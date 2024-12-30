package zeev.fraiman.workwithexcel;

import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {

    Context context;
    ListView lv;
    ArrayList<User> all;
    ArrayAdapter<String> adapter;
    TextView tv;
    Button bSave, bFindFile, bView;
    private Uri fileUri;
    private ActivityResultLauncher<Intent> filePickerLauncher;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        //startActivity(new Intent(MainActivity.this, FromCloudeFile.class));
        ActivityCompat.requestPermissions(this,
                new String[]{android.Manifest.permission.READ_EXTERNAL_STORAGE},1);

        initComponents();

        bFindFile.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                openFilePicker();
            }
        });

        bSave.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                writeToExcel(fileUri,"Ququ", 2000);
            }
        });

        bView.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                viewFromFile(fileUri);
            }
        });
    }

    private void initComponents() {
        context=this;
        bFindFile=findViewById(R.id.bFindFile);
        bView=findViewById(R.id.bView);
        bSave=findViewById(R.id.bSave);
        lv=findViewById(R.id.lv);
        all=new ArrayList<>();
        tv=findViewById(R.id.tv);

        filePickerLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == RESULT_OK && result.getData() != null) {
                        fileUri = result.getData().getData();
                        //tv.setText(""+fileUri.toString());
                        //String path = getPathFromUri(fileUri);
                        //processExcelFile(path);
                    }
                }
        );
    }

    private void viewFromFile(Uri fileUri) {
        try {
            InputStream is = getContentResolver().openInputStream(fileUri);
            Workbook workbook = new XSSFWorkbook(is);
            Sheet sheet=workbook.getSheetAt(0);
            for (Row row:sheet)  {
                if (row.getRowNum()==0)  {
                    continue;
                }
                String userName=row.getCell(0).getStringCellValue();
                int bdYear= (int) row.getCell(1).getNumericCellValue();
                User user=new User(userName,""+bdYear);
                all.add(user);
            }
            String[] allUsers=new String[all.size()];
            for (int i = 0; i < allUsers.length; i++) {
                allUsers[i]=all.get(i).toString();
            }
            adapter=new ArrayAdapter<>(MainActivity.this,
                    android.R.layout.simple_list_item_1,
                    allUsers);
            lv.setAdapter(adapter);
            is.close();
        } catch (FileNotFoundException e) {
            tv.setText(e.toString());
        } catch (IOException e) {
            tv.setText(e.toString());
        }
    }

    private void writeToExcel(Uri fileUri, String userName, int bdYear) {
        try {
            //File file = new File(getFilesDir(), "Users.xlsx");
            Workbook workbook;
            Sheet sheet;
            if (fileUri!=null) {
                // גישה לקובץ אם קיים
                //FileInputStream fis = new FileInputStream(file);
                InputStream is = getContentResolver().openInputStream(fileUri);
                workbook = new XSSFWorkbook(is);
                sheet = workbook.getSheetAt(0);
                is.close();
            } else {
                // אם קובץ לא קיים אז יוצרים חדש ריק
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Users");
                // כותרות לשדות
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("User Name");
                headerRow.createCell(1).setCellValue("Birth Year");
            }

            // מגיעים לשורה האחרונה
            int lastRowNum = sheet.getLastRowNum();
            Row newRow = sheet.createRow(lastRowNum + 1);

            // רושמים נתונים בשורה האחרונה
            newRow.createCell(0).setCellValue(userName);
            newRow.createCell(1).setCellValue(bdYear);

            // סוגרים גישה עם שמירת שינויים בקובץ
            //FileOutputStream fos = new FileOutputStream(file);
            OutputStream os = getContentResolver().openOutputStream(fileUri);
            workbook.write(os);
            os.close();
            workbook.close();

            Toast.makeText(this, "Data written successfully", Toast.LENGTH_SHORT).show();
        } // end of try
        catch (IOException e) {
            e.printStackTrace();
            Toast.makeText(this, "Error writing data", Toast.LENGTH_SHORT).show();
        }
    }

    private void openFilePicker() {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        //startActivityForResult(intent, PICK_EXCEL_FILE);
        filePickerLauncher.launch(intent);
    }

    private String getPathFromUri(Uri uri) {
        String filePath = null;
        String[] projection = {MediaStore.Files.FileColumns.DATA};
        Cursor cursor = getContentResolver().query(uri, projection, null, null, null);
        if (cursor != null) {
            if (cursor.moveToFirst()) {
                int columnIndex = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DATA);
                filePath = cursor.getString(columnIndex);
            }
            cursor.close();
        }
        else {
            filePath="Noting";
        }
        tv.setText("result="+filePath);
        return filePath;
    }

}