package zeev.fraiman.workwithexcel;

public class User {
    private String userName;
    private String userBDYear;

    public User(String userName, String userBDYear) {
        this.userName = userName;
        this.userBDYear = userBDYear;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getUserBDYear() {
        return userBDYear;
    }

    public void setUserBDYear(String userBDYear) {
        this.userBDYear = userBDYear;
    }

    @Override
    public String toString() {
        return "User{" +
                "userName='" + userName + '\'' +
                ", userBDYear='" + userBDYear + '\'' +
                '}';
    }
}
