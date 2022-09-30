package _JDBC.Gun2;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

public class _05_InsertXMLIntoMySQL {

    private static Connection connection;
    protected static Statement statement;

    @BeforeTest
    public void DBConnectionOpen() {
        String url = "jdbc:mysql://db-technostudy.ckr1jisflxpv.us-east-1.rds.amazonaws.com:3306/z_nazim";
        String username = "root";
        String password = "'\"-LhCB'.%k[4S]z";

        try {
            connection = DriverManager.getConnection(url, username, password);
            statement = connection.createStatement();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }
    @Test
    public void test() throws SQLException {
        String path = "src/test/java/_JDBC/Gun2/XMLtoSQL.xlsx";
        FileInputStream inputStream;
        Workbook workbook;
        try {
            inputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = workbook.getSheet("Sayfa1");
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            int cellCount = 1;
            String g = "INSERT INTO contactInfo(ad, soyad, tel) VALUES (?,?,?)";
            PreparedStatement ps = connection.prepareStatement(g);
            ps.setString(1, row.getCell(cellCount).toString());
            ps.setString(2, row.getCell(cellCount + 1).toString());
            ps.setString(3, row.getCell(cellCount+ 2).toString());
            ps.executeUpdate();
//                String wholeCell="";
//                String cellnum1=row.getCell(j).toString();
//                String cellnum2=row.getCell(j+1).toString();
//                String cellnum3=row.getCell(j+2).toString();
//                wholeCell="'"+cellnum1+"','"+cellnum2+"','"+cellnum3+"'";
//                statement.executeUpdate("INSERT INTO contactInfo(ad, soyad, tel) VALUES ("+wholeCell+")");
        }
    }

    @AfterTest
    public void DBConnectionClose() {
        try {
            connection.close();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }
}
