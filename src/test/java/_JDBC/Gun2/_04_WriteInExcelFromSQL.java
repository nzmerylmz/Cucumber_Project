package _JDBC.Gun2;

import _JDBC.JDBCParent;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class _04_WriteInExcelFromSQL extends JDBCParent {
    @Test
    public void test() throws SQLException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("SQLQuery");
        ResultSet rs = statement.executeQuery("select * from actor");
        ResultSetMetaData rsmd = rs.getMetaData();
        Row newRow = sheet.createRow(0);
        for (int i = 1, j = 0; i <= rsmd.getColumnCount(); i++) {
            newRow.createCell(j++).setCellValue(rsmd.getColumnName(i));
        }
        int rowNumber = 1;
        while (rs.next()) {
            Row newRow1 = sheet.createRow(rowNumber);
            for (int k = 0; k < rsmd.getColumnCount(); k++) {
                newRow1.createCell(k).setCellValue(rs.getString(k + 1));
            }
            rowNumber++;
        }
        try {
            FileOutputStream outputStream = new FileOutputStream("src/test/java/_JDBC/Gun2/WriteSQLDataInExcel.xlsx");
            workbook.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.out.println("File created");
    }
}
