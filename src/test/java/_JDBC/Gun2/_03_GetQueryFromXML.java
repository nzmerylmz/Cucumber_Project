package _JDBC.Gun2;

import _JDBC.JDBCParent;
import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class _03_GetQueryFromXML extends JDBCParent
{
    @Test
    public void test() throws SQLException {
        Cell cell = null;
        String path = "src/test/java/_JDBC/Gun2/QueryPath.xlsx";
        FileInputStream inputStream;
        Workbook workbook;
        try {
            inputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = workbook.getSheet("Query");
        int rowCount=sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rowCount; i++) {
            Row row=sheet.getRow(i);
            int cellCount=row.getPhysicalNumberOfCells();
            for (int j = 0; j < cellCount; j++) {
                cell=row.getCell(j);
            }
        }
        getTable(cell.toString());
    }
    public static void getTable(String queryPath) throws SQLException {
        ResultSet rs = statement.executeQuery(queryPath);
        ResultSetMetaData rsmd = rs.getMetaData();
        for (int i = 1; i <= rsmd.getColumnCount(); i++) {
            System.out.printf("%-20s",rsmd.getColumnName(i));
        }
        System.out.println();
        while (rs.next()){
            for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                System.out.printf("%-20s",rs.getString(i));
            }
            System.out.println();
        }
    }
}
