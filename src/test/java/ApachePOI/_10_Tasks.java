package ApachePOI;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class _10_Tasks {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Calculator");
        Map<Integer, Object[]> verticalData = new TreeMap<>();
        int verticalRowNum = 0;
        int i = 1;
        while (verticalRowNum != 100) {
            for (int j = 1; j <= 10; j++) {
                verticalData.put(j, new Object[]{i, "X", j, (i * j)});
            }
            Set<Integer> keyset = verticalData.keySet();
            for (Integer key : keyset) {
                Row row = sheet.createRow(verticalRowNum++);
                Object[] objectArray = verticalData.get(key);
                int verticalCellNum = 0;
                for (Object object : objectArray) {
                    Cell cell = row.createCell(verticalCellNum++);
                    if (object instanceof String)
                        cell.setCellValue((String) object);
                    else if (object instanceof Integer)
                        cell.setCellValue((Integer) object);
                }
            }
            i++;
        }
        Map<Integer, Object[]> horizontalData = new TreeMap<>();
        int horizontalCellnum = 4;
        int horizontalRowNum = 0;
        i = 2;
        while (i != 11) {
            for (int j = 1; j <= 10; j++) {
                horizontalData.put(j, new Object[]{i, "X", j, (i * j)});
            }
            Set<Integer> keyset = horizontalData.keySet();
            for (Integer key : keyset) {
                Row newRow = sheet.getRow(horizontalRowNum++);
                Object[] objectArray = horizontalData.get(key);
                for (Object object : objectArray) {
                    Cell cell = newRow.createCell(horizontalCellnum++);
                    if (object instanceof String)
                        cell.setCellValue((String) object);
                    else if (object instanceof Integer)
                        cell.setCellValue((Integer) object);
                }
                horizontalCellnum-=4;
            }
            horizontalCellnum+=4;
            horizontalRowNum=0;
            i++;
        }
        try {
            FileOutputStream outputStream = new FileOutputStream("src/test/java/ApachePOI/resource/NewTask.xlsx");
            workbook.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.out.println("File created");
    }
}


