package ApachePOI;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

/**
 * Bir önceki yapılan soruda, Kullanıcıya 1.sütundaki tüm değer bir liste halinde kullanıcıya
 * yanlarında bi numara olarak sunalım. Kullanıcı hangi numararyı girerese o satırın
 * yanındaki tüm bilgiler gösterilsin.
 * Username için 1
 * Password için 2
 * ....
 * ...
 * giririniz = 2
 */
public class _06_Soru {
    public static void main(String[] args) {
        String path = "src/test/java/ApachePOI/resource/QueryPath.xlsx";
        Workbook workbook;

        try {
            FileInputStream inputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = workbook.getSheetAt(0);
        int satirSayisi = sheet.getPhysicalNumberOfRows();
        Row row;
        for (int i = 0; i < satirSayisi; i++) {
            row = sheet.getRow(i);
            Cell cell = row.getCell(0);
            System.out.println((i+1) + ". " + cell.toString());
        }
        Scanner sc = new Scanner(System.in);
        System.out.print("Your choice: ");
        int a = sc.nextInt();
        Row rowYeni = sheet.getRow(a-1);
        String sonuc = "";
        for (int i = 0; i < rowYeni.getPhysicalNumberOfCells(); i++) {
            if (rowYeni.getCell(i + 1) != null) {
                sonuc += rowYeni.getCell(i + 1) + " ";
            }
        }
        System.out.println(sonuc);
    }
}
