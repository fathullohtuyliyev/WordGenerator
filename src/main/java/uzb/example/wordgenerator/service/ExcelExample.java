package uzb.example.wordgenerator.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 11/03/2023.
 */
public class ExcelExample {
    public static void main(String[] args) throws IOException {
        // Yangi Workbook yaratish
        Workbook workbook = new XSSFWorkbook();

        // Hujjat yaratish
        Sheet sheet = workbook.createSheet("MySheet");

        // Qator va ustunlarni yaratish
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("FileInputStream Java dasturlash tilida fayllarni o'qish uchun foydalaniladigan klassdir.");

        // Hujjatni saqlash
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Workbookni yopish
        workbook.close();
    }
}
