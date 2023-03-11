package uzb.example.wordgenerator.service;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 11/03/2023.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExample1 {
    public static void main(String[] args) {
        try {
            // Excel faylni ochish
            FileInputStream file = new FileInputStream(new File("test.xlsx"));

            // Workbook yaratish
            Workbook workbook = new XSSFWorkbook(file);

            // Hujjatni tanlash
            Sheet sheet = workbook.getSheetAt(0);

            // Qatorni tanlash
            Row row = sheet.getRow(0);

            // Qatorni tanlash
            Cell cell = row.getCell(0);

            // Qator va qatorni chiqarish
            System.out.println(cell.getStringCellValue());

            // Ma'lumotlarni yozish
            row.createCell(1).setCellValue("Hello");
            file.close();

            // Yangi ma'lumotlarni saqlash
            FileOutputStream outFile = new FileOutputStream(new File("test.xlsx"));
            workbook.write(outFile);
            outFile.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}