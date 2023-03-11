package uzb.example.wordgenerator.service;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 11/03/2023.
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
//Excel-da xaridorlar to'plamiga qarab hisob kitoblarni yaratish
public class ExcelFile {
    public static void main(String[] args) {
        // Excel faylini yaratish
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Yangi Excel jadvali yaratish
        XSSFSheet sheet = workbook.createSheet("Xaridorlar");

        // Ma'lumotlarni kiritish
        Object[][] xaridorlar = {
                {"Xaridor", "Mahsulot", "Narx", "To'lov turi"},
                {"Ali", "Telefon", "1000", "Naqd"},
                {"Vali", "Kompyuter", "1500", "Plastik"},
                {"Soli", "Televizor", "1200", "Bank karta"},
        };

        int rowNum = 0;
        System.out.println("Excel fayliga yozish boshlandi...");
        for (Object[] xaridor : xaridorlar) {
            // Yangi qator yaratish
            Row row = sheet.createRow(rowNum++);
            //Ma'lumotlarni jadvalga kiritish:
            int colNum = 0;
            for (Object field : xaridor) {
                // Qatorga ma'lumotlarni kiritish
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try {
            // Excel faylini saqlash
            FileOutputStream outputStream = new FileOutputStream("Xaridorlar.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel fayliga yozish yakunlandi!");
    }
}

