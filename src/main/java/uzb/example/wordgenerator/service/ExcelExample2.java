package uzb.example.wordgenerator.service;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 11/03/2023.
 */

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
//Excel-da qator va ustunlarni ochish, qisqartirish va boshqa tuzatishlar.
public class ExcelExample2 {
    public static void main(String[] args) throws IOException {

        // Yangi Workbook yaratish
        Workbook workbook = new XSSFWorkbook();

        // Yangi Sheet-varaq yaratish
        Sheet sheet = workbook.createSheet("MySheet");

        // Qator va ustunlarni yaratish
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(2);
        Cell cell2 = row.createCell(5);
        cell1.setCellValue("Hello");
        cell2.setCellValue("World");

        // Qatorlarni qisqartirish
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(false);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        row.setRowStyle(style);

        // Ustunlarni qisqartirish
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        Cell headerCell2 = headerRow.createCell(1);
        headerCell1.setCellValue("Column 1");
        headerCell2.setCellValue("Column 2");
        headerCell1.setCellStyle(headerStyle);
        headerCell2.setCellStyle(headerStyle);

        // Hujjatni saqlash
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Workbookni yopish
        workbook.close();
    }





}


