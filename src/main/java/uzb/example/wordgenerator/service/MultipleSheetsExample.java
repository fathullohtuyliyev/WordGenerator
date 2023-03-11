package uzb.example.wordgenerator.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
//Excel-da ko'p jadvallarni biriktirish.
public class MultipleSheetsExample {
    public static void main(String[] args) throws IOException {
        Workbook workbook = WorkbookFactory.create(true);
        Sheet sheet1 = workbook.createSheet("Sheet1");
        Sheet sheet2 = workbook.createSheet("Sheet2");

        // Sheet1-da ma'lumotlarni kiritish
        Row row1 = sheet1.createRow(0);
        Cell cellA1 = row1.createCell(0);
        cellA1.setCellValue("Sheet1dagi ma'lumotlar");
        Cell cellB1 = row1.createCell(1);
        cellB1.setCellValue("Sana");
        Cell cellC1 = row1.createCell(2);
        cellC1.setCellValue("Narx");

        Row row2 = sheet1.createRow(1);
        Cell cellA2 = row2.createCell(0);
        cellA2.setCellValue("1");
        Cell cellB2 = row2.createCell(1);
        cellB2.setCellValue("01.01.2022");
        Cell cellC2 = row2.createCell(2);
        cellC2.setCellValue("1000");

        Row row3 = sheet1.createRow(2);
        Cell cellA3 = row3.createCell(0);
        cellA3.setCellValue("2");
        Cell cellB3 = row3.createCell(1);
        cellB3.setCellValue("02.01.2022");
        Cell cellC3 = row3.createCell(2);
        cellC3.setCellValue("1500");

        // Sheet2-da ma'lumotlarni kiritish
        Row row4 = sheet2.createRow(0);
        Cell cellA4 = row4.createCell(0);
        cellA4.setCellValue("Sheet2dagi ma'lumotlar");
        Cell cellB4 = row4.createCell(1);
        cellB4.setCellValue("Ism");
        Cell cellC4 = row4.createCell(2);
        cellC4.setCellValue("Yoshi");

        Row row5 = sheet2.createRow(1);
        Cell cellA5 = row5.createCell(0);
        cellA5.setCellValue("1");
        Cell cellB5 = row5.createCell(1);
        cellB5.setCellValue("Ali");
        Cell cellC5 = row5.createCell(2);
        cellC5.setCellValue("30");

        Row row6 = sheet2.createRow(2);
        Cell cellA6 = row6.createCell(0);
        cellA6.setCellValue("2");
        Cell cellB6 = row6.createCell(1);
        cellB6.setCellValue("Vali");
        Cell cellC6 = row6.createCell(2);
        cellC6.setCellValue("25");

        // Sheet1 va Sheet2 sarlavhalarini biriktirish
        CellRangeAddress mergedRegion1 = new CellRangeAddress(0, 0, 0, 2);
        sheet1.addMergedRegion(mergedRegion1);
        // Faylni saqlash
        FileOutputStream fileOut = new FileOutputStream("ko'p_jadval.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        System.out.println("Fayl muvaffaqiyatli yaratildi.");
    }

}
