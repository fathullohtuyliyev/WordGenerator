package uzb.example.wordgenerator.service;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 11/03/2023.
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

//Excel-da funktsiyalar va hisob kitoblari yaratish
public class ExcelExample3 {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("My Sheet");
        Row row1 = sheet.createRow(0);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue(10);
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(20);
        Row row3 = sheet.createRow(2);
        Cell cell3 = row3.createCell(0);
        cell3.setCellValue(30);
        Row row4 = sheet.createRow(3);
        Cell cell4 = row4.createCell(0);
        cell4.setCellFormula("SUM(A1:A3)");

        // Formulani baholang
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateFormulaCell(cell4);

        // Write the output to a file-Chiqishni faylga yozing
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // workbookni yoping
        workbook.close();
    }
}


