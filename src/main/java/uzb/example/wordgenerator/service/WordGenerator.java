package uzb.example.wordgenerator.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 09/03/2023.
 */
//public class WordGenerator {
//    public void generateDocument() throws IOException {
//        // Word hujjati shablonini yuklang
//        XWPFDocument document = new XWPFDocument(getClass().getResourceAsStream("/templates/document_template.docx"));
//
//        // To'ldiruvchilarni dinamik tarkib bilan almashtiring
//        for (XWPFParagraph paragraph : document.getParagraphs()) {
//            for (XWPFRun run : paragraph.getRuns()) {
//                String text = run.getText(0);
//                if (text != null && text.contains("${name}")) {
//                    text = text.replace("${name}", "John Doe");
//                    run.setText(text, 0);
//                }
//            }
//        }
//
//        // Yaratilgan Word hujjatini faylga saqlang
//        FileOutputStream outputStream = new FileOutputStream("generated_document.docx");
//        document.write(outputStream);
//        outputStream.close();
//    }


//}
