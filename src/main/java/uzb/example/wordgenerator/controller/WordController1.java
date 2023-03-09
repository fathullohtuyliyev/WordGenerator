//package uzb.example.wordgenerator.controller;
//
//import org.apache.poi.xwpf.usermodel.XWPFDocument;
//import org.apache.poi.xwpf.usermodel.XWPFParagraph;
//import org.apache.poi.xwpf.usermodel.XWPFRun;
//import org.springframework.stereotype.Controller;
//import org.springframework.web.bind.annotation.GetMapping;
//import org.springframework.web.bind.annotation.ResponseBody;
//
//import java.io.File;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
///**
// * @project WordGenerator
// * @author: Fathullo To'yliyev on 09/03/2023.
// */
//@Controller
//public class WordController1 {
//    @GetMapping("/generate-word")
//    @ResponseBody
//    public String generateWord() throws IOException {
//        // Yangi Word hujjatini yaratish
//        XWPFDocument document = new XWPFDocument();
//
//        // Yangi paragraf yaratish
//        XWPFParagraph paragraph = document.createParagraph();
//
//        //Paragrafga matn qo'shish
//        XWPFRun run = paragraph.createRun();
//        run.setText("Apache POI Java kutubxonasi boʻlib, Word, Excel " +
//                "va PowerPoint kabi Microsoft Office hujjatlarini yaratish, oʻqish va oʻzgartirish imkonini beradi");
//
//        // Hujjatni diskda saqlash
//        FileOutputStream out = new FileOutputStream(new File("document.docx"));
//        document.write(out);
//        out.close();
//        return "Word hujjati yaratildi va diskda saqlanadi.";
//    }
//}
