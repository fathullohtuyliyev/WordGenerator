package uzb.example.wordgenerator.controller;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * @project WordGenerator
 * @author: Fathullo To'yliyev on 09/03/2023.
 */
@Controller
// 2-usul
public class WordGenerationController {
    @GetMapping("/generate-word")
    public ResponseEntity<byte[]> generateWordDocument() throws IOException {

        // Yangi word hujjatini yaratish
        XWPFDocument document = new XWPFDocument();

        // Yangi paragraf yaratish va ishga tushirish
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        // Matn qo'shish
        run.setText("ByteArrayOutputStream - bu Java tilidagi klass bo'lib, u xotiradagi bayt massiviga ma'lumotlarni yozish usulini ta'minlaydi. Bu diskdagi fayl o'rniga xotiradagi buferga ma'lumotlarni" +
                " yozish imkonini beradi. ByteArrayOutputStream-ga yozilgan ma'lumotlar bayt massivida saqlanadi.");

        // Hujjatni ByteArrayOutputStream-ga saqlash
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        document.write(baos);

        // ByteArrayOutputStreamni bayt massiviga aylantiring
        byte[] bytes = baos.toByteArray();

        // Yangi HttpHeaders ob'ektini yarating va kontent turi va joylashuvini o'rnatish
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "document.docx");

        // Bayt massivi va HttpHeaders bilan ResponseEntity ni qaytaring
        return ResponseEntity.ok()
                .headers(headers)
                .body(bytes);
    }
}
