package uzb.example.wordgenerator.service;

import java.io.*;

import org.apache.poi.xslf.usermodel.*;

public class PowerPointExample {
    public static void main(String[] args) {
        // PowerPoint presentatsiyasini yaratish
        XMLSlideShow ppt = new XMLSlideShow();

        // Yangi slide yaratish
        XSLFSlide slide = ppt.createSlide();

        // Yangi paragraf yaratish va matn yozish
        XSLFTextShape title = slide.createTextBox();
        title.setText("Assalomu alaykum, dunyo!");

        // PowerPoint presentatsiyasini saqlash
        try (FileOutputStream out = new FileOutputStream("presentation.pptx")) {
            ppt.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // PowerPoint presentatsiyasini o'qish
        try (FileInputStream fileInputStream = new FileInputStream("presentation.pptx")) {
            XMLSlideShow readPpt = new XMLSlideShow(fileInputStream);
            XSLFSlide readSlide = readPpt.getSlides().get(0);
            XSLFTextShape readTitle = (XSLFTextShape) readSlide.getShapes().get(0);
            String text = readTitle.getText();
            System.out.println(text);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // PowerPoint presentatsiyasiga yangi slide qo'shish
        XSLFSlide newSlide = ppt.createSlide();
        XSLFTextShape newTitle = newSlide.createTextBox();
        newTitle.setText("Bugungi kunda havo yaxshi!");

        // PowerPoint presentatsiyasini saqlash
        try (FileOutputStream out = new FileOutputStream("presentation.pptx")) {
            ppt.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
