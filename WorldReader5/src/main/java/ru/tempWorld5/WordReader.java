package ru.tempWorld5;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.util.List;

/**
 * Вывод содержимого файла на экран
 */
public class WordReader {
    public static void main(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream("F:/Apache POI Word Test.docx")) {
            //try (FileInputStream fileInputStream = new FileInputStream("Temp.docx")) {

            // открываем файл и считываем его содержимое в объект XWPFDocument
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);

            // считываем верхний колонтитул (херед документа)
            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
            System.out.println(docHeader.getText());

            // печатаем содержимое всех параграфов документа в консоль
            List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
            for (XWPFParagraph p : paragraphs) {
                System.out.println(p.getText());
            }
            // считываем нижний колонтитул (футер документа)
            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
            System.out.println(docFooter.getText());

            /*System.out.println("_____________________________________");
            // печатаем все содержимое Word файла
            XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
            System.out.println(extractor.getText());*/

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}