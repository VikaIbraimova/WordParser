package ru.tempWorld5;

//import org.apache.poi.ss.usermodel.*;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import java.util.List;


@Component
public class Parser {

    private Utils utils;

    @Autowired
    public Parser(Utils utils) {
        this.utils = utils;
    }

    public void parse() {
        List<String> importFiles = utils.getImportedFilesList();
        String outputFile = utils.getOutputFile();
        try {
            for (String filePathString : importFiles){
                XWPFDocument importDocument = utils.getDocumentFromFile(filePathString); // получаем документ
                XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(importDocument);

                // считываем верхний колонтитул (херед документа)
                XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
                System.out.println(docHeader.getText());

                // печатаем содержимое всех параграфов документа в консоль
                List<XWPFParagraph> paragraphs = importDocument.getParagraphs();
                for (XWPFParagraph p : paragraphs) {
                    System.out.println(p.getText());
                }
                // считываем нижний колонтитул (футер документа)
                XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
                System.out.println(docFooter.getText());

            }
        }
        catch (Exception ex) {
            ex.printStackTrace();
        }

    }
}
