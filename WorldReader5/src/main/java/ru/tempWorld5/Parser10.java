package ru.tempWorld5;

//import org.apache.poi.ss.usermodel.*;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


@Component
public class Parser10 {

    private Utils utils;

    //Данные,сюда заливаем из Word
    Map<String,String> cellData = null;

    @Autowired
    public Parser10(Utils utils) {
        this.utils = utils;
    }

    public Map<String, String> parse() {
        try (FileInputStream fileInputStream = new FileInputStream("F:/Apache POI Word Test3.docx")) {
            //try (FileInputStream fileInputStream = new FileInputStream("Temp.docx")) {

            // открываем файл и считываем его содержимое в объект XWPFDocument
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);

            // считываем верхний колонтитул (херед документа)
            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
            //System.out.println(docHeader.getText());

            String word = null;
            //String[] words2;
            List<String> cells = new ArrayList<>();
            List<String> data = new ArrayList<>();
            //Map<List<String>,List<String>> cellsData = new HashMap<>();
            Map<String, String> cellsData = null;
            //String[] concatArray(String[] a, String[] b)
            //int cellsSize = a.size();
            int cellsSize = 0;
            //int dataSize = b.size();
            int dataSize = 0;
            //ArrayList<String> result = new ArrayList(cellsSize + dataSize);
            ArrayList<String> result = null;
            //Map<String,String> cellData = null;
            // печатаем содержимое всех параграфов документа в консоль
            List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
            for (XWPFParagraph p : paragraphs) {

                //Просто вывод строк документа
                //System.out.println(p.getText());

                //Заполняем массив с адресами ячеек
                Pattern pattern = Pattern.compile("([A-Z]+[0-9])");
                Matcher matcher = pattern.matcher((p.getText()));
                while(matcher.find()) {
                    //System.out.println(matcher.group());
                    cells.add(matcher.group());
                }
                cellsSize = cells.size();
                //Заполняем массив с данными
                Pattern pattern2 = Pattern.compile("([<]+[A-Z]+[0-9]+[>])");
                //Matcher matcher = pattern2.matcher((p.getText()));
                String[] words2= pattern2.split(p.getText());
                //Вытаскиваем набережная реки Карповки д.12 и 150,23
                for(String wordT:words2) {
                    //System.out.println(word);
                    for (String retval : wordT.split(">")) {
                        //System.out.println(retval);
                        for (String retval2 : retval.split("<")) {
                            //System.out.println(retval2);
                            if(!retval2.isEmpty()){
                                data.add(retval2);
                            }

                        }
                    }
                }
                dataSize = data.size();
                result = new ArrayList(cellsSize + dataSize);
                cellData = new HashMap<>();
            }
            //Смотрим содержание массива с адресами ячеек
            /*for(String wordT:cells) {
                System.out.println(wordT);
            }*/
            //Смотрим содержание массива с данными
            /*for(String wordT:data) {
                System.out.println(wordT);
            }*/
            //Соединяем два массива - var 1
            /*if ((cells instanceof RandomAccess) && (data instanceof RandomAccess)) {
                for (int i = 0; i < cellsSize; i++) result.add(cells.get(i));
                for (int i = 0; i < dataSize; i++) result.add(data.get(i));
            }*/
            //Соединяем два массива - var 2
            for (int i = 0; i < cellsSize; i++) {
                //result.add(cells.get(i));
                //result.add(data.get(i));
                cellData.put(cells.get(i),data.get(i));
            }
            //Смотрим результирующий массив
            /*for(String temp:result) {
                System.out.println(temp);
            }*/
            /*for (Map.Entry entry : cellData.entrySet()) {
                System.out.println("Key: " + entry.getKey() + " Value: "
                        + entry.getValue());
            }*/

            // считываем нижний колонтитул (футер документа)
            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
            //System.out.println(docFooter.getText());

            /*System.out.println("_____________________________________");
            // печатаем все содержимое Word файла
            XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
            System.out.println(extractor.getText());*/

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        return cellData;
    }

    public void parseOutExcel(Map<String,String> cellData){
        for (Map.Entry entry : cellData.entrySet()) {
            System.out.println("Key: " + entry.getKey() + " Value: "
                    + entry.getValue());
        }
    }

    public void parseOutExcel2(Map<String,String> cellData){
        //сюда заливаем из cellData
        String outputFile = utils.getOutputFile();
        Workbook outputWorkbook = utils.getWorkBookFromFile(outputFile); // output книга
        Sheet outputFileSheet = outputWorkbook.getSheetAt(0); //лист, сюда будем писать
        // get the last row number to append new data

       /* int rownum = outputFileSheet.getLastRowNum();
        CellReference cellReference;
        Cell cell;
        Iterator<String> keySetIterator = cellData.keySet().iterator();
        while(keySetIterator.hasNext()){
            cell = null;
            String key = keySetIterator.next();
            System.out.println("key: " + key + " value: " + cellData.get(key));
            Row row = outputFileSheet.createRow(rownum);
            cellReference = new CellReference(key);
            System.out.println("Номер столбца: " + cellReference.getCol() + "Номер строки: " + cellReference.getRow());
            //Cell cell = row.createCell(cellReference.getCol());
            cell = row.createCell(cellReference.getCol());
            cell.setCellValue(String.valueOf(cellData.get(key)));
            utils.saveWorkbook(outputWorkbook, outputFile);
       }*/

        //Обход коллекции и запись в Excel с использованием entrySet()
        for (Map.Entry entry : cellData.entrySet()){
            //System.out.println(entry.getKey());
            CellReference cellReference2 = new CellReference((String) entry.getKey());
            System.out.println("Номер столбца: " + cellReference2.getCol() + "Номер строки: " + cellReference2.getRow());
            Row row = outputFileSheet.createRow(cellReference2.getRow());
            Cell cell2 = row.createCell(cellReference2.getCol());
            cell2.setCellValue(String.valueOf(entry.getValue()));
        }
        utils.saveWorkbook(outputWorkbook, outputFile);
    }
}
