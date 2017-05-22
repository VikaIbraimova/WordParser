package ru.tempWorld5;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import ru.tempWorld5.exceptions.ParseException;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

@Component
public class Utils {
    @Value(value = "${import_folder}")
    private String importDirectory;
    @Value(value = "${output_file}")
    private String outputFile;

    public List<String> getImportedFilesList() {
        List<String> filesStringsList = new ArrayList<>();
        try {
            Files.newDirectoryStream(Paths.get(importDirectory)).forEach(path -> {
                if (path.toFile().isFile()) {
                    filesStringsList.add(path.toString());
                }
            });
        } catch (IOException e) {
            //log.error("не удается открыть файл", e);
            throw new ParseException("can't open file\n", e);
        }
        return filesStringsList;
    }

    public String getOutputFile() {
        File file = Paths.get(outputFile).toFile();
        return file.toString();
    }

    /*public XWPFDocument getDocumentFromFile(String filepath) {
        try (InputStream is = new FileInputStream(filepath)) {
            //берем книгу
            return new  XWPFDocument(is);
        } catch (FileNotFoundException e) {
            //log.error("file not found: {}", filepath);
            throw new ParseException("file not found: " + filepath + "\n", e);
        } catch (IOException e) {
            //log.error("IOException for: {}", filepath);
            throw new ParseException("unexpected IO error while reading file " + filepath + "\n", e);
        }
    }*/
    public XSSFWorkbook getWorkBookFromFile(String filepath) {
        try (InputStream is = new FileInputStream(filepath)) {
//                берем книгу
            return new XSSFWorkbook(is);
        } catch (FileNotFoundException e) {
            //log.error("file not found: {}", filepath);
            throw new ParseException("file not found: " + filepath + "\n", e);
        } catch (IOException e) {
            //log.error("IOException for: {}", filepath);
            throw new ParseException("unexpected IO error while reading file " + filepath + "\n", e);
        }
    }

    public void saveWorkbook(Workbook workbook, String file) {
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            //log.trace("сохраняем файл: {}", file);
            workbook.write(outputStream);
        } catch (IOException e) {
            //log.error("не удаётся сохранить книгу {}", file);
            throw new ParseException("can't save book\n", e);
        }
    }

    public XWPFDocument getDocumentFromFile(String filepath) {
        try (InputStream is = new FileInputStream(filepath)) {
            //берем книгу
            return new  XWPFDocument(is);
        } catch (FileNotFoundException e) {
            //log.error("file not found: {}", filepath);
            throw new ParseException("file not found: " + filepath + "\n", e);
        } catch (IOException e) {
            //log.error("IOException for: {}", filepath);
            throw new ParseException("unexpected IO error while reading file " + filepath + "\n", e);
        }
    }
}
