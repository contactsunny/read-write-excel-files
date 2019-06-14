package com.contactsunny.poc.ReadWriteExcelFilePoc;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class App {

    private static final Logger logger = Logger.getLogger(App.class);

    public static void main(String[] args) {

        BasicConfigurator.configure();

        String filePath = args[0];
        String sheetName = args[1];
        String data = args[2];

        logger.info("filePath: " + filePath);
        logger.info("sheetName: " + sheetName);
        logger.info("data: " + data);

        writeExcel(filePath, sheetName, data);
    }

    private static void writeExcel(String filePath, String sheetName, String data) {

        File file = new File(filePath);

        if (!file.exists()) {
            CreateFile(filePath, sheetName, data);
        } else {
            updateFile(file, sheetName, data);
        }

    }

    private static void CreateFile(String filePath, String sheetName, String data) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);

        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);

        cell.setCellValue(data);

        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));

            workbook.write(out);
            out.close();
            logger.info("Excel written successfully..");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void updateFile(File file, String sheetName, String data) {

        try {
            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);
            Row row;

            int lastRow = sheet.getLastRowNum();
            logger.info("lastRow: " + lastRow);

            row = sheet.getRow(lastRow);

            logger.info("Last Cell Value: " + row.getCell(0));

            int newRowNumber = lastRow + 1;
            logger.info("newRowNumber: " + newRowNumber);

            Row newRow = sheet.createRow(newRowNumber);
            newRow.createCell(0).setCellValue(data);

            inputStream.close();

            try {
                FileOutputStream out = new FileOutputStream(file, false);

                workbook.write(out);

                out.close();

            } catch (IOException e) {
                e.printStackTrace();
            }

            logger.info("Excel written successfully..");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
