package com.telus.InputInvoiceFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

public class entityIDinvoice {

    public static void populateEntityIDinvoice(Map<String, String> BANValue, String BAN, int i) {

        // Fetch the entity_id value
        String entityID = BANValue.get("entity_id");

        // Paths to the input and output Excel files
        Properties properties1 = new Properties();
        String configFilePath = "src/main/resources/config.properties";

        try (FileInputStream input = new FileInputStream(configFilePath)) {
            properties1.load(input);
        } catch (IOException e) {
            System.out.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        // Get paths from properties file
        String excelFilePath = properties1.getProperty("inputBANFile1");
        System.out.println("excelfilepath " + properties1.getProperty("inputBANFile1"));
        String output = properties1.getProperty("outputpath");
        String outputpath = output + "\\input_invoice_YYYYMMDD" + i + ".xlsx";

        System.out.println("outputpath " + outputpath);


        try {

            String entityValue = BANValue.get("entity_id");

            // Open the input Excel file (File 1)
            FileInputStream fis2 = new FileInputStream(new File(excelFilePath));
            Workbook inputWorkbook = new XSSFWorkbook(fis2);
            Sheet inputSheet = inputWorkbook.getSheetAt(0);  // Read the first sheet (0 index)

            FileInputStream fis10 = new FileInputStream(outputpath);
            XSSFWorkbook workbook = new XSSFWorkbook(fis10);
            Sheet sheet = workbook.getSheetAt(0);

            // Find the column index of "entity_id"
            int entityIdColumnIndex = -1;
            int banColumnIndex = -1;

            Row headerRow = inputSheet.getRow(0);  // The first row is assumed to be the header row
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if ("entity_id".equals(cell.getStringCellValue())) {
                        entityIdColumnIndex = cell.getColumnIndex();
                    }
                    if ("ban".equals(cell.getStringCellValue())) {
                        banColumnIndex = cell.getColumnIndex();
                    }
                }

                // Search for the entityvalue starting from the second row
                boolean found = false;
                Iterator<Row> rowIterator = inputSheet.iterator();
                rowIterator.next();  // Skip the header row

                int counter = 2;

                System.out.println("entityValue " + entityValue);
                while (rowIterator.hasNext()) {

                    Row row = rowIterator.next();
                    Cell entityCell = row.getCell(entityIdColumnIndex);  // Get the value in the "entity_id" column

                    //System.out.println("entityCell.getNumericCellValue() " + entityCell.getNumericCellValue());
                    if (entityCell != null && entityValue.equals(String.valueOf(entityCell.getNumericCellValue()))) {

                        // If entityValue matches, print the value in the "ban" column
                        Cell banCell = row.getCell(banColumnIndex);  // Get the value in the "ban" column
                        if (banCell != null && banCell.getNumericCellValue() != Integer.parseInt(BAN)) {
                            //System.out.println("Found matching entity_id, ban: " + banCell.getNumericCellValue());

                            Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows());
                            Row rowA2 = sheet.getRow(1); // Row 1 corresponds to A2
                            for (int j = 0; j < 5; j++) {
                                Cell cellA2 = rowA2.getCell(j);
                                if (cellA2 == null) {
                                    System.out.println(" cellA2 is null ");
                                }
                                if (cellA2 != null) {
                                    System.out.println("cellA2.getCellType() "+cellA2.getCellType());
                                    switch (cellA2.getCellType()) {
                                        case NUMERIC:
                                            int cellValue = (int) cellA2.getNumericCellValue();
                                            row1.createCell(j).setCellValue(cellValue);
                                            break;
                                        case STRING:
                                            String cellValueS = cellA2.getStringCellValue();
                                            row1.createCell(j).setCellValue(cellValueS);
                                    }

                                }
                                row1.createCell(5).setCellValue(0);
                                row1.createCell(6).setCellValue(0);
                                row1.createCell(7).setCellValue(banCell.getNumericCellValue());
                                row1.createCell(8).setCellValue(1001);

                                //System.out.println("banCell.getNumericCellValue() " + banCell.getNumericCellValue());
                            }
                            found = true;
                        }

                    }

                    FileOutputStream fileOut = new FileOutputStream(new File(outputpath));
                    workbook.write(fileOut);

                }
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
