package com.telus.InputInvoiceFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

public class entityIDcm {

    public static void populateEntityIDcm(Map<String, String> BANValue, String BAN, int i, String eMail) {

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
        String outputpath = output + "\\input_exp_cm_account_change_YYYYMMDD" + i + ".xlsx";

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

                Map<String, Integer> rowMap = new HashMap<>();
                // Iterate through each cell in the header row
                Iterator<Cell> cellIterator = headerRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell1 = cellIterator.next();
                    // Get the column name (header) and its index
                    String columnName = cell1.getStringCellValue();
                    int columnIndex = cell1.getColumnIndex();
                    rowMap.put(columnName, columnIndex);
                }



                while (rowIterator.hasNext()) {

                    Row row = rowIterator.next();
                    Cell entityCell = row.getCell(entityIdColumnIndex);  // Get the value in the "entity_id" column

                    //System.out.println("entityCell.getNumericCellValue() " + entityCell.getNumericCellValue());
                    if (entityCell != null && entityValue.equals(String.valueOf(entityCell.getNumericCellValue()))) {

                        // If entityValue matches, print the value in the "ban" column
                        Cell banCell = row.getCell(banColumnIndex);  // Get the value in the "ban" column
                        if (banCell != null && (int)banCell.getNumericCellValue() != Integer.parseInt(BAN)) {
                            System.out.println("Found matching entity_id, ban: " + (int)banCell.getNumericCellValue());
                            System.out.println("BAN "+ BAN);

                            int rowNum = row.getRowNum();
                            //System.out.println("rowNum in entityIDcm "+ rowNum);
                            Row dataRow = inputSheet.getRow(rowNum);

                            //System.out.println("dataRow in entityIDcm " + dataRow);

                            Map<String, String> BANdata = new HashMap<>();
                            for (Map.Entry<String, Integer> entry : rowMap.entrySet()) {
                                String columnName = entry.getKey();
                                int columnIndex = entry.getValue();

                                // Get the cell value for the specific column and row
                                Cell cell1 = dataRow.getCell(columnIndex);

                                // Get the cell value (as String) and store it in the map
                                String cellValue = cell1 != null ? cell1.toString() : ""; // Handle nulls
                                BANdata.put(columnName, cellValue);
                            }

                            System.out.println("BANdata Entity ID CM  "+ BANValue);

                            String STATUS_DATE = BANdata.get("creation_date");
                            String Address_1= BANdata.get("address_1");
                            String Address_2= BANdata.get("address_2");
                            String Address_3= BANdata.get("address_3");
                            String Address_4= BANdata.get("address_4");
                            String Address_5= BANdata.get("address_5");
                            String Address_6= BANdata.get("address_6");
                            String Address_7= BANdata.get("address_7");
                            String Address_8= BANdata.get("address_8");
                            String Address_9= BANdata.get("address_9");
                            String Address_10= BANdata.get("address_10");
                            String Address_11= BANdata.get("address_11");
                            String Address_12= BANdata.get("address_12");
                            String Address_13= BANdata.get("address_13");
                            String Address_14= BANdata.get("address_14");
                            String Address_15= BANdata.get("address_15");

                            Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows());

                            //Populate the data in the row
                            row1.createCell(0).setCellValue(STATUS_DATE);            //STATUS_DATE
                            row1.createCell(1).setCellValue(STATUS_DATE);            //Open Date
                            row1.createCell(4).setCellValue(0);                 //amount
                            row1.createCell(5).setCellValue(banCell.getNumericCellValue());                    //BAN
                            row1.createCell(6).setCellValue(1001);                   //DATA_SOURCE_ID
                            row1.createCell(7).setCellValue(BANdata.get("ban_status"));     //BAN_STATUS
                            row1.createCell(8).setCellValue("A");                    //CUSTOMER_TYPE. Default A.
                            row1.createCell(9).setCellValue("B");                    //CUST_SUB_TYPE. Default B.
                            row1.createCell(11).setCellValue("CA");                  //PAYMENT_METHOD. Default CA.


                            String rcidString = (String) BANdata.get("rcid");       //L9_rcid. Remove ".0" and Convert to int
                            if (rcidString.contains(".0")) {
                                rcidString = rcidString.replace(".0", "");
                            }
                            try {
                                int rcid = Integer.parseInt(rcidString); // Parse string to int
                                //System.out.println("RCID: " + rcid);
                                row1.createCell(12).setCellValue(rcid);
                            } catch (NumberFormatException e) {
                                e.printStackTrace();  // Handle the case where the string is not a valid integer
                            }
                            //row2.createCell(12).setCellValue(BANValue.get("rcid"));  //L9_RCID


                            String cpidString = (String) BANdata.get("cbucid");   //L9_CPID. Remove ".0" and Convert to int
                            if (cpidString.contains(".0")) {
                                cpidString = cpidString.replace(".0", "");
                            }
                            try {
                                int cbucid = Integer.parseInt(cpidString); // Parse string to int
                                //System.out.println("CPID: " + cbucid);
                                row1.createCell(13).setCellValue(cbucid);
                            } catch (NumberFormatException e) {
                                e.printStackTrace();  // Handle the case where the string is not a valid integer
                            }
                            //row2.createCell(13).setCellValue(BANValue.get("cbucid")); //L9_CPID

                            row1.createCell(14).setCellValue("EN");                   //L9_PREFERRED_LANGUAGE
                            row1.createCell(15).setCellValue(BANdata.get("legal_name_line_1")); //LEGAL_NAME_LINE1
                            row1.createCell(16).setCellValue(BANdata.get("legal_name_line_2")); //LEGAL_NAME_LINE2
                            row1.createCell(17).setCellValue(eMail);                  //EMail
                            row1.createCell(18).setCellValue(Address_1);              //address
                            row1.createCell(19).setCellValue(Address_2);
                            row1.createCell(20).setCellValue(Address_3);
                            row1.createCell(21).setCellValue(Address_4);
                            row1.createCell(22).setCellValue(Address_5);
                            row1.createCell(23).setCellValue(Address_6);
                            row1.createCell(24).setCellValue(Address_7);
                            row1.createCell(25).setCellValue(Address_8);
                            row1.createCell(26).setCellValue(Address_9);
                            row1.createCell(27).setCellValue(Address_10);
                            row1.createCell(28).setCellValue(Address_11);
                            row1.createCell(29).setCellValue(Address_12);
                            row1.createCell(30).setCellValue(Address_13);
                            row1.createCell(31).setCellValue(Address_14);
                            row1.createCell(32).setCellValue(Address_15);


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
