package com.telus.Properties;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class createInputFiles {
	
	//This method checks the number of sheets in the BAN.xlxs file. 
	//The number of sheets should be 1 else the program exists
	public static void  createSheets(int i,String outputpath) {
		
		 // Create a new workbook (Excel file)
        

        //*****************Create file-input_aged_status************************
        // Create a sheet in the workbook
		Workbook workbook = new XSSFWorkbook();
		
		String fileName1="input_aged_status_YYYYMMDD" + i + ".xlsx";
        Sheet sheet = workbook.createSheet("input_aged_status_YYYYMMDD");

        // Create a header row (first row)
        Row headerRow = sheet.createRow(0);
        
        /*
     // Create cells in the header row
        headerRow.createCell(0).setCellValue("BAN");
        headerRow.createCell(1).setCellValue("TR-1-30");
        headerRow.createCell(2).setCellValue("TR-31-60");
        headerRow.createCell(3).setCellValue("TR-61-90");
        headerRow.createCell(4).setCellValue("TR-91-120");
        headerRow.createCell(5).setCellValue("TR-121-150");
        headerRow.createCell(6).setCellValue("TR-151-180");
        headerRow.createCell(7).setCellValue("TR-180+");
        headerRow.createCell(8).setCellValue("TR0");
        headerRow.createCell(9).setCellValue("NT-1-30");
        headerRow.createCell(10).setCellValue("NT-31-60");
        headerRow.createCell(11).setCellValue("NT-61-90");
        headerRow.createCell(12).setCellValue("NT-91-120");
        headerRow.createCell(013).setCellValue("NT-121-150");
        headerRow.createCell(014).setCellValue("NT-151-180");
        headerRow.createCell(015).setCellValue("NT-180+");
        headerRow.createCell(016).setCellValue("NT0");
        headerRow.createCell(017).setCellValue("AD-1-30");
        headerRow.createCell(18).setCellValue("AD-31-60");
        headerRow.createCell(19).setCellValue("AD-61-90");
        headerRow.createCell(20).setCellValue("AD-91-120");
        headerRow.createCell(21).setCellValue("AD-121-150");
        headerRow.createCell(022).setCellValue("AD-151-180");
        headerRow.createCell(023).setCellValue("AD-180+");
        headerRow.createCell(024).setCellValue("AD0");
        headerRow.createCell(025).setCellValue("NineH-1-30");
        headerRow.createCell(026).setCellValue("NineH-31-60");
        headerRow.createCell(027).setCellValue("NineH-61-90");
        headerRow.createCell(28).setCellValue("NineH-91-120");
        headerRow.createCell(29).setCellValue("NineH-121-150");
        headerRow.createCell(30).setCellValue("NineH-151-180");
        headerRow.createCell(31).setCellValue("NineH-180+");
        headerRow.createCell(32).setCellValue("NineH0");
        headerRow.createCell(33).setCellValue("EQ-1-30");
        headerRow.createCell(34).setCellValue("EQ-31-60");
        headerRow.createCell(35).setCellValue("EQ-61-90");
        headerRow.createCell(36).setCellValue("EQ-91-120");
        headerRow.createCell(37).setCellValue("EQ-121-150");
        headerRow.createCell(38).setCellValue("EQ-151-180");
        headerRow.createCell(39).setCellValue("EQ-180+");
        headerRow.createCell(40).setCellValue("EQ0");
        headerRow.createCell(41).setCellValue("THIRDP-1-30");
        headerRow.createCell(42).setCellValue("THIRDP-31-60");
        headerRow.createCell(43).setCellValue("THIRDP-61-90");
        headerRow.createCell(44).setCellValue("THIRDP-91-120");
        headerRow.createCell(45).setCellValue("THIRDP-121-150");
        headerRow.createCell(46).setCellValue("THIRDP-151-180");
        headerRow.createCell(47).setCellValue("THIRDP-180+");
        headerRow.createCell(48).setCellValue("THIRDP0");
        headerRow.createCell(49).setCellValue("TOTAL-1-30");
        headerRow.createCell(50).setCellValue("TOTAL-31-60");
        headerRow.createCell(51).setCellValue("TOTAL-61-90");
        headerRow.createCell(52).setCellValue("TOTAL-91-120");
        headerRow.createCell(53).setCellValue("TOTAL-121-150");
        headerRow.createCell(54).setCellValue("TOTAL-151-180");
        headerRow.createCell(55).setCellValue("TOTAL-180+");
        headerRow.createCell(56).setCellValue("TOTAL0");
        headerRow.createCell(57).setCellValue("FLAG New or OLD");
        */
        
        try (FileOutputStream fileOut = new FileOutputStream(new File(outputpath + fileName1))) {
            workbook.write(fileOut);
            System.out.println("Excel file input_aged_status created successfully!" + i);
        }catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error in creating Excel file -  input_aged_status" + i);
            
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Error in closing Excel file - input_aged_status" + i);
            }
        }
        
       //***********************************************************************
        
      //*****************Create file-input_exp_cm_account_change_YYYYMMDD************************
        
        Workbook workbook1 = new XSSFWorkbook();
        
        String fileName2="input_exp_cm_account_change_YYYYMMDD" + i + ".xlsx";
             
        Sheet sheet1 = workbook1.createSheet("input_exp_cm_account_change_YYYYMMDD");

        // Create a header row (first row)
        Row headerRow1 = sheet1.createRow(0);
        
        /*
        // Create cells in the header row
        headerRow1.createCell(0).setCellValue("STATUS_DATE");
        headerRow1.createCell(1).setCellValue("OPEN_DATE");
        headerRow1.createCell(2).setCellValue("ACTIVITY_DATE");
        headerRow1.createCell(3).setCellValue("DEPOSIT_AMT");
        headerRow1.createCell(4).setCellValue("AR_BALANCE");
        headerRow1.createCell(5).setCellValue("BAN");
        headerRow1.createCell(6).setCellValue("DATA_SOURCE_ID");
        headerRow1.createCell(7).setCellValue("BAN_STATUS");
        headerRow1.createCell(8).setCellValue("CUSTOMER_TYPE");
        headerRow1.createCell(9).setCellValue("CUST_SUB_TYPE");
        headerRow1.createCell(10).setCellValue("L9_PREF_COMMUNICATION");
        headerRow1.createCell(11).setCellValue("PAYMENT_METHOD");
        headerRow1.createCell(12).setCellValue("L9_RCID");
        headerRow1.createCell(013).setCellValue("L9_CPID");
        headerRow1.createCell(014).setCellValue("L9_PREFERRED_LANGUAGE");
        headerRow1.createCell(015).setCellValue("LEGAL_NAME_LINE1");
        headerRow1.createCell(016).setCellValue("LEGAL_NAME_LINE2");
        headerRow1.createCell(017).setCellValue("EMAIL");
        headerRow1.createCell(18).setCellValue("Address Field 1");
        headerRow1.createCell(19).setCellValue("Address Field 2");
        headerRow1.createCell(20).setCellValue("Address Field 3");
        headerRow1.createCell(21).setCellValue("Address Field 4");
        headerRow1.createCell(022).setCellValue("Address Field 5");
        headerRow1.createCell(023).setCellValue("Address Field 6");
        headerRow1.createCell(024).setCellValue("Address Field 7");
        headerRow1.createCell(025).setCellValue("Address Field 8");
        headerRow1.createCell(026).setCellValue("Address Field 9");
        headerRow1.createCell(027).setCellValue("Address Field 10");
        headerRow1.createCell(28).setCellValue("Address Field 11");
        headerRow1.createCell(29).setCellValue("Address Field 12");
        headerRow1.createCell(30).setCellValue("Address Field 13");
        headerRow1.createCell(31).setCellValue("Address Field 14");
        headerRow1.createCell(32).setCellValue("Address Field 15");
        headerRow1.createCell(33).setCellValue("PHONE1");
        headerRow1.createCell(34).setCellValue("PHONE2");
        headerRow1.createCell(35).setCellValue("PHONE3");
        headerRow1.createCell(36).setCellValue("PHONE4");
        headerRow1.createCell(37).setCellValue("WRITE_OFF_STATUS");
        headerRow1.createCell(38).setCellValue("FLAG  New or CLOSED");
        headerRow1.createCell(39).setCellValue("RCID CHANGE FLAG");
        headerRow1.createCell(40).setCellValue("OTHER CHANGE FLAG");
        */
        
        try (FileOutputStream fileOut = new FileOutputStream(new File(outputpath+fileName2))) {
            workbook1.write(fileOut);
            System.out.println("Excel file input_exp_cm_account_change_YYYYMMDD created successfully!"+i);
        }catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error in creating Excel file -  input_exp_cm_account_change_YYYYMMDD"+i);
            
        } finally {
            try {
                workbook1.close();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Error in closing Excel file - input_exp_cm_account_change_YYYYMMDD"+i);
            }
        }
        
      //***********************************************************************
        
      //*****************Create file-input_invoice_YYYYMMDD************************
        Workbook workbook2 = new XSSFWorkbook();
        
        String fileName3="input_invoice_YYYYMMDD" + i + ".xlsx";
        
        Sheet sheet2 = workbook2.createSheet("input_invoice_YYYYMMDD");

        // Create a header row (first row)
        Row headerRow2 = sheet2.createRow(0);
        
        
        // Create cells in the header row
        headerRow2.createCell(0).setCellValue("BILLING CYCLE START SATE");
        headerRow2.createCell(1).setCellValue("BILLING CYCLE END SATE");
        headerRow2.createCell(2).setCellValue("INSTANCE");
        headerRow2.createCell(3).setCellValue("BILLING CYCLE YEAR");
        headerRow2.createCell(4).setCellValue("CYCLE CODE");
        headerRow2.createCell(5).setCellValue("TOTAL INVOICE AMT");
        headerRow2.createCell(6).setCellValue("TOTAL TAX AMT");
        headerRow2.createCell(7).setCellValue("BAN");
        headerRow2.createCell(8).setCellValue("DATA SOURCE ID");
        
        
        try (FileOutputStream fileOut = new FileOutputStream(new File(outputpath+fileName3))) {
            workbook2.write(fileOut);
            System.out.println("Excel file input_invoice_YYYYMMDD created successfully!"+i);
        }catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error in creating Excel file -  input_invoice_YYYYMMDD"+i);
            
        } finally {
            try {
                workbook2.close();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Error in closing Excel file - input_invoice_YYYYMMDD"+i);
            }
        }
        
      //***********************************************************************
       
    }
		
	
	
}