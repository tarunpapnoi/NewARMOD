package com.telus.InputInvoiceFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;

import com.telus.Properties.GetHeaderColumnNames;
import java.util.ArrayList;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class createCmAccountFIle {
	
	public static void  populateCMAccountFile(String outputpath, int rowNum, Map<String, String> BANValue, String BAN, int amount, String eMail,
			List<String> testResults, int i) throws FileNotFoundException {
		
		//Input invoice file
		
		String filePath = outputpath+"\\input_exp_cm_account_change_YYYYMMDD"+rowNum+".xlsx";
		//get index of columns from BAN doc
		FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        
        try {
            // Open the Excel file
            File file = new File(filePath);
            fis = new FileInputStream(file);
            
            // Create workbook instance from the file
            workbook = new XSSFWorkbook(fis);
            
            Sheet sheet = workbook.getSheetAt(0);  // Access the first sheet (0 index)
            
            Row row = sheet.createRow(0);  // Create row at index 0 with column names
            row.createCell(0).setCellValue("STATUS_DATE");
            row.createCell(1).setCellValue("OPEN_DATE");
            row.createCell(2).setCellValue("ACTIVITY_DATE");
            row.createCell(3).setCellValue("DEPOSIT_AMT");
            row.createCell(4).setCellValue("AR_BALANCE");
            row.createCell(5).setCellValue("BAN");
            row.createCell(6).setCellValue("DATA_SOURCE_ID");
            row.createCell(7).setCellValue("BAN_STATUS");
            row.createCell(8).setCellValue("CUSTOMER_TYPE");
            row.createCell(9).setCellValue("CUST_SUB_TYPE");
            row.createCell(10).setCellValue("L9_PREF_COMMUNICATION");
            row.createCell(11).setCellValue("PAYMENT_METHOD");
            row.createCell(12).setCellValue("L9_RCID");
            row.createCell(13).setCellValue("L9_CPID");
            row.createCell(14).setCellValue("L9_PREFERRED_LANGUAGE");
            row.createCell(15).setCellValue("LEGAL_NAME_LINE1");
            row.createCell(16).setCellValue("LEGAL_NAME_LINE2");
            row.createCell(17).setCellValue("EMAIL");
            row.createCell(18).setCellValue("Address Field 1");
            row.createCell(19).setCellValue("Address Field 2");
            row.createCell(20).setCellValue("Address Field 3");
            row.createCell(21).setCellValue("Address Field 4");
            row.createCell(22).setCellValue("Address Field 5");
            row.createCell(23).setCellValue("Address Field 6");
            row.createCell(24).setCellValue("Address Field 7");
            row.createCell(25).setCellValue("Address Field 8");
            row.createCell(26).setCellValue("Address Field 9");
            row.createCell(27).setCellValue("Address Field 10");
            row.createCell(28).setCellValue("Address Field 11");
            row.createCell(29).setCellValue("Address Field 12");
            row.createCell(30).setCellValue("Address Field 13");
            row.createCell(31).setCellValue("Address Field 14");
            row.createCell(32).setCellValue("Address Field 15");
            row.createCell(33).setCellValue("PHONE1");
            row.createCell(34).setCellValue("PHONE2");
            row.createCell(35).setCellValue("PHONE3");
            row.createCell(36).setCellValue("PHONE4");
            row.createCell(37).setCellValue("WRITE_OFF_STATUS");
            row.createCell(38).setCellValue("FLAG  New or CLOSED");
            row.createCell(39).setCellValue("RCID CHANGE FLAG");
            row.createCell(40).setCellValue("OTHER CHANGE FLAG");
            
      
            Row row2 = sheet.createRow(1);  // Create 2nd Data row
            
            //Fetch STATUS_DATE, OPEN_DATE, address from  BANValue
            String STATUS_DATE = BANValue.get("creation_date");
            String Address_1= BANValue.get("address_1");
            String Address_2= BANValue.get("address_2");
            String Address_3= BANValue.get("address_3");
            String Address_4= BANValue.get("address_4");
            String Address_5= BANValue.get("address_5");
            String Address_6= BANValue.get("address_6");
            String Address_7= BANValue.get("address_7");
            String Address_8= BANValue.get("address_8");
            String Address_9= BANValue.get("address_9");
            String Address_10= BANValue.get("address_10");
            String Address_11= BANValue.get("address_11");
            String Address_12= BANValue.get("address_12");
            String Address_13= BANValue.get("address_13");
            String Address_14= BANValue.get("address_14");
            String Address_15= BANValue.get("address_15");
            
            //Populate the data in the row
            row2.createCell(0).setCellValue(STATUS_DATE);            //STATUS_DATE
            row2.createCell(1).setCellValue(STATUS_DATE);            //Open Date
            row2.createCell(4).setCellValue(amount);                 //amount
            row2.createCell(5).setCellValue(BAN);                    //BAN
            row2.createCell(6).setCellValue(1001);                   //DATA_SOURCE_ID
            row2.createCell(7).setCellValue(BANValue.get("ban_status"));     //BAN_STATUS
            row2.createCell(8).setCellValue("A");                    //CUSTOMER_TYPE. Default A.
            row2.createCell(9).setCellValue("B");                    //CUST_SUB_TYPE. Default B.
            row2.createCell(11).setCellValue("CA");                  //PAYMENT_METHOD. Default CA.
            
           
            String rcidString = (String) BANValue.get("rcid");       //L9_rcid. Remove ".0" and Convert to int
            if (rcidString.contains(".0")) {
                rcidString = rcidString.replace(".0", "");
            }
            try {
                int rcid = Integer.parseInt(rcidString); // Parse string to int
                //System.out.println("RCID: " + rcid);
                row2.createCell(12).setCellValue(rcid);
            } catch (NumberFormatException e) {
                e.printStackTrace();  // Handle the case where the string is not a valid integer
            }
            //row2.createCell(12).setCellValue(BANValue.get("rcid"));  //L9_RCID
            
            
            String cpidString = (String) BANValue.get("cbucid");   //L9_CPID. Remove ".0" and Convert to int
            if (cpidString.contains(".0")) {
            	cpidString = cpidString.replace(".0", "");
            }
            try {
                int cbucid = Integer.parseInt(cpidString); // Parse string to int
                //System.out.println("CPID: " + cbucid);
                row2.createCell(13).setCellValue(cbucid);
            } catch (NumberFormatException e) {
                e.printStackTrace();  // Handle the case where the string is not a valid integer
            }
            //row2.createCell(13).setCellValue(BANValue.get("cbucid")); //L9_CPID
            
            row2.createCell(14).setCellValue("EN");                   //L9_PREFERRED_LANGUAGE
            row2.createCell(15).setCellValue(BANValue.get("legal_name_line_1")); //LEGAL_NAME_LINE1
            row2.createCell(16).setCellValue(BANValue.get("legal_name_line_2")); //LEGAL_NAME_LINE2
            row2.createCell(17).setCellValue(eMail);                  //EMail
            row2.createCell(18).setCellValue(Address_1);              //address
            row2.createCell(19).setCellValue(Address_2);
            row2.createCell(20).setCellValue(Address_3);
            row2.createCell(21).setCellValue(Address_4);
            row2.createCell(22).setCellValue(Address_5);
            row2.createCell(23).setCellValue(Address_6);
            row2.createCell(24).setCellValue(Address_7);
            row2.createCell(25).setCellValue(Address_8);
            row2.createCell(26).setCellValue(Address_9);
            row2.createCell(27).setCellValue(Address_10);
            row2.createCell(28).setCellValue(Address_11);
            row2.createCell(29).setCellValue(Address_12);
            row2.createCell(30).setCellValue(Address_13);
            row2.createCell(31).setCellValue(Address_14);
            row2.createCell(32).setCellValue(Address_15);

            
            
            // Write the changes back to the file
            fos = new FileOutputStream(file);
            workbook.write(fos);

            System.out.println("Data written to CM Account file successfully!");    
            testResults.add("Test case " + i + " is passed - input_exp_cm_account_change");
	        

        }		catch (IOException e) {
        	System.out.println("catch exception"); 
            e.printStackTrace();
        } finally {
            // Close the resources
            try {
                if (fis != null) fis.close();
                if (fos != null) fos.close();
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }	

		
}
        }
	}