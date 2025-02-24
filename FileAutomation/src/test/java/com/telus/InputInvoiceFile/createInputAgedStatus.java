package com.telus.InputInvoiceFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class createInputAgedStatus {
	
	public static void  populateInputAgedStatusFile(Map<String, String> BANValue,String outputpath, int rowNum, String BAN, int amount, int del_Cycle,
			List<String> testResults, int i) throws FileNotFoundException {
		
		//Input invoice file 
		
		String filePath = outputpath+"\\input_aged_status_YYYYMMDD"+rowNum+".xlsx";
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
            row.createCell(0).setCellValue("BAN");
            row.createCell(1).setCellValue("TR-1-30");
            row.createCell(2).setCellValue("TR-31-60");
            row.createCell(3).setCellValue("TR-61-90");
            row.createCell(4).setCellValue("TR-91-120");
            row.createCell(5).setCellValue("TR-121-150");
            row.createCell(6).setCellValue("TR-151-180");
            row.createCell(7).setCellValue("TR-180+");
            row.createCell(8).setCellValue("TR0");
            row.createCell(9).setCellValue("NT-1-30");
            row.createCell(10).setCellValue("NT-31-60");
            row.createCell(11).setCellValue("NT-61-90");
            row.createCell(12).setCellValue("NT-91-120");
            row.createCell(13).setCellValue("NT-121-150");
            row.createCell(14).setCellValue("NT-151-180");
            row.createCell(15).setCellValue("NT-180+");
            row.createCell(16).setCellValue("NT0");
            row.createCell(17).setCellValue("AD-1-30");
            row.createCell(18).setCellValue("AD-31-60");
            row.createCell(19).setCellValue("AD-61-90");
            row.createCell(20).setCellValue("AD-91-120");
            row.createCell(21).setCellValue("AD-121-150");
            row.createCell(22).setCellValue("AD-151-180");
            row.createCell(23).setCellValue("AD-180+");
            row.createCell(24).setCellValue("AD0");
            row.createCell(25).setCellValue("NineH-1-30");
            row.createCell(26).setCellValue("NineH-31-60");
            row.createCell(27).setCellValue("NineH-61-90");
            row.createCell(28).setCellValue("NineH-91-120");
            row.createCell(29).setCellValue("NineH-121-150");
            row.createCell(30).setCellValue("NineH-151-180");
            row.createCell(31).setCellValue("NineH-180+");
            row.createCell(32).setCellValue("NineH0");
            row.createCell(33).setCellValue("EQ-1-30");
            row.createCell(34).setCellValue("EQ-31-60");
            row.createCell(35).setCellValue("EQ-61-90");
            row.createCell(36).setCellValue("EQ-91-120");
            row.createCell(37).setCellValue("EQ-121-150");
            row.createCell(38).setCellValue("EQ-151-180");
            row.createCell(39).setCellValue("EQ-180+");
            row.createCell(40).setCellValue("EQ0");
            row.createCell(41).setCellValue("THIRDP-1-30");
            row.createCell(42).setCellValue("THIRDP-31-60");
            row.createCell(43).setCellValue("THIRDP-61-90");
            row.createCell(44).setCellValue("THIRDP-91-120");
            row.createCell(45).setCellValue("THIRDP-121-150");
            row.createCell(46).setCellValue("THIRDP-151-180");
            row.createCell(47).setCellValue("THIRDP-180+");
            row.createCell(48).setCellValue("THIRDP0");
            row.createCell(49).setCellValue("TOTAL-1-30");
            row.createCell(50).setCellValue("TOTAL-31-60");
            row.createCell(51).setCellValue("TOTAL-61-90");
            row.createCell(52).setCellValue("TOTAL-91-120");
            row.createCell(53).setCellValue("TOTAL-121-150");
            row.createCell(54).setCellValue("TOTAL-151-180");
            row.createCell(55).setCellValue("TOTAL-180+");
            row.createCell(56).setCellValue("TOTAL0");
            row.createCell(57).setCellValue("FLAG New or OLD");
      
            float installment;
            Row row2 = sheet.createRow(1);  // Create 2nd Data row
            
            for (int j=0;j<57;j++) {           //set all cells to 0
            	row2.createCell(j).setCellValue(0);
            }
            
            row2.createCell(0).setCellValue(BAN);            //BAN
            row2.createCell(57).setCellValue("NEW");         //Flag
            
            
            
            if (del_Cycle==1) {               //set amounts
            	installment=amount/5;
                row2.createCell(1).setCellValue(installment);   
                row2.createCell(9).setCellValue(installment);
                row2.createCell(17).setCellValue(installment);
                row2.createCell(25).setCellValue(installment);
                row2.createCell(33).setCellValue(amount-(installment*4));
                row2.createCell(49).setCellValue(amount);
            	
            }else
            	if(del_Cycle==2) {
            		installment=amount/10;
            		row2.createCell(1).setCellValue(installment);
                    row2.createCell(9).setCellValue(installment);
                    row2.createCell(17).setCellValue(installment);
                    row2.createCell(25).setCellValue(installment);
                    row2.createCell(33).setCellValue(installment);
                    row2.createCell(49).setCellValue(installment*5);
                    
                    row2.createCell(2).setCellValue(installment);
                    row2.createCell(10).setCellValue(installment);
                    row2.createCell(18).setCellValue(installment);
                    row2.createCell(26).setCellValue(installment);
                    row2.createCell(34).setCellValue(amount-(installment*9));
                    row2.createCell(50).setCellValue(amount - (installment*5));
            		
            	}
            	else if (del_Cycle==3){
            		installment=amount/15;
            		row2.createCell(1).setCellValue(installment);
                    row2.createCell(9).setCellValue(installment);
                    row2.createCell(17).setCellValue(installment);
                    row2.createCell(25).setCellValue(installment);
                    row2.createCell(33).setCellValue(installment);
                    row2.createCell(49).setCellValue(installment*5);
                    
                    row2.createCell(2).setCellValue(installment);
                    row2.createCell(10).setCellValue(installment);
                    row2.createCell(18).setCellValue(installment);
                    row2.createCell(26).setCellValue(installment);
                    row2.createCell(34).setCellValue(installment);
                    row2.createCell(50).setCellValue(installment*5);
                    
                    row2.createCell(3).setCellValue(installment);
                    row2.createCell(11).setCellValue(installment);
                    row2.createCell(19).setCellValue(installment);
                    row2.createCell(27).setCellValue(installment);
                    row2.createCell(35).setCellValue(amount-(installment*14));
                    row2.createCell(51).setCellValue(amount - (installment*10));
            		
            	}
            

            // Write the changes back to the file
            fos = new FileOutputStream(file);
            workbook.write(fos);


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

            System.out.println("populateEntityID");
            entityIDAged.populateEntityIDAged(BANValue,  BAN, i);

            System.out.println("Data written to Input aged Status file successfully!");
            testResults.add("Test case " + i + " is passed - input_aged_status_YYYYMMDD");

		
}
        }
	}