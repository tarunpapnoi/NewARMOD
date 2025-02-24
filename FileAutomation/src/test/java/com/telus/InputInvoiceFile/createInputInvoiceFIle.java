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

import static com.telus.InputInvoiceFile.entityIDinvoice.populateEntityIDinvoice;

public class createInputInvoiceFIle {
	
	public static void  populateIIFile(String outputpath,int rowNum, Map<String, String> BANValue, String billCycleStartDate, String BAN, int amount,
			List<String> testResults, int i) throws FileNotFoundException {
		
		//Input invoice file
		
		String filePath = outputpath+"\\input_invoice_YYYYMMDD"+rowNum+".xlsx";
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
            
            Row row = sheet.createRow(1);  // Create row at index 0 with column names
            row.createCell(0).setCellValue("BILLING CYCLE START DATE");
            row.createCell(1).setCellValue("BILLING CYCLE END DATE");
            row.createCell(2).setCellValue("INSTANCE");
            row.createCell(3).setCellValue("BILLING CYCLE YEAR");
            row.createCell(4).setCellValue("CYCLE CODE");
            row.createCell(5).setCellValue("TOTAL INVOICE AMT");
            row.createCell(6).setCellValue("TOTAL TAX AMT");
            row.createCell(7).setCellValue("BAN");
            row.createCell(8).setCellValue("DATA SOURCE ID");
            
         // Break date into yyyy / mm / dd
            String billDate = billCycleStartDate;  // billCycleStartDate is the start date

            // Parse the string into a LocalDate object
            //LocalDate date = LocalDate.parse(billDate, DateTimeFormatter.ISO_LOCAL_DATE);
            DateTimeFormatter formatter1 = DateTimeFormatter.ofPattern("dd-MM-yyyy");
            LocalDate date = LocalDate.parse(billDate, formatter1);

            // Format the date (remove dashes)
            String formattedDate = date.toString().replace("-", "");   // formattedDate in yyyyMMdd format

            // Given start date in yyyyMMdd format (from formatted date)
            String startDateStr = formattedDate;

            // Define the formatter to parse and format the date
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");

            // Parse the start date string into a LocalDate object
            LocalDate startDate = LocalDate.parse(startDateStr, formatter);

            // Move to the next month and subtract one day
            LocalDate endDate = startDate.plusMonths(1).minusDays(1);
            LocalDate startDateNextMonth= startDate.plusMonths(1).minusDays(0);   //Start date of next month
            //System.out.println("startDateNextMonth "+ startDateNextMonth);

            // Format the end date back into yyyyMMdd format
            String endDateStr = endDate.format(formatter);
            String startDate2 = startDateNextMonth.format(formatter);

            // Parse the endDateStr into a LocalDate object
            LocalDate endDate1 = LocalDate.parse(endDateStr, formatter);
            LocalDate startDateNext = LocalDate.parse(startDate2, formatter);

            // Extract year, month, and day from endDate (not billDate)
            int yyyy = endDate1.getYear();  // Year from endDate
            int mm = endDate1.getMonthValue();  // Month from endDate
            int dd = startDateNext.getDayOfMonth();  // Day+1 from endDate

            // Print the result
            System.out.println("End Date: " + endDateStr);

            // Add data to the next rows in the Excel sheet
            Row row2 = sheet.createRow(1);  // Row 1
            row2.createCell(0).setCellValue(formattedDate);   // BILLING CYCLE START DATE
            row2.createCell(1).setCellValue(endDateStr);      // BILLING CYCLE END DATE
            row2.createCell(2).setCellValue(mm);              // Instance (Month from endDate)
            row2.createCell(3).setCellValue(yyyy);            // BILLING CYCLE YEAR (from endDate)
            row2.createCell(4).setCellValue(dd);              // CYCLE CODE (Day from endDate)
            row2.createCell(5).setCellValue(amount);   // TOTAL INVOICE AMT
            row2.createCell(6).setCellValue(0);               // TOTAL TAX AMT
            row2.createCell(7).setCellValue(BAN);             // BAN
            row2.createCell(8).setCellValue(1001);            // DATA SOURCE ID

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
            entityIDinvoice.populateEntityIDinvoice(BANValue,  BAN,  i);

            System.out.println("Data written to Input Invoice file successfully!");
            testResults.add("Test case " + i + " is passed - Input Invoice");
		
}
        }
	}