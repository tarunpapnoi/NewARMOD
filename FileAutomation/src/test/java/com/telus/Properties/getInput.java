package com.telus.Properties;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.telus.InputInvoiceFile.getBANDetails;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class getInput {

   
	public static void readExcelData(String filePath, String outputpath,String excelFilePath, String inputData,
			List<String> testResults) {

        
        int rowNum = 0;
        String billCycleStartDate = null;   //Date use yyyy-mm-dd FORMAT\
        Date billDate;
        int amount = 0;                  //Amount 
        String BAN_status = null;            //BAN Status
        String portfolio_cat = null;       //Portfolio
        int random_digit = 0;              //Random Digit  
        String customer_risk = null;             //Customer Risk
        String eMail = null; //Email
        int del_Cycle =0;                      //DelqCycle
        String outputpath1=outputpath;
        

        try (FileInputStream fis = new FileInputStream(new File(filePath))) {
            // Create a workbook object from the Excel file
            Workbook workbook = new XSSFWorkbook(fis);

            // Get the first sheet (sheet at index 0)
            Sheet sheet = workbook.getSheetAt(0);

            // Ensure the sheet is not null
            if (sheet == null) {
                System.out.println("The first sheet could not be found.");
                System.exit(0);
            }
            
            //Create a list to store the test result of each test case and pass this to every finction
              
            
            int noOfTestCases=sheet.getPhysicalNumberOfRows()-1;
            System.out.println("The no if test cases in user input file is " + noOfTestCases);

            // Iterate over rows starting from the second row (index 1)
            //This code iterates over the rows rows of User input where each row is a test case.
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            	
            	System.out.println("\n\n************Initiating Test Case "+i + "***********\n");
            	System.out.print("Details of test case number "+ i);
            	
            	
                Row row = sheet.getRow(i);
                
                
                
                // Skip empty rows
                if (row == null) continue;

                for (int col=0;col<9;col++) {
                	Cell cell = row.getCell(col);
                	
                	switch (col) {
                	
                	case 0:    //s.N from user input file
                		rowNum=(int) cell.getNumericCellValue();
                		System.out.println("\nRownumber " + rowNum);
                		break;
                	case 1:    //date from Userinput file
                		/*if (DateUtil.isCellDateFormatted(cell)) {
                            Date date = cell.getDateCellValue();
                            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy"); // Specify the desired date format
                            billCycleStartDate = sdf.format(date); // Convert date to string
                		}
                		else {
                            // If it's not a date, treat it as a numeric value
                            billCycleStartDate = String.valueOf(cell.getNumericCellValue());
                        } */
                		if (cell.getCellType() == CellType.NUMERIC) {
                            System.out.println("The cell billCycleStartDate contains numeric data.");
                            // Further handling if it's numeric
                        } else if (cell.getCellType() == CellType.STRING) {
                            System.out.println("The cell contains string data.");
                        } else if (cell.getCellType() == CellType.BOOLEAN) {
                            System.out.println("The cell contains boolean data.");
                        } else if (cell.getCellType() == CellType.FORMULA) {
                            System.out.println("The cell contains a formula.");
                        } else {
                            System.out.println("The cell contains other type of data.");
                        }
                		//billCycleStartDate=cell.getStringCellValue();
                		billDate=cell.getDateCellValue();
                		System.out.println(billDate);
                		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy"); // Desired format
                        billCycleStartDate = sdf.format(billDate); // Convert Date to String
                        System.out.println("Bill cycle start date "+ billCycleStartDate);
                		//billCycleStartDate=cell.getDateCellValue();
                		break;            //amount from user input
                	case 2:
                		amount=(int) cell.getNumericCellValue();
                		System.out.println("Amount "+amount);
                		break;
                	case 3:       //ban from user status file
                		BAN_status=cell.getStringCellValue();
                		System.out.println("BAN Status "+BAN_status);
                		break;
                	case 4:       //portfoli0 from user status file 
                		portfolio_cat=cell.getStringCellValue();
                		System.out.println("Portfolio "+portfolio_cat);
                		break;
                	case 5:    //random digit from user status file
                		random_digit=(int) cell.getNumericCellValue();
                		System.out.println("Random Digit "+random_digit);
                		break;
                	case 6:    //customer risk from user status file
                		customer_risk=cell.getStringCellValue();
                		System.out.println("Customer risk "+customer_risk);
                		break;
                	case 7:     //email from user status file
                		eMail=cell.getStringCellValue();
                		System.out.println("Email "+eMail);
                		break;
                	case 8:      //del cycle from user status file
                		del_Cycle=(int) cell.getNumericCellValue();
                		System.out.println("Del Cycle "+ del_Cycle);
                		System.out.println("******* ********\n");
                		break;
                	
                	}
                }
                
              //create files
                createInputFiles.createSheets(i,outputpath1);

              //common check
              
           	 getBANDetails.commonMethod(outputpath,excelFilePath, inputData,rowNum,billCycleStartDate,amount,BAN_status,portfolio_cat,random_digit,customer_risk,eMail,del_Cycle, testResults,i );
            }

            // Close the workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

       
    }


}

 