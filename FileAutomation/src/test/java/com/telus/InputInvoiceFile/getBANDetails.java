package com.telus.InputInvoiceFile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.io.File;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class getBANDetails {
	
	public static void  commonMethod(String outputpath,String excelFilePath, String inputData, int rowNum,
    String billCycleStartDate,
    int amount,
    String BAN_status,
    String portfolio_cat,
    int random_digit,
    String customer_risk,
    String eMail,
    int del_Cycle,
    List<String> testResultstestResults,int i ) throws FileNotFoundException {
		
		/*
		//Inputs for BAN selection
		//String rule="rule2";      //mention the rule here. To be implemented in future
		String billCycleStartDate="2024-12-20";   //Date use yyyy-mm-dd FORMAT
        int amount=2000;                  //Amount 
        String BAN_status="O";            //BAN Status
        String portfolio_cat="SMB";       //Portfolio
        int random_digit=15;              //Random Digit  
        String customer_risk= "";             //Customer Risk
        String eMail= "mohd.sadiq@telus.com"; //Email
        int del_Cycle=3;                      //DelqCycle
        */
		
		
		String excelFilePath1 = excelFilePath;
		//get index of columns from BAN doc
		try (FileInputStream fis = new FileInputStream(new File(excelFilePath1));
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheetAt(0);
	            Iterator<Row> rowIterator = sheet.iterator();

	           
	            Row headerRow = rowIterator.next(); // Skipping header row
	            
	            //init the column index here
	            int banIndex= -1;
	            int creation_dateIndex= -1;
	            int rcidIndex= -1;
	            int closing_cycleIndex= -1;
	            int ban_statusIndex= -1;
	            int entity_idIndex= -1;
	            int random_digitIndex = -1;
	            int legal_name_line_1Index = -1;
	            int legal_name_line_2Index = -1;
	            int CbucidIndex= -1;
	            int portfolioIndex=-1;
	            int customer_riskIndex=-1;
	            int address_1Index	 = -1;
	            int address_2Index	 = -1;
	            int address_3Index	 = -1;
	            int address_4Index	 = -1;
	            int address_5Index	 = -1;
	            int address_6Index	 = -1;
	            int address_7Index	 = -1;
	            int address_8Index	 = -1;
	            int address_9Index	 = -1;
	            int address_10Index	 = -1;
	            int address_11Index	 = -1;
	            int address_12Index	 = -1;
	            int address_13Index	 = -1;
	            int address_14Index	 = -1;
	            int address_15Index = -1;
	            
	            ArrayList<Integer> statusIndex = new ArrayList<>();
	            
	            Iterator<Cell> headerCellIterator = headerRow.cellIterator();
	            
	            //Get index of columns from BAN sheet********************************
	            while (headerCellIterator.hasNext()) {
	                Cell headerCell = headerCellIterator.next();
	                String str1=headerCell.getStringCellValue().toLowerCase();
	                
	                switch (str1) {
	                case ("ban"):
	                	banIndex=headerCell.getColumnIndex();
	                    //System.out.println("ban index is "+ banIndex);
	                    break;
	                case ("creation_date"):
	                	creation_dateIndex=headerCell.getColumnIndex();
	                    //System.out.println("creation_date index is "+ creation_dateIndex);
	                    break;
	                case ("rcid"):
	                	rcidIndex=headerCell.getColumnIndex();
	                    //System.out.println("rcid index is "+ rcidIndex);
                        break;
	                case ("closing_cycle"):
	                	closing_cycleIndex=headerCell.getColumnIndex();
                        break;
	                case ("ban_status"):
	                	ban_statusIndex=headerCell.getColumnIndex();
	                    break;
	                case ("entity_id"):
	                	entity_idIndex=headerCell.getColumnIndex();
	                    break;
	                case ("random_digit"):
	                	random_digitIndex=headerCell.getColumnIndex();
	                    break;
	                case ("legal_name_line_1"):
	                	legal_name_line_1Index=headerCell.getColumnIndex();
	                    break;
	                case ("legal_name_line_2"):
	                	legal_name_line_2Index=headerCell.getColumnIndex();
	                    break;
	                case ("cbucid"):
	                	CbucidIndex=headerCell.getColumnIndex();
	                    break;
	                case("portfolio_cat"):{
	                	portfolioIndex=headerCell.getColumnIndex();
	                	break;
	                    }
	                case ("address_1"):
	                	address_1Index=headerCell.getColumnIndex();
	                    break;
	                case ("address_2"):
	                	address_2Index=headerCell.getColumnIndex();
	                break;
	                case ("address_3"):
	                	address_3Index=headerCell.getColumnIndex();
	                break;
	                case ("address_4"):
	                	address_4Index=headerCell.getColumnIndex();
	                break;
	                case ("address_5"):
	                	address_5Index=headerCell.getColumnIndex();
	                break;
	                case ("address_6"):
	                	address_6Index=headerCell.getColumnIndex();
	                break;
	                case ("address_7"):
	                	address_7Index=headerCell.getColumnIndex();
	                break;
	                case ("address_8"):
	                	address_8Index=headerCell.getColumnIndex();
	                break;
	                case ("address_9"):
	                	address_9Index=headerCell.getColumnIndex();
	                break;
	                case ("address_10"):
	                	address_10Index=headerCell.getColumnIndex();
	                break;
	                case ("address_11"):
	                	address_11Index=headerCell.getColumnIndex();
	                break;
	                case ("address_12"):
	                	address_12Index=headerCell.getColumnIndex();
	                break;
	                case ("address_13"):
	                	address_13Index=headerCell.getColumnIndex();
	                break;
	                case ("address_14"):
	                	address_14Index=headerCell.getColumnIndex();
	                break;
	                case ("address_15"):
	                	address_15Index=headerCell.getColumnIndex();
	                break;
	                case ("customer_risk"):
	                	customer_riskIndex=headerCell.getColumnIndex();
	                break;
	                	
	                }
                        
                        
	            }
	            //System.out.println("portfolio index is "+ portfolioIndex);
	            //System.out.println("customer_risk index is "+ customer_riskIndex);

	            
	          //check if status column is present in the Bans list********************************
	            if (ban_statusIndex == -1 || banIndex== -1 || portfolioIndex== -1 || random_digitIndex== -1 || customer_riskIndex== -1) {
	                System.out.println("One or more mandatory columns are missing in BAN list. please check the BAN file");
	                System.out.println(" Status, BAN, portfolio, random digit or cust risk column not found in the Excel file.");
	                System.exit(0);
	            }
	            
	            //System.out.println("customer_risk " + customer_risk);
	            
		       // get index of rows where status is C/O and save in an arraylist based on rule****************
	            System.out.println("Iterating over rows and columns to fetch the index of rows based on our requirements");
	            while (rowIterator.hasNext()) {
	                Row row = rowIterator.next();
	                Cell statusCell = row.getCell(ban_statusIndex);
	                Cell portfolioCell = row.getCell(portfolioIndex);
	                Cell rDigitCell = row.getCell(random_digitIndex);
	                Cell cRiskCell = row.getCell(customer_riskIndex);
	                
	                //System.out.println(row.getRowNum());
	                //System.out.println(" statusCell.getCellType()" +  statusCell.getCellType());
	                //System.out.println("BAN_status.equalsIgnoreCase(statusCell.getStringCellValue()" + BAN_status.equalsIgnoreCase(statusCell.getStringCellValue()));
	                //System.out.println("portfolio_cat.equalsIgnoreCase(portfolioCell.getStringCellValue() " + portfolio_cat.equalsIgnoreCase(portfolioCell.getStringCellValue()));
	                //System.out.println("rDigitCell.getNumericCellValue() "+ rDigitCell.getNumericCellValue());
	                //System.out.println("customer_risk==cRiskCell.getStringCellValue() "+ customer_risk==cRiskCell.getStringCellValue());
	                
	                
	                // Filter the BANs based on Inputs
	                //System.out.println("customer_risk " + customer_risk);
	                //System.out.println("cRiskCell.getStringCellValue() " + cRiskCell.getStringCellValue());
	                //System.out.println("customer_risk " + customer_risk);
	                if (statusCell != null && statusCell.getCellType() == CellType.STRING 
	                        && BAN_status.equalsIgnoreCase(statusCell.getStringCellValue())
	                        && portfolioCell != null
	                        && portfolio_cat.equalsIgnoreCase(portfolioCell.getStringCellValue())
	                        && rDigitCell.getNumericCellValue()>=random_digit
	                        && customer_risk.equalsIgnoreCase(cRiskCell.getStringCellValue())
	                        ) {

	                	statusIndex.add(row.getRowNum());
	                	//System.out.println("Status Index added a BAN");

	                }
	              /* if(statusCell != null && statusCell.getCellType() == CellType.STRING ) {
	                	if(BAN_status.equalsIgnoreCase(statusCell.getStringCellValue())) {
	                		if(portfolioCell != null  && portfolio_cat.equalsIgnoreCase(portfolioCell.getStringCellValue())) {
	                			if(rDigitCell.getNumericCellValue()>=random_digit ) {
	                				if(customer_risk==cRiskCell.getStringCellValue() ) {
	                					statusIndex.add(row.getRowNum());
	                				}
	                				//else {System.out.println("Issue in customer risk");}
	                			}//else {System.out.println("Issue in random_digit");}
	                		}//else {System.out.println("Issue in portfolio");}
	                	}//else {System.out.println("Issue in BAN_status");}
	                }//else {System.out.println("Issue in getCellType");}
	                */
	                
	               }

	            System.out.println("Status Index" +statusIndex);
	            
	           //********************************************************************** 
	          //fetch ban from excel sheet from random index no and save to variable************
	            String BAN;
	            if (statusIndex.size() == 0) {
	                System.out.println("The ArrayList statusIndex is empty.");
	            } else {
	                System.out.println("The ArrayList  statusIndex is not empty.");
	            }
	            
	            Random random = new Random();

	            // Generate a random index between 0 and size of the ArrayList - 1
	            int randomIndex = random.nextInt(statusIndex.size());

	            // Get the random element from the ArrayList
	            Integer randomElement = statusIndex.get(randomIndex);

	            // Print the random element
	            System.out.println("Random element from the ArrayList: " + randomElement);
	            
	            Cell cell = sheet.getRow(randomElement).getCell(banIndex);
	            
	            BAN=String.valueOf((long)cell.getNumericCellValue());
	            System.out.println("Random BAN  based on that element from the ArrayList: " + BAN);
	            
	            Map<String, Integer> rowMap = new HashMap<>();
	            
	         // Get the first row (header row)
	            headerRow = sheet.getRow(0);
	            
	         // Iterate through each cell in the header row
	            Iterator<Cell> cellIterator = headerRow.cellIterator();
	            while (cellIterator.hasNext()) {
	                Cell cell1 = cellIterator.next();
	                // Get the column name (header) and its index
	                String columnName = cell1.getStringCellValue();
	                int columnIndex = cell1.getColumnIndex();
	                rowMap.put(columnName, columnIndex);
	            }
	            
	         // Now, retrieve a specific row (e.g., row 2, which is index 1)
	            Row dataRow = sheet.getRow(randomElement);
	            
	         // Create a map to store the data for the specific row, using column names as keys
	            Map<String, String> BANValue = new HashMap<>();
	            
	            // Iterate over the headerMap to fetch the corresponding data for each column
	            for (Map.Entry<String, Integer> entry : rowMap.entrySet()) {
	                String columnName = entry.getKey();
	                int columnIndex = entry.getValue();
	                
	                // Get the cell value for the specific column and row
	                Cell cell1 = dataRow.getCell(columnIndex);
	                
	                // Get the cell value (as String) and store it in the map
	                String cellValue = cell1 != null ? cell1.toString() : ""; // Handle nulls
	                BANValue.put(columnName, cellValue);
	            }
	            
	         // Print the resulting map
	            System.out.println("Data Row Map: " + BANValue);
	            
	            
	            //Row dataRow = sheet.getRow(statusIndex);
	            
	      //*****call function to populate input invoice file
	       createInputInvoiceFIle.populateIIFile(outputpath, rowNum,BANValue, billCycleStartDate, BAN, amount,testResultstestResults,i);
	       
	     //*****call function to populate input cm account change  file
	       createCmAccountFIle.populateCMAccountFile(outputpath, rowNum,BANValue, BAN, amount, eMail,testResultstestResults,i);
	       
	     //*****call function to populate input cm account change  file
	       createInputAgedStatus.populateInputAgedStatusFile(outputpath, rowNum,BAN, amount, del_Cycle,testResultstestResults,i);

	            
	            } catch (IOException e) {
	            e.printStackTrace();
	        }
		
		
		

	    }
				

		
}