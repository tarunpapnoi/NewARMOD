package com.telus.Properties;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.telus.InputInvoiceFile.createInputAgedStatus;
import com.telus.InputInvoiceFile.getBANDetails;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class ExcelReader {
    public static void main(String[] args) throws IOException {
    	
    	//*******************************************************************************************
    	// Load properties from config file
        Properties properties = new Properties();
        String configFilePath = "src/main/resources/config.properties";  // Adjust path accordingly

        try (FileInputStream input = new FileInputStream(configFilePath)) {
            properties.load(input);
        } catch (IOException e) {
            System.out.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        // Get paths from properties file
        String excelFilePath = properties.getProperty("inputBANFile1");
        String inputData = properties.getProperty("UserInput");
        String outputpath = properties.getProperty("outputpath");
        List<String> testResults = new ArrayList<>(); 
        
        //**********************************************************************************************
    	
    	//String excelFilePath = "D:\\Users\\tarun.papnoi\\ARMODFile\\FileAutomation\\TestData\\BusCon validation.xlsx";
    	//*String excelFilePath = "D:\\Users\\tarun.papnoi\\ARMODFile\\FileAutomation\\TestData\\Copy of ITN01-BANs-20241007.xlsx";
    	//*String inputData = "C:\\Users\\tarun.papnoi\\Downloads\\UserInput.xlsx";
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath) );  //create the object of file
        	Workbook workbook = new XSSFWorkbook(fis)) 		
        {
        	// Check Number of sheets. Exit the program if no of sheet > 1.
        	//GetNumberOfSheet.numberOfSheet(workbook); 
        	
        	 Sheet sheet = workbook.getSheetAt(0);
        	 
        	 //Save the header columns name and index in a map
        	 GetHeaderColumnNames.getColumnNames(sheet);

        	 
        	 getInput.readExcelData(inputData,outputpath,excelFilePath,inputData,testResults);
        	 
        	 //common check
        	 //getBANDetails.commonMethod();
        	 
        	 //Print the Test case result
        	 System.out.println("\n\n*********The test results are as follows: **********");
        	 for (String result : testResults) {
                 System.out.println(result);
             }
        	 
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
    
}
    
}