package com.telus.Properties;

import org.apache.poi.ss.usermodel.Workbook;

public class GetNumberOfSheet {
	
	//This method checks the number of sheets in the BAN.xlxs file. 
	//The number of sheets should be 1 else the program exists
	public static void  numberOfSheet(Workbook workbook) {
		
		int number = workbook.getNumberOfSheets();
		if (number>1) {
			System.out.println("The input BAN excel document has more than 1 sheet. Please check again. ");
			System.exit(0);
		
		}
	}
	
}