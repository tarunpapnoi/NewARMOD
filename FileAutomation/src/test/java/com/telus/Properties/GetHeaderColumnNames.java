package com.telus.Properties;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


public class GetHeaderColumnNames {
	
	//This method checks the number of sheets in the BAN.xlxs file. 
	//The number of sheets should be 1 else the program exists
	public static void  getColumnNames(Sheet sheet) {
		
		Row headerRow = sheet.getRow(0);
   	 	
   	 final Map<String, Integer> headerColumnIndexMap = new HashMap<>();
   	 Iterator<Cell> cellIterator = headerRow.cellIterator();
   	 while(cellIterator.hasNext()) {
   		 
   		 Cell cell = cellIterator.next();
   		 if(cell.getCellType()==CellType.STRING) {
   			 String columnName = cell.getStringCellValue();
   			 int columnIndex = cell.getColumnIndex();
   			 
   			 headerColumnIndexMap.put(columnName, columnIndex);
   		 }
   	 }
   	 
   	 //System.out.println("Column Names and their Indexes:");
   	 int banInit = 0;
   	 for(Map.Entry<String, Integer> entry : headerColumnIndexMap.entrySet()) {
   		 //System.out.println(entry.getKey()  + " -- " + entry.getValue());
   		 if(entry.getKey().toLowerCase() == "ban") {
   			banInit=1;
   			//System.out.println("BAN column found" + banInit);
   			
   		 }
   		 
   	 }
   	 
  // Alternatively, provide a getter method
     
   	 
   	//if(banInit== 0) {
		//	System.out.println("There is no column for BANs in the input file" + banInit);
			//System.exit(0);
		 //}
	}
	
}