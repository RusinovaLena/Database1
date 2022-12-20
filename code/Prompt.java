package net.codejava;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Prompt {
	ArrayList<String> advice = new ArrayList<>();
	
	public ArrayList<String> suggestion() {			
		try {   		
			InputStream inputStream = new FileInputStream("Reestr.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
	        Sheet currentSheet = workbook.getSheetAt(0);
	        Iterator<Row> rowIterator = currentSheet.iterator();
	        rowIterator.next();
	        Row nextRow2 = rowIterator.next(); 
	        
	        for (Row row: currentSheet) {
	        	if (row.getRowNum() > 1) {
			        for (Cell cell: row) {			        	
			        	if (cell.getColumnIndex() == 1 && cell.getStringCellValue().length() != 0) {
                        	String s = cell.getStringCellValue();
                        	advice.add(s);                             	
                        }
			        }
	        	}    
	        }    
	        
            Set set = new HashSet(advice);
            advice.clear();
            advice = new ArrayList(set); 
	            
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			catch (IOException e) {
				e.printStackTrace();
			}
			return advice;	
			
		}
}
