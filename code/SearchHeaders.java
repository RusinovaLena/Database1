package net.codejava;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SearchHeaders {
	
	public int search(String headerCurrent) throws IOException {
		
        InputStream inputStream = new FileInputStream("Reestr.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet currentSheet = workbook.getSheetAt(0);
        int currentRow = currentSheet.getLastRowNum() + 1;
        Iterator<Row> rowIterator = currentSheet.iterator();
        rowIterator.next();	        
        
        boolean b = false;
        boolean b2 = false;
        for (Row row: currentSheet) {
        	if (row.getRowNum() > 1) {
		        for (Cell cell: row) {
		        	
		        	if (cell.getColumnIndex() == 2 && cell.getStringCellValue().equals(headerCurrent)) {
	                	b = true;
	                	b2 = false;
	                	
	                } else {
	                	b2 = true;
	                }
	                
	                if (b == true && b2 == true) {
	                	if (headerCurrent.length() != 0) {
	                		currentRow = cell.getRowIndex() + 1;
	                	} else {
	                		currentRow = currentSheet.getLastRowNum() + 1;
	                	}
	                	b = false;
	                    b2 = false;
	                	break;
	                }
		        }
        	}    
        }    	    
	    return currentRow;    
	}    
	
	public int[] searchTwoRow(String headerCurrent) throws IOException {
		
        InputStream inputStream = new FileInputStream("Reestr.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet currentSheet = workbook.getSheetAt(0);
        int[] currentRows = new int[2];
        
        currentRows[1] = currentSheet.getLastRowNum() + 1;
        Iterator<Row> rowIterator = currentSheet.iterator();
        rowIterator.next();	   
        
        boolean b = false;
        boolean b2 = false;
        int ch = 0;
        
        for (Row row: currentSheet) {
        	if (row.getRowNum() > 1) {
		        for (Cell cell: row) {
		        	
		        	if (cell.getColumnIndex() == 2 && cell.getStringCellValue().equals(headerCurrent)) {
                    	if (ch == 0) {
                    		// начальная строка заголовка
                    		currentRows[0] = cell.getRowIndex() - 2;
                    		ch++;
                    	}	
                    	b = true;
                    	b2 = false;
                    } else {
                    	b2 = true;
                    }
                    
                    if (b == true && b2 == true) {
                    	// конечная строка заголовка
                    	currentRows[1] = cell.getRowIndex() - 1;
                    	b = false;
                        b2 = false;
                    	break;
                    }
		        }
        	}    
        }
	    return currentRows;    
	}
}    
