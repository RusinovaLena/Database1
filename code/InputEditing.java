package net.codejava;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InputEditing {
	public String[] inputValues(int rowIditing, int nList, int n) throws IOException {
		String[] initialData  = new String[n];
		
		try {			
			ArrayList data = new ArrayList();
			InputStream inputStream = new FileInputStream("Reestr.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			HSSFSheet currentSheet = workbook.getSheetAt(nList);
	        Iterator<Row> rowIterator = currentSheet.iterator();
	        
	        for (Row row: currentSheet) {
		        for (Cell cell: row) {
		        	if (row.getRowNum() == rowIditing) {
			        	switch (cell.getCellTypeEnum()) {
				        	case STRING:  				        	
					        	if ( cell.getColumnIndex() <= n ) {   
				                    if (cell.getColumnIndex() == n - 2) {                                   	                               	
				                    	data.add(cell.getStringCellValue());
				                    	data.add(cell.getRowIndex()); 
				                    } else {
						        		if (!cell.getStringCellValue().equals("")) {
						        			data.add(cell.getStringCellValue());   
						        		} else {
						        			data.add(""); 
						        		}
				                    }
			                    }		                                        
			                    break;
			        	default:	                    
		                    if (cell.getColumnIndex() <= n) {
		                    	if (nList == 0 && ( cell.getColumnIndex() == 9 | cell.getColumnIndex() == 10 ) ) {
				                    if (cell.getColumnIndex() == 9) {	                        	                                           		                            	
		                            	if (cell.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
		                                    data.add(ft.format(cell.getDateCellValue()));
		                            	}			                             		                    		                                 	
			                        }		                    
		                        
		                            if (cell.getColumnIndex() == 10) {
		                            	if (cell.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
		                                    data.add(ft.format(cell.getDateCellValue()));
		                            	} 
		                            }	
		                    	} else {
		                    		if ( cell.getColumnIndex() == n - 2) {
				                    	if (cell.getDateCellValue() == null) {
				                			data.add("");
				                		} else {
				                			data.add(String.valueOf((int) cell.getNumericCellValue()));
				                		}	
						        		data.add(cell.getRowIndex()); 
						        	} else {
				                    	if (cell.getDateCellValue() == null) {
				                			data.add("");
				                		} else {
				                			data.add(String.valueOf((int) cell.getNumericCellValue()));
				                		}
						        	}	
		                    	}
		                    }		                                                    		                    	
		                    break;
		        	}		        	
	        	}
        	}  
        }
        
	    System.out.println( "data: " + data );
	        
        for (int i = 0; i < data.size(); i ++) {
        	initialData[i]  = data.get(i).toString();
        }
	        
		} catch (IOException ex1) {
	        System.out.println("Error reading file");
	        ex1.printStackTrace();
	    }
		
	    catch (NullPointerException n1) {	        
	    }
		return initialData;
	}
}


