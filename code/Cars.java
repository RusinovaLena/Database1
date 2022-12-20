package net.codejava;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import javax.swing.*;
import javax.swing.border.LineBorder;
import javax.swing.table.*;
import java.io.InputStream;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class Cars extends JFrame  { // приборы и расходники 
	
	final static Object[] mainHeaders = {"№", "Марка, модель", "VIN",
		    "<html> <center>Регистрационный<br> номер", "<html> <center>Год<br>выпуска", "Пробег", "ПТС", "СТС",
		    "<html> <center>Страховая<br>компания", "№ полиса", "<html> <center>Срок<br>действия", "Страховая компания", "№ полиса",
		    "<html> <center>Срок<br>действия",  
		    "<html> <center>Техническое<br> состояние", 
		    "<html> <center>Форма<br> собственности", 
            "<html> <center>Владелец<br> оборудавания", "Местонахождение", "<html> <center>Ответственный<br> владелец", "Примечание", " "};
		    // "<html> <center>Владелец<br> оборудавания", "Местонахождение", "<html> <center>Ответственный<br> владелец", "Примечание"};
	           
    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
    
    int startSize = 0;
    JTextField yearRelease = new JTextField(15);
    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
    
    static JFrame frame = new JFrame();
    static ArrayList <String>copyData = new ArrayList<String>();
    static ArrayList jointUpload = new ArrayList();
    
    ArrayList listYears = new ArrayList();
    ArrayList kindControl = new ArrayList();
    ArrayList formsOwnership = new ArrayList();
    ArrayList equipmentOwners = new ArrayList();
    ArrayList allLocations = new ArrayList();
    ArrayList modelBrand = new ArrayList();  
    
    ArrayList<Integer> dateGreen10 = new ArrayList<Integer>();
    ArrayList<Integer> dateYellow10 = new ArrayList<Integer>();
    ArrayList<Integer> dateRed10 = new ArrayList<Integer>();
    
    ArrayList<Integer> dateGreen13 = new ArrayList<Integer>();
    ArrayList<Integer> dateYellow13 = new ArrayList<Integer>();
    ArrayList<Integer> dateRed13 = new ArrayList<Integer>();
    
    static int cH = 0;
    static int cW = 0;
    
    public void start(JFrame frame, int hThree, JTable tableStart) throws IOException, ParseException {
    	
    	// установка другой иконки для JFrame
    	/*
    	ImageIcon liderIcon = new ImageIcon(new Cars().getClass().getClassLoader().getResource(".png"));
        Image image = liderIcon.getImage();
        frame.setIconImage(image);
        */
        
    	frame.getContentPane().setLayout(new BorderLayout());
    	Font myFont = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 15);
    	JButton buttonSearch = new JButton("Поиск");
        JButton buttonStart = new JButton("<html><center>" + "Сбросить параметры поиска" + "<center><html>"); 
      
        JPanel panelMain = new JPanel();
   	 	panelMain.add(yearRelease).setFont(myFont);
   	    
        JPanel panelB = new JPanel(new GridLayout(0, 7, 0, 0));
        
        panelB.add(new JLabel(" "));
        JCheckBox fieldSub = new JCheckBox("", false);
   	    new WorkingCheckBox().scaleCheckBoxIcon(fieldSub, 25);
   	    
   	    panelB.add(new JLabel("<html>Добавить данные в<br>совместную выгрузку:</html>", SwingConstants.RIGHT)).setFont(bigFontTR);
   	    panelB.add(fieldSub).setFont(bigFontTR); 	            
        panelB.add(buttonSearch).setFont(bigFontTR);
        panelB.add(new JLabel(" "));
        panelB.add(buttonStart).setFont(bigFontTR);
        panelB.add(new JLabel(" "));
               
        ArrayList data = new ArrayList();

        try {
        	InputStream inputStream = new FileInputStream("Reestr.xls");
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet currentSheet = workbook.getSheetAt(1);
            String sD10 = null;
            String sD13 = null;
            
		    for (Row row: currentSheet) {
		    	boolean nRowYellow10 = false;   
                boolean nRowRed10 = false; 
                boolean nRowYellow13 = false;   
                boolean nRowRed13 = false; 
                String number = "";
                String name = "";		    	
		        for (Cell cell: row) {
			    	if (row.getRowNum() > 2) {
				    	int columnIndex = cell.getColumnIndex();
				    	
				        switch (cell.getCellTypeEnum()) {
					        case STRING:			        				        
					        	
		                        if (cell.getColumnIndex() == 10) {
		                        	boolean sameDay = false; 
		                        	Date dateNow = new Date();
		                        	Calendar cal1 = Calendar.getInstance();
		                            Calendar cal2 = Calendar.getInstance();
		                            	
		                        	String regex = "(\\d{2}.\\d{2}.\\d{4})";
		                    		Matcher m = Pattern.compile(regex).matcher(cell.getStringCellValue());
		                    		
		                    		if (m.find()) {
		                    			Date docDate = null;
		                    			
										try {
											docDate = new SimpleDateFormat("dd.MM.yyyy").parse(m.group(1));
										} catch (ParseException e1) {
											e1.printStackTrace();
										}
		                                cal1.setTime(dateNow);
		                                cal2.setTime(docDate);
		                    		} else {
		                    			sameDay = false;
		                    			sD10 = String.valueOf(sameDay);
		                    			break;
		                    		}
		                            
		                            // анализ 10 столбца
		                            if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
		                            	sameDay = true; 
		                            	
		                            	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
		                            		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                		nRowYellow10 = true;
		                                	} else {
		                                		nRowYellow10 = false;
		                                	}
		                                	
		                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                		nRowRed10 = true;
		                                	} else {
		                                		nRowRed10 = false;
		                                	}
		                            	}
		                            	
		                            } else {
		                                if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
		                                    sameDay = false;
		                                } else {
		                                	// если сегодняшний месяц меньше
		                                    if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
		                                    	sameDay = true;
		                                    	
		                                           // если текущий месяц больше на 1
		                                        if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
		                                        	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                        		nRowYellow10 = true;
		                                        	} else {
		                                        		nRowYellow10 = false;
		                                        	}
		                                        	
		                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                        		nRowRed10 = true;
		                                        	} else {
		                                        		nRowRed10 = false;
		                                        	}
		                                    	}                                                                     
		                                    } else {
		
		                                        if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
		                                            sameDay = false;  
		                                        } else {          
		                                        	// если месяцы равны
		                                            if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
		                                                sameDay = true;
		                                                nRowYellow10 = true;
		                                                
		                                             // если месяцы равны и текущий день <= 7
		                                            	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                            		nRowRed10 = true;
		                                            	} else {
		                                            		nRowRed10 = false;
		                                            	}
		                                            } else {
		                                                sameDay = false;  	                                                                                        
		                                            }
		                                        }
		                                    }
		                                }
		                            }                                                              
		                            sD10 = String.valueOf(sameDay);	
		                        }
		                        if (cell.getColumnIndex() == 13) {
		                        	boolean sameDay = false; 
		                        	Date dateNow = new Date();
		                        	Calendar cal1 = Calendar.getInstance();
		                            Calendar cal2 = Calendar.getInstance();
		                            	
		                        	String regex = "(\\d{2}.\\d{2}.\\d{4})";
		                    		Matcher m = Pattern.compile(regex).matcher(cell.getStringCellValue());
		                    		
		                    		if (m.find()) {
		                    			Date docDate = null;
		                    			
										try {
											docDate = new SimpleDateFormat("dd.MM.yyyy").parse(m.group(1));
										} catch (ParseException e1) {
											e1.printStackTrace();
										}
		                                cal1.setTime(dateNow);
		                                cal2.setTime(docDate);
		                    		} else {
		                    			sameDay = false;
		                    			sD13 = String.valueOf(sameDay);
		                    			break;
		                    		}
		                            
		                            // анализ 10 столбца
		                            if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
		                            	sameDay = true; 
		                            	
		                            	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
		                            		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                		nRowYellow13 = true;
		                                	} else {
		                                		nRowYellow13 = false;
		                                	}
		                                	
		                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                		nRowRed13 = true;
		                                	} else {
		                                		nRowRed13 = false;
		                                	}
		                            	}
		                            	
		                            } else {
		                                if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
		                                    sameDay = false;
		                                } else {
		                                	// если сегодняшний месяц меньше
		                                    if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
		                                    	sameDay = true;
		                                    	
		                                           // если текущий месяц больше на 1
		                                        if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
		                                        	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                        		nRowYellow13 = true;
		                                        	} else {
		                                        		nRowYellow13 = false;
		                                        	}
		                                        	
		                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                        		nRowRed13 = true;
		                                        	} else {
		                                        		nRowRed13 = false;
		                                        	}
		                                    	}                                                                     
		                                    } else {
		
		                                        if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
		                                            sameDay = false;  
		                                        } else {          
		                                        	// если месяцы равны
		                                            if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
		                                                sameDay = true;
		                                                nRowYellow13 = true;
		                                                
		                                             // если месяцы равны и текущий день <= 7
		                                            	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                            		nRowRed13 = true;
		                                            	} else {
		                                            		nRowRed13 = false;
		                                            	}
		                                            } else {
		                                                sameDay = false;  	                                                                                        
		                                            }
		                                        }
		                                    }
		                                }
		                            }                                                              
		                            sD13 = String.valueOf(sameDay);	
		                        }
		                        break;
					        default:               	   
		                	   
		                	   if (cell.getColumnIndex() == 10) {
		                       	boolean sameDay = false; 
		                       	Date dateNow = new Date();
		                       	Calendar cal1 = Calendar.getInstance();
		                        Calendar cal2 = Calendar.getInstance();
		                       	sameDay = false; 
		                       	
		                       	if (cell.getDateCellValue() != null) {                                                                                       
		                               cal1.setTime(dateNow);
		                               cal2.setTime(cell.getDateCellValue());
		                       	} else {
		                       		break;
		                       	}
		                       	
		                        // анализ 10 столбца
		                       if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
		                       	sameDay = true; 
		                       	
		                       	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
		                       		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                           		nRowYellow10 = true;
		                           	} else {
		                           		nRowYellow10 = false;
		                           	}		                           	
		                           	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                           		nRowRed10 = true;
		                           	} else {
		                           		nRowRed10 = false;
		                           	}
		                       	}		                       	
		                       } else {
		                           if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
		                               sameDay = false;
		                           } else {
		                           	// если сегодняшний месяц меньше
		                               if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
		                               	sameDay = true;
		                               	
		                                      // если текущий месяц больше на 1
		                                   if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
		                                   	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                   		nRowYellow10 = true;
		                                   	} else {
		                                   		nRowYellow10 = false;
		                                   	}
		                                   	
		                                   	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                   		nRowRed10 = true;
		                                   	} else {
		                                   		nRowRed10 = false;
		                                   	}
		                               	}                                                                     
		                               } else {	
		                                   if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
		                                       sameDay = false;  
		                                   } else {          
		                                   	// если месяцы равны
		                                       if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
		                                           sameDay = true;
		                                           nRowYellow10 = true;
		                                           
		                                        // если месяцы равны и текущий день <= 7
		                                       	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                       		nRowRed10 = true;
		                                       	} else {
		                                       		nRowRed10 = false;
		                                       	}
		                                       } else {
		                                           sameDay = false;  	                                                                                        
		                                       }
		                                   }
		                               }
		                           }
		                       }                                                              
		                       sD10 = String.valueOf(sameDay);	
		                   } 
	                	   if (cell.getColumnIndex() == 13) {
		                       	boolean sameDay = false; 
		                       	Date dateNow = new Date();
		                       	Calendar cal1 = Calendar.getInstance();
		                        Calendar cal2 = Calendar.getInstance();
		                       	sameDay = false; 
		                       	
		                       	if (cell.getDateCellValue() != null) {                                                                                       
		                               cal1.setTime(dateNow);
		                               cal2.setTime(cell.getDateCellValue());
		                       	} else {
		                       		break;
		                       	}
		                       	
		                        // анализ 10 столбца
		                       if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
		                       	sameDay = true; 
		                       	
		                       	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
		                       		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                           		nRowYellow13 = true;
		                           	} else {
		                           		nRowYellow13 = false;
		                           	}
		                           	
		                           	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                           		nRowRed13 = true;
		                           	} else {
		                           		nRowRed13 = false;
		                           	}
		                       	}
		                       	
		                       } else {
		                           if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
		                               sameDay = false;
		                           } else {
		                           	// если сегодняшний месяц меньше
		                               if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
		                               	sameDay = true;
		                               	
		                                      // если текущий месяц больше на 1
		                                   if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
		                                   	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
		                                   		nRowYellow13 = true;
		                                   	} else {
		                                   		nRowYellow13 = false;
		                                   	}
		                                   	
		                                   	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                   		nRowRed13 = true;
		                                   	} else {
		                                   		nRowRed13 = false;
		                                   	}
		                               	}                                                                     
		                               } else {
		
		                                   if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
		                                       sameDay = false;  
		                                   } else {          
		                                   	// если месяцы равны
		                                       if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
		                                           sameDay = true;
		                                           nRowYellow13 = true;
		                                           
		                                        // если месяцы равны и текущий день <= 7
		                                       	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
		                                       		nRowRed13 = true;
		                                       	} else {
		                                       		nRowRed13 = false;
		                                       	}
		                                       } else {
		                                           sameDay = false;  	                                                                                        
		                                       }
		                                   }
		                               }
		                           }
		                       }                                                              
		                       sD13 = String.valueOf(sameDay);	
		                   }    
		                   break;
				        }
				      }
				    }
			        for (Cell cell2: row) {
			        	if (row.getRowNum() > 2) {
			        		
		        		if (sD10 == "true") {
	                    	if (data.size() == 0) {
	                    		dateGreen10.add(0);
	                    	} else {
	                    		dateGreen10.add(data.size()/21);
	                    	}
	                    	sD10 = "false";
	                    }
		        		if (sD13 == "true") {
	                    	if (data.size() == 0) {
	                    		dateGreen13.add(0);
	                    	} else {
	                    		dateGreen13.add(data.size()/21);
	                    	}
	                    	sD13 = "false";
	                    }
			        	switch (cell2.getCellType()) {
				        	case Cell.CELL_TYPE_STRING:  	
			                	
				        	if (cell2.getColumnIndex() == 0) {            
				        		if (!cell2.getStringCellValue().equals("")) {
				        			number = cell2.getStringCellValue();
				        			data.add(number);   
				        		} else {
				        			data.add(""); 
				        		}
		                    }		                    
		                    if (cell2.getColumnIndex() == 1) {
		                    	if (cell2.getStringCellValue().length() == 0) {                        		
		                    		data.add("");
		                    	} else {
		                    		data.add(cell2.getStringCellValue());
		                    		modelBrand.add(cell2.getStringCellValue());
		                    	}	
		                    }		                    
		                    if (cell2.getColumnIndex() == 2) {
		                    	if (cell2.getStringCellValue().equals("")) {                        		
		                    		data.add("");
		                    	} else {
		                    		data.add(cell2.getStringCellValue());		                    		
		                    	}
		                    }		                    
		                    if (cell2.getColumnIndex() == 3) {                                
		                    	name = cell2.getStringCellValue();
		                    	data.add(name);   
		                    }		                    
		                    if (cell2.getColumnIndex() >= 4 && cell2.getColumnIndex() <= 9) {
		                    	if (cell2.getColumnIndex() == 4) {
			                    	if (cell2.getStringCellValue().length() == 0) {
			                			data.add("");
			                		} else {
			                			data.add(cell2.getStringCellValue());
			                			listYears.add(cell2.getStringCellValue());	                		
			                			}	
			                    } else {
			                    	data.add(cell2.getStringCellValue());
			                    }	                    	                    	
		                    }  		                    
		                    if (cell2.getColumnIndex() == 10) {
		                    	data.add(cell2.getStringCellValue());                                                       		                           	 
		                    }		                    
		                    if (cell2.getColumnIndex() >= 11 && cell2.getColumnIndex() <= 13) {   
		                    	if (cell2.getColumnIndex() == 13) {
		                    		data.add(cell2.getStringCellValue());		                    		
		                    	} else {
		                    		data.add(cell2.getStringCellValue());
		                    	}	
		                    }		
		                    if (cell2.getColumnIndex() == 14) {                                   	
		                    	if (!cell2.getStringCellValue().equals("")) {		                    		
		                    	}	                    	
		                    	data.add(cell2.getStringCellValue());
		                    }		                    
		                    if (cell2.getColumnIndex() == 15) {
		                    	if (!cell2.getStringCellValue().equals("")) {
		                    		formsOwnership.add(cell2.getStringCellValue());		                    		
		                    	}	                    	
		                    	data.add(cell2.getStringCellValue());
		                    }	                    
		                    if (cell2.getColumnIndex() == 16) {                   	
		                    	data.add(cell2.getStringCellValue());
		                    	if (!cell2.getStringCellValue().equals("")) {
		                    		equipmentOwners.add(cell2.getStringCellValue());
		                    	}
		                    }
		                    if (cell2.getColumnIndex() == 17) {                   	
		                    	data.add(cell2.getStringCellValue());
		                    	if (!cell2.getStringCellValue().equals("")) {
		                    		allLocations.add(cell2.getStringCellValue());
		                    	}
		                    }
		                    if (cell2.getColumnIndex() == 18) {                   	
		                    	data.add(cell2.getStringCellValue());
		                    }
		                    if (cell2.getColumnIndex() == 19) {                                   	                               	
		                    	data.add(cell2.getStringCellValue());
		                    	data.add(cell2.getRowIndex()); // добавляем в список номер строки
		                    } 
		                    break;
				        default:
		                	if (cell2.getColumnIndex() == 0) {                                	
		                   		int x = (int) cell2.getNumericCellValue();                                   		
		                   		number = String.valueOf(x);                                		
		                   		data.add(number);	
		                    }
		                	
		                	if (cell2.getColumnIndex() == 1) {
		                		if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }
		                	if (cell2.getColumnIndex() == 2) {
		                		if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                			modelBrand.add(cell2.getStringCellValue());
		                		}	
		                    }
		                	if (cell2.getColumnIndex() == 3) {
		                		if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }
		                	if (cell2.getColumnIndex() == 4) {
		                		if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }
		                	if (cell2.getColumnIndex() == 5) {
		                		if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }		                    
		                    if (cell2.getColumnIndex() == 6) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                			listYears.add(String.valueOf((int) cell2.getNumericCellValue()));	                		
		                			}	
		                    }		                    
		                    if (cell2.getColumnIndex() == 7) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }		                    
		                    if (cell2.getColumnIndex() == 8) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }		                                                    
		                    if ( cell2.getColumnIndex() == 10 | cell2.getColumnIndex() == 13 ) {                                    	
		                    	if (cell2.getDateCellValue() == null) {
		                    		data.add("");
		                    	} else {
		                    		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
		                            data.add(ft.format(cell2.getDateCellValue()));
		                    	}
		                    }
		                    if (cell2.getColumnIndex() == 9) {
		                		
		                    	if (cell2.getDateCellValue() == null) {
		                    		data.add("");
		                    	} else {
		                    		data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                    	} 
		                    }
		                    if (cell2.getColumnIndex() == 11) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 12) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 14) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                			equipmentOwners.add(cell2.getStringCellValue());
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 15) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                			allLocations.add(cell2.getStringCellValue());
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 16) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 17) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 18) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
		                		}	
		                    }
		                    if (cell2.getColumnIndex() == 19) {
		                    	if (cell2.getDateCellValue() == null) {
		                			data.add("");
		                		} else {
		                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                		}	
				        		data.add(cell2.getRowIndex()); // добавляем в список номер строки
				        	}	
		                    break;
			        	}		        	
			        }
		        }   		        
			    if (nRowYellow10 == true && nRowRed10 == false) {	    			
	                if (data.size() == 0) {
	                	dateYellow10.add(0);
	                } else {
	                	dateYellow10.add((data.size()-21)/21);
	                }                   			
	    			nRowYellow10 = false;
	    		} 	                
	        	if (nRowRed10 == true) {          	        		
			      	if (data.size() == 0) {
			      		dateRed10.add(0);
			      	} else {
			      		dateRed10.add((data.size()-21)/21);
			      	}						      	
	    			nRowRed10 = false;
	    		}	        	
	        	if (nRowYellow13 == true && nRowRed13 == false) {	    			
	                if (data.size() == 0) {
	                	dateYellow13.add(0);
	                } else {
	                	dateYellow13.add((data.size()-21)/21);
	                }                   			
	    			nRowYellow13 = false;
	    		} 	                
	        	if (nRowRed13 == true) {          	        		
			      	if (data.size() == 0) {
			      		dateRed13.add(0);
			      	} else {
			      		dateRed13.add((data.size()-21)/21);
			      	}						      	
	    			nRowRed13 = false;
	    		} 
		    }            
            Set set3 = new HashSet(modelBrand);
            modelBrand.clear();
            modelBrand = new ArrayList(set3); 
            
            int cl = 21;
            int rw = data.size() / cl;
            int j = 0;
            int k = 0;
            String str[][] = new String[rw][cl];
            
            for (Object someString2 : data) {          	
                if (someString2 == null) {
                    str[j][k] = " ";
                    k++;
                } else {               	
                    if (j < rw) {                  	
                        if (k < cl) {
                            str[j][k] = someString2.toString();
                            k++;
                        } else {
                            k = 0;
                            j++;
                            str[j][k] = someString2.toString();
                            k++;
                        }
                    }
                }
            }
            
            Object[][] dt = new Object[rw][cl];
            
            for (int i = 0; i < rw; i++) {          	
                for (int j2 = 0; j2 < cl; j2++) {
                    dt[i][j2] = str[i][j2];
                }
            }
            
		    TableModel dm = new DefaultTableModel( dt, mainHeaders );
		    
            JTable table = new JTable(dm) { };            
            workbook.close();
            
            table.changeSelection(0, 0, false, false);
            JScrollPane scrollPane = new JScrollPane( table );
            getContentPane().add(scrollPane);
                       
            JPanel panel = new JPanel(new BorderLayout(2, 2));
            JPanel panelF1 = new JPanel(new BorderLayout(2, 2));
            JButton buttonUnload = new JButton("Выгрузить в excel");               
            JButton buttonAdd = new JButton("Добавить");
            JPanel three = new JPanel(new GridLayout(0, 5, 5, 5));
            
            // three.add(new Label(" "));
            three.add(new Label(" "));
            three.add(buttonUnload).setFont(bigFontTR); 
            three.add(new Label(" "));                            
            three.add(buttonAdd);
            three.add(new Label(" "));           
            
            panelF1.setPreferredSize(new Dimension(0, 110));
            
            three.setPreferredSize(new Dimension(0, hThree));
            three.setMaximumSize(new Dimension(0, hThree));
            three.setMinimumSize(new Dimension(0, 10));
            
            setLayout(new GridLayout(3, 1, 100, 100));

            JPanel panelFF2 = new JPanel(new GridBagLayout());
            
            GridBagConstraints c2 = new GridBagConstraints();
            c2.fill = GridBagConstraints.VERTICAL;
            c2.gridx = 1;
            c2.gridy = 0;            
            c2.weightx = 0.95;
            c2.weighty = 2;
            c2.fill = GridBagConstraints.BOTH;
            panelFF2.add(new JScrollPane(table), c2);
            
            c2.fill = GridBagConstraints.VERTICAL;
            c2.gridx = 1;
            c2.gridy = 1;            
            c2.weightx = 1;
            c2.weighty = 0.1;
            c2.fill = GridBagConstraints.BOTH;           
            panelFF2.add(three, c2);

            setLayout(new GridLayout(2, 1, 10, 10));
            panel.add(panelF1, BorderLayout.NORTH);
            panel.add(panelFF2, BorderLayout.CENTER);
        		
            frame.add(new JScrollPane(panel));
            frame.pack();
            frame.getRootPane().setDefaultButton(buttonSearch);
            frame.setVisible(true);        
            
            copyData = new ArrayList<String>(data);
                        			
			buttonAdd.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
            		try {
						new AddDataCar().inputValues();
					} catch (IOException e1) {
						e1.printStackTrace();
					}               	
                }
            });
							           	
        	frame.getContentPane().removeAll();
            
            data.clear();
            data = new ArrayList();    
            
            JTable table1 = new JTable(dm) {
            	// объединяем ячейки
                protected JTableHeader createDefaultTableHeader() {
                    return new GroupableTableHeader(columnModel); 
                }                
                // запрет на редактирование ячеек в таблице
                private static final long serialVersionUID = 1L;                
                // кнопку редактирования изменять можно
                public boolean isCellEditable(int row, int column) {                
                	if (column != 20) {
                        return false;   
                	} else {
                		return true; 
                	}               
                 };      
                 // закрашивание строки заголовка
		        public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
		        	
		        	DefaultTableCellRenderer rightRenderer = new DefaultTableCellRenderer();
	                rightRenderer.setHorizontalAlignment(SwingConstants.CENTER);
	                getColumnModel().getColumn(0).setCellRenderer(rightRenderer);
	                
	                Component c = super.prepareRenderer(renderer, row, column);	                                  	
	                c.setBackground(Color.white);
             	
	                if (column == 10) {
	                	 c.setBackground(new java.awt.Color(234, 234, 234));
	                }				                				                
	                for (int i = 0; i < dateGreen10.size(); i++) {
	                	if (row == ((int) dateGreen10.get(i)) && column == 10) {
	                		c.setBackground(Color.white);
	                	}
	                }	   	                				                
	                for (int i = 0; i < dateYellow10.size(); i++) {
	                	if (row == ((int) dateYellow10.get(i)) && column == 10) {
	                		c.setBackground(new java.awt.Color(247, 239, 162));
	                	}
	                }  				                
	                for (int i = 0; i < dateRed10.size(); i++) {
	                	if (row == ((int) dateRed10.get(i)) && column == 10) {
	                		c.setBackground(new java.awt.Color(213, 92, 95));
	                	}
	                }
	                
	                if (column == 13) {
	                	 c.setBackground(new java.awt.Color(234, 234, 234));
	                }				                				                
	                for (int i = 0; i < dateGreen13.size(); i++) {
	                	if (row == ((int) dateGreen13.get(i)) && column == 13) {
	                		c.setBackground(Color.white);
	                	}
	                }	   	                				                
	                for (int i = 0; i < dateYellow13.size(); i++) {
	                	if (row == ((int) dateYellow13.get(i)) && column == 13) {
	                		c.setBackground(new java.awt.Color(247, 239, 162));
	                	}
	                }  				                
	                for (int i = 0; i < dateRed13.size(); i++) {
	                	if (row == ((int) dateRed13.get(i)) && column == 13) {
	                		c.setBackground(new java.awt.Color(213, 92, 95));
	                	}
	                }
	                return c;	               	                
	            }
            };
                       
            table1.changeSelection(0, 0, false, false);
            JScrollPane scrollPane1 = new JScrollPane(table1);
            getContentPane().add(scrollPane1);
            
            // кнопка "редактировать"    
            table1.getColumn(" ").setCellRenderer(new ButtonRendererCars(frame));
            table1.getColumn(" ").setCellEditor(new ButtonEditorCars(new JCheckBox(), frame));
            table1.getColumnModel().getColumn(20).setPreferredWidth(30);
            // кнопка "редактировать"               
        	
        	// горизонтальная прокрутка заголовков
        	table1.getTableHeader().setPreferredSize(new Dimension(10000, 120));        	
        	table1.setRowHeight(25);
            
        	for (int i = 0; i <= mainHeaders.length - 2; i++) {
		        switch (i) {
		        	case 0:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(35);
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 2:
		        	case 3:
		        	case 17:
		        	case 18:
		        	case 19:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(130); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 5:
		        	case 6:
		        	case 7:
		        	case 9:
		        	case 12:
		        	case 14:
		        	case 15:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(110); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 4:
		        	case 10:
		        	case 13:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(70); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 1:
		        	case 16:
		        	case 11:
		        	case 8:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(160); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        }
        	}
        	
            JPanel panel12 = new JPanel(new BorderLayout(10, 10));
            JPanel panel01 = new JPanel(new BorderLayout(0, 25));
            
            table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);                      
            table1.setPreferredScrollableViewportSize(tableStart.getPreferredSize());           
            panel01.setPreferredSize(new Dimension(0, 267));
            
            three.setPreferredSize(new Dimension(0, hThree));
            three.setMaximumSize(new Dimension(0, hThree));
            three.setMinimumSize(new Dimension(0, 10));
            
            JPanel panelL = new JPanel(new GridLayout(5, 2, 10, 10));             
            Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);                    
            
            Set setKC = new HashSet( kindControl );
            kindControl.clear();
            kindControl = new ArrayList(setKC);        
            
            panelL.add(new JLabel("Марка, модель:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceModel().outputPanel(modelBrand, "Все марки и модели")).setFont(bigFontTR);             
    	    panelL.add(new JLabel("" , SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(new JLabel("", SwingConstants.RIGHT)).setFont(bigFontTR);                  
      	    
       	    Set setTO = new HashSet(formsOwnership);
            formsOwnership.clear();
            formsOwnership = new ArrayList(setTO);
            
            panelL.add(new JLabel("Форма собственности:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceFormOwnership().outputPanel(formsOwnership, "Все формы собственности")).setFont(bigFontTR);            
    	    panelL.add(new JLabel("" , SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(new JLabel("", SwingConstants.RIGHT)).setFont(bigFontTR);    	    
            
            Set setEO = new HashSet(equipmentOwners);
            equipmentOwners.clear();
            equipmentOwners = new ArrayList(setEO);
            
            panelL.add(new JLabel("Владелец оборудования:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceEquipmentOwners().outputPanel(equipmentOwners, "Все владельцы оборудования")).setFont(bigFontTR);           
            panelL.add(new JLabel("" , SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(new JLabel("")).setFont(bigFontTR);
    	    
            Set setL = new HashSet(allLocations);
            allLocations.clear();
            allLocations = new ArrayList(setL);  
            panelL.add(new JLabel("Местонахождение:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceLocation().outputPanel(allLocations, "Все местонахождения")).setFont(bigFontTR);          
            panelL.add(new JLabel("")).setFont(bigFontTR);
            panelL.add(new JLabel("")).setFont(bigFontTR);

            panelL.add(new JLabel("Год выпуска:", SwingConstants.RIGHT)).setFont(bigFontTR);
            yearRelease.setMaximumSize(new Dimension(10, 14));
            yearRelease.setMinimumSize(new Dimension(10, 14));
            panelL.add(yearRelease);
    	    
       	    JCheckBox field3 = new JCheckBox("", false);
       	    new WorkingCheckBox().scaleCheckBoxIcon(field3, 25);
    	    panelL.add(new JLabel("Техническое состояние исправно:", SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(field3).setFont(bigFontTR);
    	    
            TableColumnModel cm1 = table1.getColumnModel();
            ColumnGroup g_name1 = new ColumnGroup("КАСКО");
            g_name1.add(cm1.getColumn(8));
            g_name1.add(cm1.getColumn(9));
            g_name1.add(cm1.getColumn(10));
            
            ColumnGroup g_name2 = new ColumnGroup("ОСАГО");
            g_name2.add(cm1.getColumn(11));
            g_name2.add(cm1.getColumn(12));
            g_name2.add(cm1.getColumn(13));
            
            GroupableTableHeader th1 = (GroupableTableHeader) table1.getTableHeader();
            th1.addColumnGroup( g_name1 );
            th1.addColumnGroup( g_name2 );
        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));
            
            JPanel panelFF22 = new JPanel(new GridBagLayout());
            GridBagConstraints c22 = new GridBagConstraints();

            c22.fill = GridBagConstraints.VERTICAL;
            c22.gridx = 1;
            c22.gridy = 0;            
            c22.weightx = 0.95;
            c22.weighty = 2;
            c22.fill = GridBagConstraints.BOTH;
            panelFF22.add(new JScrollPane(table1), c22);
            
            c22.fill = GridBagConstraints.VERTICAL;
            c22.gridx = 1;
            c22.gridy = 1;          
            c22.weightx = 1;
            c22.weighty = 0.1;
            c22.fill = GridBagConstraints.BOTH;         
            panelFF22.add(three, c22);

            JButton buttonDC = new JButton("<html><center>" + "Приборы и<br> расходники"  + "<center><html>");
            buttonDC.setBackground(Color.LIGHT_GRAY);          
            
            JButton buttonCars = new JButton("Автомобили");
            buttonCars.setForeground(Color.WHITE);
            buttonCars.setBackground(Color.LIGHT_GRAY);
            
            JButton buttonTechnology = new JButton("Орг. техника");
            buttonTechnology.setBackground(Color.LIGHT_GRAY);
                                   
            JPanel buttonsPanels = new JPanel(new GridLayout(3, 1, 40, 40));         
            buttonsPanels.add( buttonDC );
            buttonsPanels.add( buttonCars );
            buttonsPanels.add( buttonTechnology );
            
            JPanel panelsButtons = new JPanel(new GridBagLayout());
            GridBagConstraints cPB = new GridBagConstraints();

            cPB.fill = GridBagConstraints.HORIZONTAL;
            cPB.gridx = 0;
            cPB.gridy = 1;            
            cPB.weightx = 1;
            cPB.ipady = -18;
            cPB.insets = new Insets(5, 0, 0, 0);
            cPB.fill = GridBagConstraints.BOTH;
            
            panelsButtons.add(panelL, cPB);
                        
            cPB.fill = GridBagConstraints.HORIZONTAL;
            cPB.gridx = 1;
            cPB.gridy = 1;           
            cPB.weightx = 0.01;
            cPB.ipadx = 10;
            // вверх, вниз, влево, вправо
            cPB.insets = new Insets(15, 0, 2, 8);
            cPB.fill = GridBagConstraints.BOTH;    
            
            panelsButtons.add(buttonsPanels, cPB);
            
            setLayout(new GridLayout(2, 1, 1, 1));
            panel01.add(panelsButtons, BorderLayout.NORTH);
            panel01.add(panelB, BorderLayout.CENTER);
            
            setLayout(new GridLayout(2, 1, 1, 1));
            panel12.add(panel01, BorderLayout.NORTH);
            panel12.add(panelFF22, BorderLayout.CENTER);
            
            frame.getContentPane().setLayout(new BorderLayout());                       
            frame.add(panel12);
            frame.pack();        
            frame.getRootPane().setDefaultButton(buttonSearch);
            frame.setVisible(true);
            frame.revalidate();
            frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
            
            cH = panelL.getHeight() - 190;
            cW = 30;
            
            buttonTechnology.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                	yearRelease.setText("");
                    yearRelease = new JTextField(14);
    			    jointUpload.clear();
    			    copyData.clear();
                    frame.setPreferredSize(frame.getSize());
                    
                	try {
    					new Technics().start(frame, three.getHeight(), table1);
    				} catch (FileNotFoundException e1) {
    					e1.printStackTrace();
    				} catch (IOException e1) {
    					e1.printStackTrace();
    				} catch (ParseException e1) {
    					e1.printStackTrace();
    				}  
                }
            });  
            
            buttonDC.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                	yearRelease.setText("");
                    yearRelease = new JTextField(14);
    			    jointUpload.clear();
    			    copyData.clear();
                    frame.setPreferredSize(frame.getSize());
                    
                	try {
    					new DeviceAndConsumables().start(frame, 1);
    				} catch (FileNotFoundException e1) {
    					e1.printStackTrace();
    				} catch (IOException e1) {
    					e1.printStackTrace();
    				} catch (ParseException e1) {
    					e1.printStackTrace();
    				}  
                }
            });  
            
        buttonStart.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	yearRelease.setText("");
                yearRelease = new JTextField(15);  
			    jointUpload.clear();
			    copyData.clear();
                frame.setPreferredSize(frame.getSize());
                dateGreen10.clear(); 
			    dateYellow10.clear(); 
			    dateRed10.clear(); 
			    dateGreen13.clear(); 
			    dateYellow13.clear(); 
			    dateRed13.clear();
			    
            	try {
					start(frame, three.getHeight(), table1);
				} catch (FileNotFoundException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				} catch (ParseException e1) {
					e1.printStackTrace();
				}  
            }
        });               
        
        Set setYears = new HashSet(listYears);
        listYears.clear();
        listYears = new ArrayList(setYears);
                  		
		new AutoSuggestor(yearRelease, frame, null, Color.WHITE.brighter(), Color.DARK_GRAY, Color.RED, 0.8f, cH, cW) {		 	
			boolean wordTyped(String typedWord)  {	     	             		
				setDictionary(listYears);
		 		return super.wordTyped(typedWord);
		    }                                       
		};		
		
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		buttonSearch.addActionListener(new ActionListener() {			
			@Override
			public void actionPerformed(ActionEvent e) {
				
			    ArrayList data = new ArrayList();		
			    data.clear();
			    dateGreen10.clear(); 
			    dateYellow10.clear(); 
			    dateRed10.clear(); 
			    dateGreen13.clear(); 
			    dateYellow13.clear(); 
			    dateRed13.clear();
			    
			    ArrayList selectedModel = new ChoiceModel().selectedModel();
			    ArrayList selectedFormsOwnership = new ChoiceFormOwnership().selectedHeaders();
			    ArrayList selectedEquipmentOwners = new ChoiceEquipmentOwners().selectedHeaders();
			    ArrayList selectedLocations = new ChoiceLocation().selectedLocations();			    
			    ArrayList whiteСell = new ArrayList();
			    
				try {			
				    InputStream inputStream = new FileInputStream("Reestr.xls");
				    Workbook workbook = new HSSFWorkbook(inputStream);
				    Sheet currentSheet = workbook.getSheetAt(1);
				    Iterator<Row> rowIterator = currentSheet.iterator();
				    rowIterator.next(); 
				    rowIterator.next();
				    String number = "";
				    String name = "";
				    String c2String = "";
				    boolean h = false;                        						    
				    String sD10 = null;
				    String sD13 = null;
				    int currentRow = 0;
				    
				    for (Row row: currentSheet) {
			            boolean b = false;
			            boolean b2 = false; 
	                    boolean column2 = false;	
	                    boolean column15 = false; 
	                    boolean column16 = false; 
	                    boolean column17 = false; 
	                    boolean column14 = false;
			            boolean nRowYellow = false;   
			            boolean nRowRed = false; 
			            boolean nRowRedWithoutYellow = false; 
				    	boolean nRowYellow10 = false;   
		                boolean nRowRed10 = false; 
		                boolean nRowYellow13 = false;   
		                boolean nRowRed13 = false;
		                
				        for (Cell cell: row) {				            
					    	if (row.getRowNum() > 1) {
						    	int columnIndex = cell.getColumnIndex();
						    	
						        switch (cell.getCellTypeEnum()) {
							        case STRING:			        				        
							        	
							        	if (cell.getColumnIndex() == 0) {				                        		
			                        		number = cell.getStringCellValue();			                        	
				                        }
				                        
				                        if (cell.getColumnIndex() == 1) {
				                        	c2String = cell.getStringCellValue();

				                        	if (selectedModel.get(0).equals("Все марки и модели")) {
				                        		column2 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedModel.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedModel.get(i))) {				                        		
						                        		column2 = true;
						                        	} 
					                        	}
				                        	}	
				                        }
				                        
				                        if (cell.getColumnIndex() == 2) {							                        	
				                        	name = cell.getStringCellValue();			                        	
				                        }		
				                        
				                        if (cell.getColumnIndex() == 4) {					                        	
				                        	String str = cell.getStringCellValue();					                        
					                        int indexM = str.indexOf(yearRelease.getText());				                        	
				                        	if (indexM != -1) {
				                        		b = true;
				                        	}				                        				                        				                        	
				                        	if (str.replaceAll("\\s","").indexOf(yearRelease.getText()) != -1) {
				                        		b = true;
				                        	}
				                        }	

				                        if ( cell.getColumnIndex() == 14 ) {				                        	
			                        		
				                        	if (cell.getStringCellValue().equals("Исправен") | cell.getStringCellValue().equals("Исправен.") | 
				                        			cell.getStringCellValue().equals("Исправно") | cell.getStringCellValue().equals("испр")) {			                        		
				                        		column14 = true;
				                        	}
				                        }
				                        
				                        if ( cell.getColumnIndex() == 10 ) {
				                        	boolean sameDay = false; 
				                        	Date dateNow = new Date();
				                        	Calendar cal1 = Calendar.getInstance();
				                            Calendar cal2 = Calendar.getInstance();
				                            	
				                        	String regex = "(\\d{2}.\\d{2}.\\d{4})";
				                    		Matcher m = Pattern.compile(regex).matcher(cell.getStringCellValue());
				                    		
				                    		if (m.find()) {
				                    			Date docDate = null;
				                    			
												try {
													docDate = new SimpleDateFormat("dd.MM.yyyy").parse(m.group(1));
												} catch (ParseException e1) {
													e1.printStackTrace();
												}
				                                cal1.setTime(dateNow);
				                                cal2.setTime(docDate);
				                    		} else {
				                    			sameDay = false;
				                    			sD10 = String.valueOf(sameDay);
				                    			break;
				                    		}
				                            
				                            // анализ 10 столбца
				                            if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
				                            	sameDay = true; 
				                            	
				                            	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
				                            		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
				                                		nRowYellow10 = true;
				                                	} else {
				                                		nRowYellow10 = false;
				                                	}
				                                	
				                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                		nRowRed10 = true;
				                                	} else {
				                                		nRowRed10 = false;
				                                	}
				                            	}
				                            	
				                            } else {
				                                if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
				                                    sameDay = false;
				                                } else {
				                                	// если сегодняшний месяц меньше
				                                    if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
				                                    	sameDay = true;
				                                    	
				                                           // если текущий месяц больше на 1
				                                        if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
				                                        	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
				                                        		nRowYellow10 = true;
				                                        	} else {
				                                        		nRowYellow10 = false;
				                                        	}
				                                        	
				                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                        		nRowRed10 = true;
				                                        	} else {
				                                        		nRowRed10 = false;
				                                        	}
				                                    	}                                                                     
				                                    } else {
				
				                                        if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
				                                            sameDay = false;  
				                                        } else {          
				                                        	// если месяцы равны
				                                            if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
				                                                sameDay = true;
				                                                nRowYellow10 = true;
				                                                
				                                             // если месяцы равны и текущий день <= 7
				                                            	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                            		nRowRed10 = true;
				                                            	} else {
				                                            		nRowRed10 = false;
				                                            	}
				                                            } else {
				                                                sameDay = false;  	                                                                                        
				                                            }
				                                        }
				                                    }
				                                }
				                            }                                                              
				                            sD10 = String.valueOf(sameDay);	
				                        }
				                        if (cell.getColumnIndex() == 13) {
				                        	boolean sameDay = false; 
				                        	Date dateNow = new Date();
				                        	Calendar cal1 = Calendar.getInstance();
				                            Calendar cal2 = Calendar.getInstance();
				                            	
				                        	String regex = "(\\d{2}.\\d{2}.\\d{4})";
				                    		Matcher m = Pattern.compile(regex).matcher(cell.getStringCellValue());
				                    		
				                    		if (m.find()) {
				                    			Date docDate = null;
				                    			
												try {
													docDate = new SimpleDateFormat("dd.MM.yyyy").parse(m.group(1));
												} catch (ParseException e1) {
													e1.printStackTrace();
												}
				                                cal1.setTime(dateNow);
				                                cal2.setTime(docDate);
				                    		} else {
				                    			sameDay = false;
				                    			sD13 = String.valueOf(sameDay);
				                    			break;
				                    		}
				                            
				                            // анализ 10 столбца
				                            if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
				                            	sameDay = true; 
				                            	
				                            	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
				                            		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
				                                		nRowYellow13 = true;
				                                	} else {
				                                		nRowYellow13 = false;
				                                	}
				                                	
				                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                		nRowRed13 = true;
				                                	} else {
				                                		nRowRed13 = false;
				                                	}
				                            	}
				                            	
				                            } else {
				                                if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
				                                    sameDay = false;
				                                } else {
				                                	// если сегодняшний месяц меньше
				                                    if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
				                                    	sameDay = true;
				                                    	
				                                           // если текущий месяц больше на 1
				                                        if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
				                                        	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
				                                        		nRowYellow13 = true;
				                                        	} else {
				                                        		nRowYellow13 = false;
				                                        	}
				                                        	
				                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                        		nRowRed13 = true;
				                                        	} else {
				                                        		nRowRed13 = false;
				                                        	}
				                                    	}                                                                     
				                                    } else {
				
				                                        if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
				                                            sameDay = false;  
				                                        } else {          
				                                        	// если месяцы равны
				                                            if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
				                                                sameDay = true;
				                                                nRowYellow13 = true;
				                                                
				                                             // если месяцы равны и текущий день <= 7
				                                            	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
				                                            		nRowRed13 = true;
				                                            	} else {
				                                            		nRowRed13 = false;
				                                            	}
				                                            } else {
				                                                sameDay = false;  	                                                                                        
				                                            }
				                                        }
				                                    }
				                                }
				                            }                                                              
				                            sD13 = String.valueOf(sameDay);	
				                        }
				                        if (cell.getColumnIndex() == 16) {  
				                        	
				                        	if (selectedEquipmentOwners.get(0).equals("Все владельцы оборудования")) {
				                        		column16 = true;
				                        	} else {
				                        		
					                        	for (int i = 0; i < selectedEquipmentOwners.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedEquipmentOwners.get(i))) {				                        		
						                        		column16 = true;
						                        	} 
					                        	}
				                        	}
				                        }
				                       	
				                        if (cell.getColumnIndex() == 15) {   
				                        	
				                        	if (selectedFormsOwnership.get(0).equals("Все формы собственности")) {
				                        		column15 = true;
				                        	} else {
					                        	for (int i = 0; i < selectedFormsOwnership.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedFormsOwnership.get(i))) {				                        		
						                        		column15 = true;
						                        		
						                        	} 
					                        	}
				                        	}
				                        	currentRow = cell.getRowIndex();
				                        }
				                        if (cell.getColumnIndex() == 17) {  
				                        	
				                        	if ( selectedLocations.get(0).equals("Все местонахождения") ) {
				                        		column17 = true;
				                        	} else {
				                        		
					                        	for (int i = 0; i < selectedLocations.size(); i++) {
						                        	if (cell.getStringCellValue().equals( selectedLocations.get(i) )) {				                        		
						                        		column17 = true;
						                        	} 
					                        	}
				                        	}
				                        }
				                        break;
							        default:               	   
							        	if (cell.getColumnIndex() == 0) {				                        		                                                		
			                        		int x = (int) cell.getNumericCellValue();                                   		
			                        		number = String.valueOf(x);			                        	
				                        }
				                        
				                        if (cell.getColumnIndex() == 1) {
				                        	c2String = cell.getStringCellValue();
				                        	
				                        	if (selectedModel.get(0).equals("Все марки и модели")) {
				                        		column2 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedModel.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedModel.get(i))) {				                        		
						                        		column2 = true;
						                        	} 
					                        	}
				                        	}	
				                        }
				                        
				                        if (cell.getColumnIndex() == 2) {					                        					                        	
				                        	name = cell.getStringCellValue();			                        	
				                        }		
				                        
				                        if (cell.getColumnIndex() == 4) {	
				                        	
				                        	String str;
				                        	if (cell.getCellTypeEnum() != NUMERIC) {   				                        		
				                        		str = cell.getStringCellValue();
				                        	} else {				                        		                                                		
				                        		str = String.valueOf(cell.getNumericCellValue());
				                        	}
					                        
					                        int indexM = str.indexOf(yearRelease.getText());
				                        	
				                        	if (indexM != -1) {
				                        		b = true;
				                        	}				                        	
				                        				                        	
				                        	if (str.replaceAll("\\s","").indexOf(yearRelease.getText()) != -1) {
				                        		b = true;
				                        	}
				                        }	
				                        
				                        if (cell.getColumnIndex() == 11) {
 	
				                        }
				                        if (cell.getColumnIndex() == 14) {				                        	
			                        		
				                        	if (cell.getStringCellValue().equals("Исправен") | cell.getStringCellValue().equals("Исправен.") | 
				                        			cell.getStringCellValue().equals("Исправно") | cell.getStringCellValue().equals("испр")) {			                        		
				                        		column14 = true;
				                        	}
				                        }
				                        
				                        if (cell.getColumnIndex() == 10) {
					                       	boolean sameDay = false; 
					                       	Date dateNow = new Date();
					                       	Calendar cal1 = Calendar.getInstance();
					                        Calendar cal2 = Calendar.getInstance();
					                       	sameDay = false; 
					                       	
					                       	if (cell.getDateCellValue() != null) {                                                                                       
					                               cal1.setTime(dateNow);
					                               cal2.setTime(cell.getDateCellValue());
					                       	} else {
					                       		break;
					                       	}
					                       	
					                        // анализ 10 столбца
					                       if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
					                       	sameDay = true; 
					                       	
					                       	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
					                       		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
					                           		nRowYellow10 = true;
					                           	} else {
					                           		nRowYellow10 = false;
					                           	}		                           	
					                           	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
					                           		nRowRed10 = true;
					                           	} else {
					                           		nRowRed10 = false;
					                           	}
					                       	}		                       	
					                       } else {
					                           if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
					                               sameDay = false;
					                           } else {
					                           	// если сегодняшний месяц меньше
					                               if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
					                               	sameDay = true;
					                               	
					                                      // если текущий месяц больше на 1
					                                   if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
					                                   	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
					                                   		nRowYellow10 = true;
					                                   	} else {
					                                   		nRowYellow10 = false;
					                                   	}
					                                   	
					                                   	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
					                                   		nRowRed10 = true;
					                                   	} else {
					                                   		nRowRed10 = false;
					                                   	}
					                               	}                                                                     
					                               } else {	
					                                   if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
					                                       sameDay = false;  
					                                   } else {          
					                                   	// если месяцы равны
					                                       if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
					                                           sameDay = true;
					                                           nRowYellow10 = true;
					                                           
					                                        // если месяцы равны и текущий день <= 7
					                                       	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
					                                       		nRowRed10 = true;
					                                       	} else {
					                                       		nRowRed10 = false;
					                                       	}
					                                       } else {
					                                           sameDay = false;  	                                                                                        
					                                       }
					                                   }
					                               }
					                           }
					                       }                                                              
					                       sD10 = String.valueOf(sameDay);	
					                   } 
				                	   if (cell.getColumnIndex() == 13) {
					                       	boolean sameDay = false; 
					                       	Date dateNow = new Date();
					                       	Calendar cal1 = Calendar.getInstance();
					                        Calendar cal2 = Calendar.getInstance();
					                       	sameDay = false; 
					                       	
					                       	if (cell.getDateCellValue() != null) {                                                                                       
					                               cal1.setTime(dateNow);
					                               cal2.setTime(cell.getDateCellValue());
					                       	} else {
					                       		break;
					                       	}
					                       	
					                        // анализ 10 столбца
					                       if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
					                       	sameDay = true; 
					                       	
					                       	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
					                       		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
					                           		nRowYellow13 = true;
					                           	} else {
					                           		nRowYellow13 = false;
					                           	}
					                           	
					                           	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
					                           		nRowRed13 = true;
					                           	} else {
					                           		nRowRed13 = false;
					                           	}
					                       	}
					                       	
					                       } else {
					                           if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
					                               sameDay = false;
					                           } else {
					                           	// если сегодняшний месяц меньше
					                               if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
					                               	sameDay = true;
					                               	
					                                      // если текущий месяц больше на 1
					                                   if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
					                                   	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
					                                   		nRowYellow13 = true;
					                                   	} else {
					                                   		nRowYellow13 = false;
					                                   	}
					                                   	
					                                   	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
					                                   		nRowRed13 = true;
					                                   	} else {
					                                   		nRowRed13 = false;
					                                   	}
					                               	}                                                                     
					                               } else {
					
					                                   if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
					                                       sameDay = false;  
					                                   } else {          
					                                   	// если месяцы равны
					                                       if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
					                                           sameDay = true;
					                                           nRowYellow13 = true;
					                                           
					                                        // если месяцы равны и текущий день <= 7
					                                       	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
					                                       		nRowRed13 = true;
					                                       	} else {
					                                       		nRowRed13 = false;
					                                       	}
					                                       } else {
					                                           sameDay = false;  	                                                                                        
					                                       }
					                                   }
					                               }
					                           }
					                       }                                                              
					                       sD13 = String.valueOf(sameDay);	
					                   }    
				                        
				                        if (cell.getColumnIndex() == 16) {  
				                        	
				                        	if (selectedEquipmentOwners.get(0).equals("Все владельцы оборудования")) {
				                        		column16 = true;
				                        	} else {
				                        		
					                        	for (int i = 0; i < selectedEquipmentOwners.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedEquipmentOwners.get(i))) {				                        		
						                        		column16 = true;
						                        	} 
					                        	}
				                        	}
				                        }
				                        
				                        if (cell.getColumnIndex() == 15) {   
				                        	
				                        	if (selectedFormsOwnership.get(0).equals("Все формы собственности")) {
				                        		column15 = true;
				                        	} else {
					                        	for (int i = 0; i < selectedFormsOwnership.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedFormsOwnership.get(i))) {				                        		
						                        		column15 = true;
						                        	} 
					                        	}
				                        	}
				                        	currentRow = cell.getRowIndex();
				                        }		
				                        if (cell.getColumnIndex() == 17) {  
				                        	
				                        	if (selectedLocations.get(0).equals("Все владельцы оборудования")) {
				                        		column17 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedLocations.size(); i++) {
						                        	if (cell.getStringCellValue().equals(selectedLocations.get(i))) {				                        		
						                        		column17 = true;
						                        	} 
					                        	}
				                        	}
				                        }
				                    break;
					        	}
					        }
			            }
			        	if (field3.isSelected() == false) {
                    		column14 = true;
                    	}
                    
                    	int nRow = 0;	
                    	
	                    if ( b == true && column2 == true && column15 == true && column16 == true && column17 == true && b == true 
	                    		&& column14 == true ) {
	                	
				        for (Cell cell2: row) {
				        	if ( row.getRowNum() > 2 ) {
				        		if (sD10 == "true") {
			                    	if (data.size() == 0) {
			                    		dateGreen10.add(0);
			                    	} else {
			                    		dateGreen10.add(data.size()/21);
			                    	}
			                    	sD10 = "false";
			                    }
				        		if (sD13 == "true") {
			                    	if (data.size() == 0) {
			                    		dateGreen13.add(0);
			                    	} else {
			                    		dateGreen13.add(data.size()/21);
			                    	}
			                    	sD13 = "false";
			                    }
				        		switch (cell2.getCellTypeEnum()) {
					        		case STRING:  				        	
							        	if (cell2.getColumnIndex() == 0) {            
							        		if (!cell2.getStringCellValue().equals("")) {
							        			number = cell2.getStringCellValue();
							        			data.add(number);   
							        		} else {
							        			data.add(""); 
							        		}
					                    }		                    
					                    if (cell2.getColumnIndex() == 1) {
					                    	if (cell2.getStringCellValue().length() == 0) {                        		
					                    		data.add("");
					                    	} else {
					                    		data.add(cell2.getStringCellValue());
					                    	}	
					                    }		                    
					                    if (cell2.getColumnIndex() == 2) {
					                    	if (cell2.getStringCellValue().equals("")) {                        		
					                    		data.add("");
					                    	} else {
					                    		data.add(cell2.getStringCellValue());
					                    	}
					                    }		                    
					                    if (cell2.getColumnIndex() == 3) {                                
					                    	name = cell2.getStringCellValue();
					                    	data.add(name);   
					                    }		                    
					                    if (cell2.getColumnIndex() >= 4 && cell2.getColumnIndex() <= 9) {
					                    	if (cell2.getColumnIndex() == 6) {
						                    	if (cell2.getStringCellValue().length() == 0) {
						                			data.add("");
						                		} else {
						                			data.add(cell2.getStringCellValue());                		
						                			}	
						                    } else {
						                    	data.add(cell2.getStringCellValue());
						                    }	                    	                    	
					                    }  		                    
				                    
					                    if (cell2.getColumnIndex() >= 10 && cell2.getColumnIndex() <= 13) {   
					                    	if (cell2.getColumnIndex() == 13) {
					                    		data.add(cell2.getStringCellValue());
					                    	} else {
					                    		data.add(cell2.getStringCellValue());
					                    	}	
					                    }		
					                    if (cell2.getColumnIndex() == 14) {       
					                    	data.add(cell2.getStringCellValue());
					                    }		                    
					                    if (cell2.getColumnIndex() == 15) {        	
					                    	data.add(cell2.getStringCellValue());
					                    }	                    
					                    if (cell2.getColumnIndex() == 16) {                   	
					                    	data.add(cell2.getStringCellValue());
					                    }	
					                    if (cell2.getColumnIndex() == 17) {                   	
					                    	data.add(cell2.getStringCellValue());
					                    }	
					                    if (cell2.getColumnIndex() == 18) {                   	
					                    	data.add(cell2.getStringCellValue());
					                    }	
					                    if (cell2.getColumnIndex() == 19) {                                   	                               	
					                    	data.add(cell2.getStringCellValue());
					                    	data.add(cell2.getRowIndex()); // добавляем в список номер строки
					                    } 
					                    break;
							        default:
					                	if (cell2.getColumnIndex() == 0) {                                	
					                   		int x = (int) cell2.getNumericCellValue();                                   		
					                   		number = String.valueOf(x);                                		
					                   		data.add(number);	
					                    }
					                	
					                	if (cell2.getColumnIndex() == 1) {
					                		if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                	if (cell2.getColumnIndex() == 2) {
					                		if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                	if (cell2.getColumnIndex() == 3) {
					                		if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                	if (cell2.getColumnIndex() == 4) {
					                		if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                	if (cell2.getColumnIndex() == 5) {
					                		if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }		                    
					                    if (cell2.getColumnIndex() == 6) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                			listYears.add(String.valueOf((int) cell2.getNumericCellValue()));	                		
					                			}	
					                    }		                    
					                    if (cell2.getColumnIndex() == 7) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }		                    
					                    if (cell2.getColumnIndex() == 8) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }		                                                    
					                    if (cell2.getColumnIndex() == 9) {                                    	
					                    	if (cell2.getDateCellValue() == null) {
					                    		data.add("");
					                    	} else {
					                    		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
					                            data.add(ft.format(cell2.getDateCellValue()));
					                    	}
					                    }
					                    if ( cell2.getColumnIndex() == 10 | cell2.getColumnIndex() == 13 ) {
					                		
					                    	if (cell2.getDateCellValue() == null) {
					                    		data.add("");
					                    	} else {
					                    		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
					                            data.add(ft.format(cell2.getDateCellValue())); 
					                    	} 
					                    }
					                    if (cell2.getColumnIndex() == 11) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 12) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 14) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 15) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 16) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 17) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 18) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));	                			
					                		}	
					                    }
					                    if (cell2.getColumnIndex() == 19) {
					                    	if (cell2.getDateCellValue() == null) {
					                			data.add("");
					                		} else {
					                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
					                		}	
							        		data.add(cell2.getRowIndex()); // добавляем в список номер строки
							        	}	
					                    break;
						        	}		        	
		        				}
				        	}   
	                    }
	                    if (nRowYellow10 == true && nRowRed10 == false) {	    			
	    	                if (data.size() == 0) {
	    	                	dateYellow10.add(0);
	    	                } else {
	    	                	dateYellow10.add((data.size()-21)/21);
	    	                }                   			
	    	    			nRowYellow10 = false;
	    	    		} 	                
	    	        	if (nRowRed10 == true) {          	        		
	    			      	if (data.size() == 0) {
	    			      		dateRed10.add(0);
	    			      	} else {
	    			      		dateRed10.add((data.size()-21)/21);
	    			      	}						      	
	    	    			nRowRed10 = false;
	    	    		}	        	
	    	        	if (nRowYellow13 == true && nRowRed13 == false) {	    			
	    	                if (data.size() == 0) {
	    	                	dateYellow13.add(0);
	    	                } else {
	    	                	dateYellow13.add((data.size()-21)/21);
	    	                }                   			
	    	    			nRowYellow13 = false;
	    	    		} 	                
	    	        	if (nRowRed13 == true) {          	        		
	    			      	if (data.size() == 0) {
	    			      		dateRed13.add(0);
	    			      	} else {
	    			      		dateRed13.add((data.size()-21)/21);
	    			      	}						      	
	    	    			nRowRed13 = false;
	    	    		}
			        }       
				    
					int cl = 21;
				    int rw = data.size() / cl;
				    int j = 0;
				    int k = 0;
				    String str[][] = new String[rw][cl];
				
				    for (Object someString2 : data) {
				
				        if (someString2 == null) {
				            str[j][k] = " ";
				            k++;
				        } else {				
				            if (j < rw) {					
				                if (k < cl) {
				                    str[j][k] = someString2.toString();
				                    k++;
				                } else {
				                    k = 0;
				                    j++;
				                    str[j][k] = someString2.toString();
				                    k++;
				                }
				            }
				        }
				    }
				
				    Object[][] dt1 = new Object[rw][cl];
				    int c13 = 0;
				    
				    for (int i = 0; i < rw; i++) {
				        for (int j2 = 0; j2 < cl; j2++) {
			        		if (j2 == cl-1 ) {                               			
			        			dt1[i][c13] = str[i][j2];	     
			            		c13 = 0;
			            	} else {
			            		dt1[i][c13] = str[i][j2];	     
			            		c13++;
			            	}	                             	
				        }                               
				    }
				    				    			    
				    DefaultTableModel dm1 = new DefaultTableModel();
				    dm1.setDataVector(dt1, mainHeaders);
		            
				    JTable table1 = new JTable(dm1) {
		            	// объединяем ячейки
		                protected JTableHeader createDefaultTableHeader() {
		                    return new GroupableTableHeader(columnModel); 
		                }                
		                // запрет на редактирование ячеек в таблице
		                private static final long serialVersionUID = 1L;                
		                // кнопку редактирования изменять можно
		                public boolean isCellEditable(int row, int column) {                
		                	if (column != 20) {
		                        return false;   
		                	} else {
		                		return true; 
		                	}               
		                 };      
		                 // закрашивание строки заголовка
				        public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
				        	
				        	DefaultTableCellRenderer rightRenderer = new DefaultTableCellRenderer();
			                rightRenderer.setHorizontalAlignment(SwingConstants.CENTER);
			                getColumnModel().getColumn(0).setCellRenderer(rightRenderer);
			                
			                Component c = super.prepareRenderer(renderer, row, column);	                                  	
			                c.setBackground(Color.white);
		             	
			                if (column == 10) {
			                	 c.setBackground(new java.awt.Color(234, 234, 234));
			                }				                				                
			                for (int i = 0; i < dateGreen10.size(); i++) {
			                	if (row == ((int) dateGreen10.get(i)) && column == 10) {
			                		c.setBackground(Color.white);
			                	}
			                }	   	                				                
			                for (int i = 0; i < dateYellow10.size(); i++) {
			                	if (row == ((int) dateYellow10.get(i)) && column == 10) {
			                		c.setBackground(new java.awt.Color(247, 239, 162));
			                	}
			                }  				                
			                for (int i = 0; i < dateRed10.size(); i++) {
			                	if (row == ((int) dateRed10.get(i)) && column == 10) {
			                		c.setBackground(new java.awt.Color(213, 92, 95));
			                	}
			                }
			                
			                if (column == 13) {
			                	 c.setBackground(new java.awt.Color(234, 234, 234));
			                }				                				                
			                for (int i = 0; i < dateGreen13.size(); i++) {
			                	if (row == ((int) dateGreen13.get(i)) && column == 13) {
			                		c.setBackground(Color.white);
			                	}
			                }	   	                				                
			                for (int i = 0; i < dateYellow13.size(); i++) {
			                	if (row == ((int) dateYellow13.get(i)) && column == 13) {
			                		c.setBackground(new java.awt.Color(247, 239, 162));
			                	}
			                }  				                
			                for (int i = 0; i < dateRed13.size(); i++) {
			                	if (row == ((int) dateRed13.get(i)) && column == 13) {
			                		c.setBackground(new java.awt.Color(213, 92, 95));
			                	}
			                }
			                return c;	               	                
			            }
		            };				    
				    
				    TableColumnModel cm1 = table1.getColumnModel();
		            ColumnGroup g_name1 = new ColumnGroup("КАСКО");
		            g_name1.add(cm1.getColumn(8));
		            g_name1.add(cm1.getColumn(9));
		            g_name1.add(cm1.getColumn(10));
		            
		            ColumnGroup g_name2 = new ColumnGroup("ОСАГО");
		            g_name2.add(cm1.getColumn(11));
		            g_name2.add(cm1.getColumn(12));
		            g_name2.add(cm1.getColumn(13));
		            
		            GroupableTableHeader th1 = (GroupableTableHeader) table1.getTableHeader();
		            th1.addColumnGroup( g_name1 );
		            th1.addColumnGroup( g_name2 );
		        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));
		        	
		        	// горизонтальная прокрутка заголовков
		        	table1.getTableHeader().setPreferredSize(new Dimension(10000, 120));
		        	
		            // кнопка "редактировать"    
		            table1.getColumn(" ").setCellRenderer(new ButtonRendererCars(frame));
		            table1.getColumn(" ").setCellEditor(new ButtonEditorCars(new JCheckBox(), frame));
		            table1.getColumnModel().getColumn(20).setPreferredWidth(30);
		            // кнопка "редактировать"
		        			        	
		        	// table1.setPreferredScrollableViewportSize(table.getPreferredSize());
		            table1.changeSelection(0, 0, false, false);
		            JScrollPane scrollPane1 = new JScrollPane( table1 );
		            getContentPane().add(scrollPane1);		            
		        	
		            table1.setRowHeight( 25 );
		            
		        	for (int i = 0; i <= mainHeaders.length - 2; i++) {
				        switch (i) {
				        	case 0:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(35);
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 2:
				        	case 3:
				        	case 17:
				        	case 18:
				        	case 19:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(130); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 5:
				        	case 6:
				        	case 7:
				        	case 9:
				        	case 12:
				        	case 14:
				        	case 15:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(110); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 4:
				        	case 10:
				        	case 13:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(70); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 1:
				        	case 16:
				        	case 11:
				        	case 8:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(160); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        }
		        	}		        	
		            
				    workbook.close();
				    
				    table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		            
				    JPanel panel = new JPanel(new BorderLayout(10, 10));
				    JPanel panelTop = new JPanel();
				    JPanel panelBt1 = new JPanel(new GridBagLayout());
				    JPanel panel2 = new JPanel(new BorderLayout(0, 0));
				    JPanel panel1 = new JPanel(new GridBagLayout());
				    GridBagConstraints c = new GridBagConstraints();
				    
		            three.setPreferredSize(new Dimension(0, hThree));
		            three.setMaximumSize(new Dimension(0, hThree));
		            three.setMinimumSize(new Dimension(0, 10));
		            
				    JPanel panelFF2 = new JPanel(new GridBagLayout());
				    
				    GridBagConstraints c2 = new GridBagConstraints();		
				    
				    c2.fill = GridBagConstraints.VERTICAL;
				    c2.gridx = 1;
				    c2.gridy = 0;				    
				    c2.weightx = 0.95;
				    c2.weighty = 2;
				    c2.fill = GridBagConstraints.BOTH;
				    panelFF2.add(new JScrollPane(table1), c2);
				   
				    c2.fill = GridBagConstraints.VERTICAL;
				    c2.gridx = 1;
				    c2.gridy = 1;				    
				    c2.weightx = 1;
				    c2.weighty = 0.1;
				    c2.fill = GridBagConstraints.BOTH;				    
				    panelFF2.add(three, c2);
				                                
				    setLayout(new GridLayout(2, 1, 1, 1));
				    panel.add(panel01, BorderLayout.NORTH);
				    panel.add(panelFF2, BorderLayout.CENTER);
				    
				    frame.setPreferredSize(frame.getSize());
				    frame.getContentPane().setLayout(new BorderLayout());
				    frame.getContentPane().removeAll();
				    frame.add(panel);
				    frame.pack();        
				    frame.getRootPane().setDefaultButton(buttonSearch);
				    frame.setVisible(true);
				    frame.revalidate();                
				    frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);				   
				    
		            copyData = new ArrayList<String>(data);
		            
				    if (fieldSub.isSelected() == true) {			    	
				    	for (int i = 0; i < copyData.size(); i++) {
				    		jointUpload.add(copyData.get(i));
				    	}
				    }
				    
				} catch (IOException ex1) {
					System.out.println("Error reading file");
			        ex1.printStackTrace();
			    }
			        
			}
		   
		});
		
		buttonUnload.addActionListener(new ActionListener() {
        
			@Override
            public void actionPerformed(ActionEvent e) {
                try {
                	if (fieldSub.isSelected() == true) {
                		new SelectingColumnsCar().selecting(jointUpload);
				    } else {
				    	jointUpload.clear();
				    	new SelectingColumnsCar().selecting(copyData);	        			
				    }                	        			        		
	        		
                } catch (FileNotFoundException e1) {
                    e1.printStackTrace();
                } catch (IOException e1) {
					
					e1.printStackTrace();
				}
            }
        }); 
		
        } catch (IOException ex1) {
                System.out.println("Error reading file");
                ex1.printStackTrace();
        } 
        catch (NullPointerException n) {
            System.out.println(n);             
        }
    }
}

//создание кнопки "редактировать"
class ButtonRendererCars extends JButton implements TableCellRenderer {
	
	JFrame frame;
	int i = 0;
	
	public ButtonRendererCars(JFrame frame) {
	 	this.frame = frame;	
	    setOpaque(true);
	}

	public Component getTableCellRendererComponent(JTable table, Object value,
         boolean isSelected, boolean hasFocus, int row, int column) {
 	
	     GridBagConstraints gbc = new GridBagConstraints();
	     gbc.gridwidth = GridBagConstraints.REMAINDER;
	     gbc.fill = GridBagConstraints.HORIZONTAL;
	
	     ImageIcon pencil = null;
	     pencil = new ImageIcon(new Cars().getClass().getClassLoader().getResource("pencil.png"));
	     Image image = pencil.getImage(); 
	     Image newimg = image.getScaledInstance(23, 23, java.awt.Image.SCALE_SMOOTH);
	     pencil = new ImageIcon(newimg);       
	     setBorderPainted(false);
	     setBorder(new LineBorder(Color.BLACK));
	
	     if ( !isSelected) {   	
	     	setForeground(table.getSelectionForeground());
	        setBackground(table.getSelectionBackground());
	     }
	     
	     setBackground(Color.white);
	     setIcon(pencil);
	     return this;
	}
}

//кнопка "редактировать"
class ButtonEditorCars extends DefaultCellEditor {
	
	public JButton button;
	String label = "";
	JFrame frame;
	public boolean isPushed;
	
	 public ButtonEditorCars(JCheckBox checkBox, JFrame frame) {
	     super(checkBox);
	     this.frame = frame;
	     button = new JButton(label);
	     button.setOpaque(true);
	     
	     button.addActionListener(new ActionListener() {
	         public void actionPerformed(ActionEvent e) {
	         	 fireEditingStopped();  
	         	 /*frame.dispose();
	          	 frame.setVisible(true);
	          	 frame.revalidate();*/
	         }
	     });
	 }
 
	 public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
	 	
		 /*frame.dispose();
	 	 frame.setVisible(true);
	 	 frame.revalidate();*/
	 	 label = "";
	     button = new JButton(label);
	     GridBagConstraints gbc = new GridBagConstraints();
	     gbc.gridwidth = GridBagConstraints.REMAINDER;
	     gbc.fill = GridBagConstraints.HORIZONTAL;
	
	     ImageIcon pencil = null;
	
	 	 /*frame.dispose();
	 	 frame.setVisible(true);
	 	 frame.revalidate();*/
	     pencil = new ImageIcon(new Cars().getClass().getClassLoader().getResource("pencil.png"));
	     Image image = pencil.getImage();
	     Image newimg = image.getScaledInstance(23, 23, java.awt.Image.SCALE_SMOOTH);
	     pencil = new ImageIcon(newimg);
	
	     button.setBorderPainted(false);
	     button.setBorder(new LineBorder(Color.BLACK));
	
	     if (isSelected) {
	         button.setForeground(table.getSelectionForeground());
	         button.setBackground(table.getSelectionBackground());
	     } else {
	     	 /*frame.dispose();
	     	 frame.setVisible(true);
	     	 frame.revalidate();*/
	     	 button.setForeground(table.getSelectionForeground());
	         button.setBackground(table.getSelectionBackground());
	     }
	     
	     button.setBackground(Color.white);
	     button.setIcon(pencil);
	     label = (value == null) ? "" : value.toString();	
	     isPushed = true;
	     TableModel tm = table.getModel();
	     String[] inputValue = new String[21];
	     
	     for (int i = 0; i < inputValue.length; i++) {
	         inputValue[i] = (String) tm.getValueAt(row, i);
	     }	    
    	 new EditDataCar().windowDataChange( inputValue );
		     
	     isPushed = true;
	     return button;
	 }
	 public Object getCellEditorValue() {		 
		 label = "";
	     isPushed = false;
	     // frame.dispose();
	 	 // frame.setVisible(true);
	 	 // frame.revalidate();
	     return label;
	 }
	 public boolean stopCellEditing() {		 
	     isPushed = true;
	     return super.stopCellEditing();
	 }
}