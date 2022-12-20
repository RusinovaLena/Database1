package net.codejava;

import net.codejava.SelectingColumnsDC.CustomerItem;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFRow;
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

public class DeviceAndConsumables extends JFrame  { // приборы и расходники 
	
	final static Object[] mainHeaders = {"№", "<html> <center>Вид<br> контроля", "<html> <center>Назначение<br>(область применения)",
		    "<html> <center>Наименование<br> прибора", "<html> <center>Тип, марка,<br> модель", 
		    "<html><center>Производитель, страна<br> производства, марка,<br> модель, основные<br> технические<br> характеристики<html>", 
            "Зав.№", "Количество", "<html> <center>Год<br>выпуска", "<html> <center>Дата поверки<br>(калибровки)", 
            "<html> <center>Дата окончания <br>поверки (калибровки)", "Документы", "<html> <center>Техническое<br> состояние",
            "<html> <center>Указание в поверке на<br>принадлежность к <br> организации", "<html> <center>Форма<br> собственности", 
            "<html> <center>Владелец<br> оборудавания", "Местонахождение", "Примечание", " "};
            // "<html> <center>Владелец<br> оборудавания", "Местонахождение", "Примечание"};
	
    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
    
    JTextField searchName = new JTextField(14);
    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
    
    ArrayList<Integer> dateGreen = new ArrayList<Integer>();
    ArrayList<Integer> dateYellow = new ArrayList<Integer>();
    ArrayList<Integer> dateRed = new ArrayList<Integer>();
    
    static JFrame frame = new JFrame();
    static ArrayList <String>copyData = new ArrayList<String>();
    static ArrayList jointUpload = new ArrayList();
    
    ArrayList appointment = new ArrayList();
    ArrayList arrayMonthToFinish = new ArrayList();
    ArrayList monthToFinish = new ArrayList();
    ArrayList weekToFinish = new ArrayList();
    ArrayList arrayWeekToFinish = new ArrayList();
    ArrayList kindControl = new ArrayList();
    ArrayList formsOwnership = new ArrayList();
    ArrayList equipmentOwners = new ArrayList();
    ArrayList allLocations = new ArrayList();
    ArrayList allNames = new ArrayList();
    
    static int cH = 0;
    static int cW = 0;
    public void start(JFrame frame, int startSize) throws IOException, ParseException {
    	
    	// установка другой иконки для JFrame
    	/*ImageIcon liderIcon = new ImageIcon(new DeviceAndConsumables().getClass().getClassLoader().getResource(".png"));
        Image image = liderIcon.getImage();
        frame.setIconImage(image);*/
        
        ArrayList whiteСell = new ArrayList();
        
    	frame.getContentPane().setLayout(new BorderLayout());
    	Font myFont = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 15);
    	JButton buttonSearch = new JButton("Поиск");
        JButton buttonStart = new JButton("<html><center>" + "Сбросить параметры поиска" + "<center><html>"); 
      
        JPanel panelMain = new JPanel();
   	 	panelMain.add(searchName).setFont(myFont);
   	 	
   	    JCheckBox field1 = new JCheckBox("", false);
   	    new WorkingCheckBox().scaleCheckBoxIcon(field1, 25);
   	    panelMain.add(field1).setFont(myFont);
   	    
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
            HSSFSheet currentSheet = workbook.getSheetAt(0);
            String sD = null;
            
		    for (Row row: currentSheet) {
		    	boolean nRowYellow = false;   
                boolean nRowRed = false; 
                boolean nRowRedWithoutYellow = false; 
                String number = "";
                String name = "";		    	
		        for (Cell cell: row) {
			    	if (row.getRowNum() > 1) {
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
	                    			sD = String.valueOf(sameDay);
	                    			break;
	                    		}
	                            
	                            // анализ 10 столбца
	                            if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
	                            	sameDay = true; 
	                            	
	                            	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
	                            		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
	                                		nRowYellow = true;
	                                	} else {
	                                		nRowYellow = false;
	                                	}
	                                	
	                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                		nRowRed = true;
	                                	} else {
	                                		nRowRed = false;
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
	                                        		nRowYellow = true;
	                                        	} else {
	                                        		nRowYellow = false;
	                                        	}
	                                        	
	                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                        		nRowRed = true;
	                                        	} else {
	                                        		nRowRed = false;
	                                        	}
	                                    	}                                                                     
	                                    } else {
	
	                                        if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
	                                            sameDay = false;  
	                                        } else {          
	                                        	// если месяцы равны
	                                            if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
	                                                sameDay = true;
	                                                nRowYellow = true;
	                                                
	                                             // если месяцы равны и текущий день <= 7
	                                            	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                            		nRowRed = true;
	                                            	} else {
	                                            		nRowRed = false;
	                                            	}
	                                            } else {
	                                                sameDay = false;  	                                                                                        
	                                            }
	                                        }
	                                    }
	                                }
	                            }                                                              
	                            sD = String.valueOf(sameDay);	
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
	                           		nRowYellow = true;
	                           	} else {
	                           		nRowYellow = false;
	                           	}
	                           	
	                           	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                           		nRowRed = true;
	                           	} else {
	                           		nRowRed = false;
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
	                                   		nRowYellow = true;
	                                   	} else {
	                                   		nRowYellow = false;
	                                   	}
	                                   	
	                                   	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                   		nRowRed = true;
	                                   	} else {
	                                   		nRowRed = false;
	                                   	}
	                               	}                                                                     
	                               } else {
	
	                                   if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
	                                       sameDay = false;  
	                                   } else {          
	                                   	// если месяцы равны
	                                       if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
	                                           sameDay = true;
	                                           nRowYellow = true;
	                                           
	                                        // если месяцы равны и текущий день <= 7
	                                       	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                       		nRowRed = true;
	                                       	} else {
	                                       		nRowRed = false;
	                                       	}
	                                       } else {
	                                           sameDay = false;  	                                                                                        
	                                       }
	                                   }
	                               }
	                           }
	                       }                                                              
	                       sD = String.valueOf(sameDay);	
	                   } 
	                   break;
			        }
			      }
			    }
		        for (Cell cell2: row) {
		        	if (row.getRowNum() > 1) {
		        	switch (cell2.getCellTypeEnum()) {
			        	case STRING:
			        	if (sD == "true") {
	                    	if (data.size() == 0) {
	                    		dateGreen.add(0);
	                    	} else {
	                    		dateGreen.add(data.size()/19);
	                    	}
	                    	sD = "false";
	                    }   				        	
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
	                    		kindControl.add(cell2.getStringCellValue());
	                    	}	
	                    }		                    
	                    if (cell2.getColumnIndex() == 2) {
	                    	if (cell2.getStringCellValue().equals("")) {                        		
	                    		data.add("");
	                    	} else {
	                    		data.add(cell2.getStringCellValue());
	                    		appointment.add(cell2.getStringCellValue());
	                    	}
	                    }		                    
	                    if (cell2.getColumnIndex() == 3) {                                
	                    	name = cell2.getStringCellValue();
	                    	data.add(name);   
	                    	allNames.add(cell2.getStringCellValue());
	                    }		                    
	                    if (cell2.getColumnIndex() >= 4 && cell2.getColumnIndex() <= 9) {                               	 
	                    	data.add(cell2.getStringCellValue());
	                    }  		                    
	                    if (cell2.getColumnIndex() == 10) {
	                    	data.add(cell2.getStringCellValue());
	                    	
	            			if (cell2.getStringCellValue().length() != 0) {		            				
	            				if (cell2.getStringCellValue().equals("Не подлежит поверке")) {
	            					whiteСell.add((int) data.size()/19);
	            				}	
	            				if (cell2.getStringCellValue().equals("Не поверяется")) {
	            					whiteСell.add((int) data.size()/19);
	            				}
	            				
	            			} else {
	            				whiteСell.add((int) data.size()/19);
	            			}                                                        		                           	 
	                    }		                    
	                    if (cell2.getColumnIndex() >= 11 && cell2.getColumnIndex() <= 13) {                               	 
	                    	data.add(cell2.getStringCellValue());
	                    }		
	                    if (cell2.getColumnIndex() == 14) {                                   	
	                    	if (!cell2.getStringCellValue().equals("")) {
	                    		formsOwnership.add(cell2.getStringCellValue());
	                    	}	                    	
	                    	data.add(cell2.getStringCellValue());
	                    }		                    
	                    if (cell2.getColumnIndex() == 15) {
	                    	if (!cell2.getStringCellValue().equals("")) {
	                    		equipmentOwners.add(cell2.getStringCellValue());
	                    	}	                    	
	                    	data.add(cell2.getStringCellValue());
	                    }	                    
	                    if (cell2.getColumnIndex() == 16) {
	                    	if (!cell2.getStringCellValue().equals("")) {
	                    		allLocations.add(cell2.getStringCellValue());
	                    	}	                    	
	                    	data.add(cell2.getStringCellValue());
	                    }		                    
	                    if (cell2.getColumnIndex() == 17) {                                   	                               	
	                    	data.add(cell2.getStringCellValue());
	                    	data.add(cell2.getRowIndex()); // добавляем в список номер строки
	                    } 
	                    break;
			        default:
	                	if (sD == "true") {
	                    	if (data.size() == 0) {
	                    		dateGreen.add(0);
	                    	} else {
	                    		dateGreen.add(data.size()/19);
	                    	}
	                    	sD = "false";
	                    }
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
	                			kindControl.add(String.valueOf((int) cell2.getNumericCellValue()));
	                		}	
	                    }
	                	if (cell2.getColumnIndex() == 2) {
	                		if (cell2.getDateCellValue() == null) {
	                			data.add("");
	                		} else {
	                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                			appointment.add(cell2.getStringCellValue());
	                		}	
	                    }
	                	if (cell2.getColumnIndex() == 3) {
	                		if (cell2.getDateCellValue() == null) {
	                			data.add("");
	                		} else {
	                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                			allNames.add(cell2.getStringCellValue());
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
	                    if (cell2.getColumnIndex() == 10) {
	                		
	                    	if (cell2.getDateCellValue() == null) {
	                    		data.add("");
	                    		whiteСell.add((int) data.size()/19);
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
	                    if (cell2.getColumnIndex() == 13) {
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
	                			formsOwnership.add(cell2.getStringCellValue());
	                		}	
	                    }
	                    if (cell2.getColumnIndex() == 15) {
	                    	if (cell2.getDateCellValue() == null) {
	                			data.add("");
	                		} else {
	                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                			equipmentOwners.add(cell2.getStringCellValue());
	                		}	
	                    }
	                    if (cell2.getColumnIndex() == 16) {
	                    	if (cell2.getDateCellValue() == null) {
	                			data.add("");
	                		} else {
	                			data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                			allLocations.add(cell2.getStringCellValue());
	                		}	
	                    }
	                    if (cell2.getColumnIndex() == 17) {
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
	            if (nRowYellow == true && nRowRed == false) {
	    			
	                if (data.size() == 0) {
	                	dateYellow.add(0);
	                } else {
	                	dateYellow.add((data.size()-19)/19);
	                }  
	                // записываем наименование, зав. номер, дату
	    			monthToFinish.add(0, data.get(data.size()-16));
	    			monthToFinish.add(1, data.get(data.size()-13));
	    			monthToFinish.add(2, data.get(data.size()-9));
	    			arrayMonthToFinish.add(monthToFinish.subList(monthToFinish.size()-3, monthToFinish.size()).toString());                      			
	    			nRowYellow = false;
	    			nRowRedWithoutYellow = true;
	    		} 
	                
	        	if (nRowRed == true) {          
	        		
			      	if (data.size() == 0) {
			      		dateRed.add(0);
			      	} else {
			      		dateRed.add((data.size()-19)/19);
			      	}
			        // записываем наименование, зав. номер, дату
			      	weekToFinish.add(data.get(data.size()-16));
			      	weekToFinish.add(data.get(data.size()-13));
			      	weekToFinish.add(data.get(data.size()-9));							      	
			        arrayWeekToFinish.add(weekToFinish.subList(weekToFinish.size()-3, weekToFinish.size()).toString());							      	
	    			nRowRed = false;
	    		} 
	            	
	        	nRowRedWithoutYellow = false;
	        }

            HashSet setG = new HashSet(dateGreen);
            dateGreen.clear();
            dateGreen = new ArrayList<Integer>(setG);

            JSONObject jsonObjectMonth = new JSONObject();
            JSONObject jsonObjectWeek = new JSONObject();
            HashMap<String, String> map1 = new HashMap<String, String>();
            HashMap<String, HashMap<String,String>> massMap = new HashMap<String, HashMap<String,String>>();
            
        	try {
        		int i = 0;
        		
        		while (i < monthToFinish.size()) {
        			if (i < 2) {
        				jsonObjectMonth.put("name", monthToFinish.get(i));
        				jsonObjectMonth.put("number", monthToFinish.get(i+1));
        				jsonObjectMonth.put("date", monthToFinish.get(i+2));    
        				
        			} else {
        				jsonObjectMonth.accumulate("name", monthToFinish.get(i));
        				jsonObjectMonth.accumulate("number", monthToFinish.get(i+1));
        				jsonObjectMonth.accumulate("date", monthToFinish.get(i+2));  
        				
        			}
        			i = i + 3;
        		}        		
        		int j = 0;
        		
        		while (j < weekToFinish.size()) {
        			if (j < 2) {
        				jsonObjectWeek.put("name", weekToFinish.get(j));
        				jsonObjectWeek.put("number", weekToFinish.get(j+1));
        				jsonObjectWeek.put("date", weekToFinish.get(j+2));
        				
        			} else {
        				jsonObjectWeek.accumulate("name", weekToFinish.get(j));
        				jsonObjectWeek.accumulate("number", weekToFinish.get(j+1));
        				jsonObjectWeek.accumulate("date", weekToFinish.get(j+2));
        				
        			}          			
        			j = j + 3;
        		}
        		
        		i = 0;
        		while (i < monthToFinish.size()) {
        				map1 = new HashMap<String, String>();
        				map1.put("name", monthToFinish.get(i).toString());
        				map1.put("number", monthToFinish.get(i+1).toString());
        				map1.put("date", monthToFinish.get(i+2).toString());    
        				massMap.put("File" + String.valueOf(i/3 + 1), map1);
        			i = i + 3;
        		}
			} catch (JSONException e1) {
				e1.printStackTrace();
			}                	                       
        	
        	// отправка данных с данными, срок годности которых скоро пройдет
        	// new ReadJSON().readingObject(jsonObjectWeek);
        	// JButton button3C = new JButton("Отправить JSON");
        	
        	/* button3C.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    	//new HttpRequestPost().givenUsingTimer_whenSchedulingDailyTask_thenCorrect(jsonObjectWeek, 1);
                        //new HttpRequestPost().givenUsingTimer_whenSchedulingDailyTask_thenCorrect(jsonObjectMonth, 2);
        				//new ReportGenerator(jsonObjectWeek).run();
        				//new ReportGenerator(jsonObjectMonth).run();
                }
            });	*/ 
        	
        	//new ReportGenerator(jsonObjectWeek).run();
			//new ReportGenerator(jsonObjectMonth).run();      
        	
            int cl = 19;
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
            
            DefaultTableModel dm = new DefaultTableModel();
            
		    dm.setDataVector(dt, mainHeaders);
            JTable table = new JTable(dm) {};
            
            workbook.close();
            
            table.setPreferredScrollableViewportSize(table.getPreferredSize());
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
                
            // начальный размер фрейма
        	if(startSize == 0) {
        		frame.setPreferredSize(new Dimension(1750, 800));
        		// frame.setPreferredSize(new Dimension(1850, 950));
        		startSize++;       		
        	}
        		
            frame.add(new JScrollPane(panel));
            frame.pack();
            frame.getRootPane().setDefaultButton(buttonSearch);
            frame.setVisible(true);        
            
            copyData = new ArrayList<String>(data);
                        			
			buttonAdd.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
            		try {
            			Set set = new HashSet(appointment);
            			appointment.clear();            			
            			appointment = new ArrayList(set);
            			appointment.add(0, " ");
						new AddDataDC().inputValues(appointment);
						appointment.remove(0);
					} catch (IOException e1) {
						e1.printStackTrace();
					}               	
                }
            });
						
        	ArrayList miniTable = new ArrayList();
        	frame.getContentPane().removeAll();
            
            data.clear();
            data = new ArrayList();
            
            dateGreen.sort(null);
            dateYellow.sort(null);
            dateRed.sort(null);     
            
            JTable table1 = new JTable(dm) {
                
                // запрет на редактирование ячеек в таблице
                private static final long serialVersionUID = 1L;
                
                // кнопку редактирования изменять можно
                public boolean isCellEditable(int row, int column) {                
                	if (column != 18) {
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
	                for (int i = 0; i < dateGreen.size(); i++) {
	                	if (row == ((int) dateGreen.get(i)) && column == 10) {
	                		//c.setBackground(new java.awt.Color(174, 212, 149));
	                		c.setBackground(Color.white);
	                	}
	                }	   	                
	                for (int i = 0; i < whiteСell.size(); i++) {
	                	if (row == ((int) whiteСell.get(i)) && column == 10) {
	                		// c.setBackground(Color.white);
	                	}
	                } 				                
	                for (int i = 0; i < dateYellow.size(); i++) {
	                	if (row == ((int) dateYellow.get(i)) && column == 10) {
	                		c.setBackground(new java.awt.Color(247, 239, 162));
	                	}
	                }  				                
	                for (int i = 0; i < dateRed.size(); i++) {
	                	if (row == ((int) dateRed.get(i)) && column == 10) {
	                		c.setBackground(new java.awt.Color(213, 92, 95));
	                	}
	                }	

	                return c;	               	                
	            }
		    };
		    
            table1.setPreferredScrollableViewportSize(table.getPreferredSize());
            table1.changeSelection(0, 0, false, false);
            JScrollPane scrollPane1 = new JScrollPane( table1 );
            getContentPane().add(scrollPane1);
            
            // кнопка "редактировать"    
            table1.getColumn(" ").setCellRenderer(new ButtonRendererDC(frame));
            table1.getColumn(" ").setCellEditor(new ButtonEditorDC(new JCheckBox(), frame));
            table1.getColumnModel().getColumn(18).setPreferredWidth(30);
            // кнопка "редактировать"   
            
            JTableHeader th1 = table1.getTableHeader();
        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));              	
        	
        	// горизонтальная прокрутка заголовков
        	table1.getTableHeader().setPreferredSize(new Dimension(10000,120));
        	
        	table1.getColumnModel().getColumn(0).setPreferredWidth(35);
        	
        	table1.getColumnModel().getColumn(1).setPreferredWidth(140); 
        	table1.getColumnModel().getColumn(1).setMaxWidth(140);
        	table1.getColumnModel().getColumn(1).setMinWidth(140);
        	
        	table1.getColumnModel().getColumn(2).setPreferredWidth(170);  
        	table1.getColumnModel().getColumn(2).setMaxWidth(170);
        	table1.getColumnModel().getColumn(2).setMinWidth(170);
        	
        	table1.getColumnModel().getColumn(3).setPreferredWidth(170);
        	table1.getColumnModel().getColumn(3).setMaxWidth(170);
        	table1.getColumnModel().getColumn(3).setMinWidth(170);
        	
        	table1.getColumnModel().getColumn(4).setPreferredWidth(170); 
        	table1.getColumnModel().getColumn(4).setMaxWidth(170);
        	table1.getColumnModel().getColumn(4).setMinWidth(170);
        	
        	table1.getColumnModel().getColumn(5).setPreferredWidth(180); 
        	table1.getColumnModel().getColumn(5).setMaxWidth(180);
        	table1.getColumnModel().getColumn(5).setMinWidth(180);
        	
			table1.getColumnModel().getColumn(6).setPreferredWidth(100);
			table1.getColumnModel().getColumn(7).setPreferredWidth(100);
			
			table1.getColumnModel().getColumn(8).setPreferredWidth(70);       	
			table1.getColumnModel().getColumn(9).setPreferredWidth(150);	
			
			table1.getColumnModel().getColumn(10).setPreferredWidth(170); 
        	table1.getColumnModel().getColumn(10).setMaxWidth(170);
        	table1.getColumnModel().getColumn(10).setMinWidth(170);
        	
			table1.getColumnModel().getColumn(11).setPreferredWidth(110);
			table1.getColumnModel().getColumn(12).setPreferredWidth(150);
			table1.getColumnModel().getColumn(13).setPreferredWidth(110);
			table1.getColumnModel().getColumn(14).setPreferredWidth(150);
			table1.getColumnModel().getColumn(15).setPreferredWidth(155);
			table1.getColumnModel().getColumn(16).setPreferredWidth(155);
			table1.getColumnModel().getColumn(17).setPreferredWidth(150);
			
			table1.setRowHeight(25);
			
			for (int i = 0; i <= mainHeaders.length - 2; i++) {
        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
        	}
			
			ArrayList rZ = new ArrayList();			
            
            JPanel panel12 = new JPanel(new BorderLayout(10, 10));
            JPanel panel01 = new JPanel(new BorderLayout(0, 25));
            
            table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);          
            
            panel01.setPreferredSize(new Dimension(0, 267)); // was 300
            
            three.setPreferredSize(new Dimension(0, 10));
            three.setMaximumSize(new Dimension(0, 10));
            three.setMinimumSize(new Dimension(0, 10));
            
            JPanel panelL = new JPanel(new GridLayout(5, 2, 10, 10));
            
            GridBagConstraints c = new GridBagConstraints();        
            
            Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);                                 
    	    
            Set setAp = new HashSet(appointment);
            appointment.clear();
            appointment = new ArrayList(setAp);
            
    	    panelL.add(new JLabel("Назначение:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceAppointment().outputPanel(appointment, "Все назначения")).setFont(bigFontTR); 
            
    	    JCheckBox field4 = new JCheckBox("", false);
    	    new WorkingCheckBox().scaleCheckBoxIcon(field4, 25);
    	    panelL.add(new JLabel("<html><right>" + "Указание в поверке на <br> принадлежность к организации:" +
    	    "<html>", SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(field4).setFont(bigFontTR);
            
            Set setKC = new HashSet(kindControl);
            kindControl.clear();
            kindControl = new ArrayList(setKC);        
            
            panelL.add(new JLabel("Вид контроля:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceKindsControl().outputPanel(kindControl, "Все виды контроля")).setFont(bigFontTR); 
            
            panelL.add(new JLabel("Действующая поверка:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(field1).setFont(bigFontTR);                     
      	    
       	    Set setTO = new HashSet(formsOwnership);
            formsOwnership.clear();
            formsOwnership = new ArrayList(setTO);
            
            Set setL = new HashSet(allLocations);
            allLocations.clear();
            allLocations = new ArrayList(setL); 
            
            Set setEO = new HashSet(equipmentOwners);
            equipmentOwners.clear();
            equipmentOwners = new ArrayList(setEO);       
            
            panelL.add(new JLabel("Форма собственности:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceFormOwnership().outputPanel(formsOwnership, "Все формы собственности")).setFont(bigFontTR);
            
            JCheckBox field2 = new JCheckBox("", false);
            new WorkingCheckBox().scaleCheckBoxIcon(field2, 25);
       	    panelL.add(new JLabel("Не поверяются:", SwingConstants.RIGHT)).setFont(bigFontTR);
       	    panelL.add(field2).setFont(bigFontTR);      	    
            
            panelL.add(new JLabel("Владелец оборудования:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceEquipmentOwners().outputPanel(equipmentOwners, "Все владельцы оборудования")).setFont(bigFontTR);
            
       	    JCheckBox field3 = new JCheckBox("", false);
       	    new WorkingCheckBox().scaleCheckBoxIcon(field3, 25);
    	    panelL.add(new JLabel("Техническое состояние исправно:", SwingConstants.RIGHT)).setFont(bigFontTR);
    	    panelL.add(field3).setFont(bigFontTR);
    	     
            panelL.add(new JLabel("Местонахождение:", SwingConstants.RIGHT)).setFont(bigFontTR);
            
            panelL.add(new ChoiceLocation().outputPanel(allLocations, "Все местонахождения")).setFont(bigFontTR);
            
            JPanel panelL2 = new JPanel();
            panelL2.add(new JLabel("Наименование:", SwingConstants.RIGHT)).setFont(bigFontTR);
            searchName.setMaximumSize(new Dimension(10, 14));
            searchName.setMinimumSize(new Dimension(10, 14));
            panelL2.add(searchName);         
            
            panelL.add(panelL2);
            
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
            buttonDC.setForeground(Color.WHITE);
            JButton buttonCars = new JButton("Автомобили");
            buttonCars.setBackground(Color.LIGHT_GRAY);
            JButton buttonTechnology = new JButton("Орг. техника");
            buttonTechnology.setBackground(Color.LIGHT_GRAY);
                                   
            JPanel buttonsPanels = new JPanel(new GridLayout(3, 1, 40, 40));
            
            buttonsPanels.add(buttonDC);
            buttonsPanels.add(buttonCars);
            buttonsPanels.add(buttonTechnology);
            
            JPanel panelsSearch = new JPanel(new GridBagLayout());
            GridBagConstraints cPB = new GridBagConstraints();

            cPB.fill = GridBagConstraints.HORIZONTAL;
            cPB.gridx = 0;
            cPB.gridy = 1;            
            cPB.weightx = 1;
            cPB.ipady = -30;
            cPB.insets = new Insets(5, 0, 0, 0);
            cPB.fill = GridBagConstraints.BOTH;
            
            panelsSearch.add(panelL, cPB);
                        
            cPB.fill = GridBagConstraints.HORIZONTAL;
            cPB.gridx = 1;
            cPB.gridy = 1;           
            cPB.weightx = 0.01;
            cPB.ipadx = 10;
            // вверх, вниз, влево, вправо
            cPB.insets = new Insets(15, 0, 0, 10);
            cPB.fill = GridBagConstraints.BOTH;    
            
            panelsSearch.add(buttonsPanels, cPB);
            
            setLayout(new GridLayout(2, 1, 1, 1));
            panel01.add(panelsSearch, BorderLayout.NORTH);
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
            
            cH = panel01.getHeight() - 100;
            cW = (panelL.getWidth() / 2) + 100;
            
            buttonTechnology.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                	searchName.setText("");
                    searchName = new JTextField(14); 
                    dateGreen.clear(); 
    			    dateYellow.clear(); 
    			    dateRed.clear();
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
            
            buttonCars.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                	searchName.setText("");
                    searchName = new JTextField(14); 
                    dateGreen.clear(); 
    			    dateYellow.clear(); 
    			    dateRed.clear();
    			    jointUpload.clear();
    			    copyData.clear();
                    frame.setPreferredSize(frame.getSize());                    
                	try {
    					new Cars().start(frame, three.getHeight(), table1);
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
	            	searchName.setText("");
	                searchName = new JTextField(14); 
	                dateGreen.clear(); 
				    dateYellow.clear(); 
				    dateRed.clear();
				    jointUpload.clear();
				    copyData.clear();
				    allNames.clear();
	                frame.setPreferredSize(frame.getSize());	                
	            	try {
						start(frame, 1);
					} catch (FileNotFoundException e1) {
						e1.printStackTrace();
					} catch (IOException e1) {
						e1.printStackTrace();
					} catch (ParseException e1) {
						e1.printStackTrace();
					}  
	            }
            });               
            
        Set setN = new HashSet(allNames);
        allNames.clear();
        allNames = new ArrayList(setN);
        
		new AutoSuggestor(searchName, frame, null, Color.WHITE.brighter(), Color.DARK_GRAY, Color.RED, 0.8f, cH, cW) {		 	
			boolean wordTyped(String typedWord)  {	     	             		
				setDictionary(allNames);
		 		return super.wordTyped(typedWord);
		    }                                       
		};		
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		buttonSearch.addActionListener(new ActionListener() {			
			@Override
			public void actionPerformed(ActionEvent e) {
			    ArrayList data = new ArrayList();
			    dateGreen.clear(); 
			    dateYellow.clear(); 
			    dateRed.clear(); 
			    rZ.clear();
			   
			    ArrayList selectedAppointment = new ChoiceAppointment().selecting();
			    ArrayList selectedKindsControl = new ChoiceKindsControl().selectedHeaders();
			    ArrayList selectedFormsOwnership = new ChoiceFormOwnership().selectedHeaders();
			    ArrayList selectedEquipmentOwners = new ChoiceEquipmentOwners().selectedHeaders();
			    ArrayList selectedLocations = new ChoiceLocation().selectedLocations();
			    ArrayList whiteСell = new ArrayList();
			    
				try {			
				    InputStream inputStream = new FileInputStream("Reestr.xls");
				    Workbook workbook = new HSSFWorkbook(inputStream);
				    int nList = 0;
				    Sheet currentSheet = workbook.getSheetAt(nList);
				    Iterator<Row> rowIterator = currentSheet.iterator();
				    rowIterator.next(); 
				    rowIterator.next();
				    String number = "";
				    String name = "";
				    String c2String = "";
				    boolean h = false;                        						    
				    String sD = null;
				    int currentRow = 0;       					        
			        
				    for (Row row: currentSheet) {
				    	boolean b = false;
	                    boolean column8 = false;
	                    boolean column3 = false; 
	                    boolean column11 = false; 
	                    boolean column12 = false; 
	                    boolean column13 = false; 
	                    boolean column14 = false; 		
	                    boolean column15 = false; 
	                    boolean column10 = false;
	                    boolean column10DP = false;
	                    boolean column2 = false;
			            boolean nRowYellow = false;   
			            boolean nRowRed = false; 
			            boolean nRowRedWithoutYellow = false; 	  	
				        for (Cell cell: row) {
					    	if (row.getRowNum() > 1) {
					    	int columnIndex = cell.getColumnIndex();
					    	
					        switch (cell.getCellTypeEnum()) {
						        case STRING:			        				        
						        	
						        	if (cell.getColumnIndex() == 0) {
			                        	if (cell.getCellTypeEnum() != NUMERIC) {   				                        		
			                        		number = cell.getStringCellValue();
			                        	} else {				                        		                                                		
			                        		int x = (int) cell.getNumericCellValue();                                   		
			                        		number = String.valueOf(x);
			                        	}
			                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 1) {
			                        	c2String = cell.getStringCellValue();

			                        	if (selectedKindsControl.get(0).equals("Все виды контроля")) {
			                        		column2 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedKindsControl.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedKindsControl.get(i))) {				                        		
					                        		column2 = true;
					                        	} 
				                        	}
			                        	}	
			                        }
			                        
			                        if (cell.getColumnIndex() == 2) {
			                        	name = cell.getStringCellValue();
			                        	if (selectedAppointment.get(0).equals("Все назначения")) {
			                        		column3 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedAppointment.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedAppointment.get(i))) {				                        		
					                        		column3 = true;
					                        	} 
				                        	}
			                        	}	
			                        }
			                        
			                        if (cell.getColumnIndex() == 3) {					                        	
			                        	
			                        	int indexM = cell.getStringCellValue().indexOf(searchName.getText());
			                        	
			                        	if (indexM != -1) {
			                        		b = true;
			                        	}				                        	

			                        	if (cell.getStringCellValue().replaceAll("\\s","").indexOf(searchName.getText()) != -1) {
			                        		b = true;
			                        	} 				                        	
			                        }				                        				                        		                        
			                        
			                        if (cell.getColumnIndex() == 12) {				                        	
			                        	if (cell.getStringCellValue().equals("Исправен") | cell.getStringCellValue().equals("Исправен.") | 
			                        			cell.getStringCellValue().equals("Исправно") | cell.getStringCellValue().equals("испр")) {			                        		
			                        		column11 = true;
			                        	} 
			                        }
			                        
			                        if (cell.getColumnIndex() == 13) {
			                        	
			                        	if (!cell.getStringCellValue().equals("нет") && !cell.getStringCellValue().equals("")
			                        		&& !cell.getStringCellValue().equals("-") | !cell.getStringCellValue().equals("???")) {				                        		
			                        		column12 = true;				                        		
			                        	}				                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 14) {  
			                        	
			                        	if (selectedFormsOwnership.get(0).equals("Все формы собственности")) {
			                        		column13 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedFormsOwnership.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedFormsOwnership.get(i))) {				                        		
					                        		column13 = true;
					                        	} 
				                        	}
			                        	}				                        					                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 15) {   
			                        	
			                        	if (selectedEquipmentOwners.get(0).equals("Все владельцы оборудования")) {
			                        		column14 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedEquipmentOwners.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedEquipmentOwners.get(i))) {				                        		
					                        		column14 = true;
					                        	} 
				                        	}
			                        	}				                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 16) { 
			                        	if (selectedLocations.get(0).equals("Все местонахождения")) {
			                        		column15 = true;
			                        	} else {
				                        	for (int i = 0; i < selectedLocations.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedLocations.get(i))) {				                        		
					                        		column15 = true;
					                        	} 
				                        	}
			                        	}
			                        }	
			                        
			                        if (cell.getColumnIndex() == 10) {

	                                	boolean sameDay = false; 
	                                	Date dateNow = new Date();
	                                	Calendar cal1 = Calendar.getInstance();
	                                    Calendar cal2 = Calendar.getInstance();
	                                    
	                                    // определяем тип данных и получаем значение 
	                                    if ((cell.getCellType() != Cell.CELL_TYPE_FORMULA && cell.getCellTypeEnum() == NUMERIC)
	                                    || (cell.getCellType() == Cell.CELL_TYPE_FORMULA && cell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC)) {
	                                    	sameDay = false; 
	                                    	
	                                    	if (cell.getDateCellValue() != null) {                                                                                       
	                                            cal1.setTime(dateNow);
	                                            cal2.setTime(cell.getDateCellValue());
	                                    	} else {
	                                    		break;
	                                    	}		
	                                    	
	                                    } else {
	                                    	
	                                    	String regex = "(\\d{2}.\\d{2}.\\d{4})";
	                                    	
	                                    	if (cell.getStringCellValue().length() == 0) {				                                    
	                                    		 column10DP = true;
	                                    	}
	                                    	if (cell.getStringCellValue().equals("-") | cell.getStringCellValue().equals("???")) {				                                    
	                                    		 column10DP = true;
	                                    	}
	                                    	if (cell.getStringCellValue().equals("Не подлежит поверке") | cell.getStringCellValue().equals("Не поверяется")) {
	                                    		column10DP = true;
                                			}
	                                    	
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
	                                			
	                                			if (cell.getStringCellValue().equals("Не подлежит поверке") | cell.getStringCellValue().equals("Не поверяется")) {
	                                				column8 = true;
	                                			}		                                					                              			
	                                			sameDay = false;
	                                			sD = String.valueOf(sameDay);
	                                		}
	                                    }	
	                                    
	                                    // анализ 10 столбца
	                                    if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
	                                    	sameDay = true; 
	                                    	
	                                    	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
	                                    		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
	                                        		nRowYellow = true;
	                                        	} else {
	                                        		nRowYellow = false;
	                                        	}
	                                        	
	                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                        		nRowRed = true;
	                                        	} else {
	                                        		nRowRed = false;
	                                        	}
	                                    	}
	                                    	
	                                    } else {
	                                        if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
	                                            sameDay = false;
	                                        } else {
	                                            if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
	                                            	sameDay = true;
	                                            	
	                                                   // если текущий месяц больше на 1
	                                                if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
	                                                	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
	                                                		nRowYellow = true;
	                                                	} else {
	                                                		nRowYellow = false;
	                                                	}
	                                                	
	                                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                                		nRowRed = true;
	                                                	} else {
	                                                		nRowRed = false;
	                                                	}
	                                            	}                                                                     
	                                            } else {

	                                                if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
	                                                    sameDay = false;  
	                                                } else {                                                       	
	                                                    if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
	                                                        sameDay = true;
	                                                        nRowYellow = true;
	                                                        
	                                                     // если месяцы равны и текущий день <= 7
	                                                    	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                                    		nRowRed = true;
	                                                    	} else {
	                                                    		nRowRed = false;
	                                                    	}
	                                                    } else {
	                                                        sameDay = false;  	                                                                                        
	                                                    }
	                                                }
	                                            }
	                                        }
	                                    }                                                              
	                                    sD = String.valueOf(sameDay);
	                                    currentRow = cell.getRowIndex();			                                
			                        } 	
			                        break;
						        default:               	   
			                	   
			                        if (cell.getColumnIndex() == 0) {
			                        	if (cell.getCellTypeEnum() != NUMERIC) {   				                        		
			                        		number = cell.getStringCellValue();
			                        	} else {				                        		                                                		
			                        		int x = (int) cell.getNumericCellValue();                                   		
			                        		number = String.valueOf(x);
			                        	}
			                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 1) {
			                        	c2String = cell.getStringCellValue();

			                        	if (selectedKindsControl.get(0).equals("Все виды контроля")) {
			                        		column2 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedKindsControl.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedKindsControl.get(i))) {				                        		
					                        		column2 = true;
					                        	} 
				                        	}
			                        	}	
			                        }
			                        
			                        if (cell.getColumnIndex() == 2) {
			                        	name = cell.getStringCellValue();
			                        	if (selectedAppointment.get(0).equals("Все назначения")) {
			                        		column3 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedAppointment.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedAppointment.get(i))) {				                        		
					                        		column3 = true;
					                        	} 
				                        	}
			                        	}	
			                        }
			                        
			                        if (cell.getColumnIndex() == 3) {					                        	
			                        	
			                        	int indexM = cell.getStringCellValue().indexOf(searchName.getText());
			                        	
			                        	if (indexM != -1) {
			                        		b = true;
			                        	}				                        	

			                        	if (cell.getStringCellValue().replaceAll("\\s","").indexOf(searchName.getText()) != -1) {
			                        		b = true;
			                        	} 				                        	
			                        }				                        				                        		                        
			                        
			                        if (cell.getColumnIndex() == 12) {				                        	
			                        	if (cell.getStringCellValue().equals("Исправен") | cell.getStringCellValue().equals("Исправен.") | 
			                        			cell.getStringCellValue().equals("Исправно") | cell.getStringCellValue().equals("испр")) {			                        		
			                        		column11 = true;
			                        	} 
			                        }
			                        
			                        if (cell.getColumnIndex() == 13) {
			                        	
			                        	if (!cell.getStringCellValue().equals("нет") && !cell.getStringCellValue().equals("")
			                        		&& !cell.getStringCellValue().equals("-") | !cell.getStringCellValue().equals("???")	) {				                        		
			                        		column12 = true;				                        		
			                        	}				                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 14) {  
			                        	
			                        	if (selectedFormsOwnership.get(0).equals("Все формы собственности")) {
			                        		column13 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedFormsOwnership.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedFormsOwnership.get(i))) {				                        		
					                        		column13 = true;
					                        	} 
				                        	}
			                        	}				                        					                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 15) {   
			                        	
			                        	if (selectedEquipmentOwners.get(0).equals("Все владельцы оборудования")) {
			                        		column14 = true;
			                        	} else {
			                        		
				                        	for (int i = 0; i < selectedEquipmentOwners.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedEquipmentOwners.get(i))) {				                        		
					                        		column14 = true;
					                        	} 
				                        	}
			                        	}				                        	
			                        }
			                        
			                        if (cell.getColumnIndex() == 16) { 
			                        	if (selectedLocations.get(0).equals("Все местонахождения")) {
			                        		column15 = true;
			                        	} else {
				                        	for (int i = 0; i < selectedLocations.size(); i++) {
					                        	if (cell.getStringCellValue().equals(selectedLocations.get(i))) {				                        		
					                        		column15 = true;
					                        	} 
				                        	}
			                        	}
			                        }	
			                        
			                        if (cell.getColumnIndex() == 10) {

	                                	boolean sameDay = false; 
	                                	Date dateNow = new Date();
	                                	Calendar cal1 = Calendar.getInstance();
	                                    Calendar cal2 = Calendar.getInstance();
	                                    
	                                    // определяем тип данных и получаем значение 
	                                    if ((cell.getCellType() != Cell.CELL_TYPE_FORMULA && cell.getCellTypeEnum() == NUMERIC)
	                                    || (cell.getCellType() == Cell.CELL_TYPE_FORMULA && cell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC)) {
	                                    	sameDay = false; 
	                                    	
	                                    	if (cell.getDateCellValue() != null) {                                                                                       
	                                            cal1.setTime(dateNow);
	                                            cal2.setTime(cell.getDateCellValue());
	                                    	} else {
	                                    		break;
	                                    	}		
	                                    	
	                                    } else {
	                                    	
	                                    	String regex = "(\\d{2}.\\d{2}.\\d{4})";
	                                    	
	                                    	if (cell.getStringCellValue().length() == 0) {				                                    
	                                    		 column10DP = true;
	                                    	}
	                                    	if (cell.getStringCellValue().equals("-") | cell.getStringCellValue().equals("???")) {				                                    
	                                    		 column10DP = true;
	                                    	}
	                                    	if (cell.getStringCellValue().equals("Не подлежит поверке") | cell.getStringCellValue().equals("Не поверяется")) {
	                                    		column10DP = true;
                                			}
	                                    	
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
	                                			
	                                			if (cell.getStringCellValue().equals("Не подлежит поверке") | cell.getStringCellValue().equals("Не поверяется")) {
	                                				column8 = true;
	                                			}		                                					                              			
	                                			sameDay = false;
	                                			sD = String.valueOf(sameDay);
	                                		}
	                                    }	
	                                    
	                                    // анализ 10 столбца
	                                    if (cal1.get(Calendar.YEAR) < cal2.get(Calendar.YEAR)) {
	                                    	sameDay = true; 
	                                    	
	                                    	if (cal1.get(Calendar.MONTH) == 11 && cal2.get(Calendar.MONTH) == 0) {	                                            		
	                                    		if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
	                                        		nRowYellow = true;
	                                        	} else {
	                                        		nRowYellow = false;
	                                        	}
	                                        	
	                                        	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                        		nRowRed = true;
	                                        	} else {
	                                        		nRowRed = false;
	                                        	}
	                                    	}
	                                    	
	                                    } else {
	                                        if (cal1.get(Calendar.YEAR) > cal2.get(Calendar.YEAR)) {
	                                            sameDay = false;
	                                        } else {
	                                            if (cal1.get(Calendar.MONTH) < cal2.get(Calendar.MONTH)) {
	                                            	sameDay = true;
	                                            	
	                                                   // если текущий месяц больше на 1
	                                                if (cal2.get(Calendar.MONTH) - cal1.get(Calendar.MONTH) == 1) {                                                     	
	                                                	if ((30 - cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 30) {
	                                                		nRowYellow = true;
	                                                	} else {
	                                                		nRowYellow = false;
	                                                	}
	                                                	
	                                                	if (30 - (cal1.get(Calendar.DAY_OF_MONTH)) +  cal2.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                                		nRowRed = true;
	                                                	} else {
	                                                		nRowRed = false;
	                                                	}
	                                            	}                                                                     
	                                            } else {

	                                                if (cal1.get(Calendar.MONTH) > cal2.get(Calendar.MONTH)) {
	                                                    sameDay = false;  
	                                                } else {                                                       	
	                                                    if (cal1.get(Calendar.DAY_OF_MONTH) <= cal2.get(Calendar.DAY_OF_MONTH)) {
	                                                        sameDay = true;
	                                                        nRowYellow = true;
	                                                        
	                                                     // если месяцы равны и текущий день <= 7
	                                                    	if (cal2.get(Calendar.DAY_OF_MONTH) - cal1.get(Calendar.DAY_OF_MONTH) <= 7) {
	                                                    		nRowRed = true;
	                                                    	} else {
	                                                    		nRowRed = false;
	                                                    	}
	                                                    } else {
	                                                        sameDay = false;  	                                                                                        
	                                                    }
	                                                }
	                                            }
	                                        }
	                                    }                                                              
	                                    sD = String.valueOf(sameDay);
	                                    currentRow = cell.getRowIndex();			                                
			                        } 			                  
			                   break;
					        }
					      }
					    }
				        
	                    if (sD == "true" && field1.isSelected() == true && column8 == false && column10DP == false) {
                        	column10 = true;      
                        	column10DP = false;
                        }  
	                    
	                    if (field1.isSelected() == false) {
	                    	column10 = true;
	                    }
	                    
	                    if (field2.isSelected() == false) {
	                    	column8 = true;
	                    }
	                    
	                    if (field2.isSelected() == true) {
	                    	whiteСell.add((int) data.size()/19);
	                    }

	                    if (field3.isSelected() == false) {
	                    	column11 = true;
	                    }
	                    if (field4.isSelected() == false) {
	                    	column12 = true;
	                    }
	                    
	                    int nRow = 0;
	                    
	                    if (b == true && column12 == true && column11 == true && column8 == true && column10 == true 
	                    		&& column2 == true && column13 == true && column14 == true && column15 == true && column3 == true) {
				        for (Cell cell2: row) {
				        	if (row.getRowNum() > 1) {
				        	switch (cell2.getCellTypeEnum()) {
					        	case STRING:
						        	if (sD == "true" && field2.isSelected() == false && column10DP == false) {
		                            	if (data.size() == 0) {
		                            		dateGreen.add(0);
		                            	} else {
		                            		dateGreen.add(data.size()/19);
		                            	}
		                            	sD = "false";
		                            }               

		                            if (cell2.getColumnIndex() == 3) {     
		                            	if (number.length() != 0) {
			                        		char chEnd = number.charAt(number.length()-1);		                        		
			                        		if (chEnd == '.') {
			                        			rZ.add(((int) data.size()/18) + 2);
			                        		}
			                            }
		                            	data.add(number);
		                            	data.add(c2String); 	                            	
		                            	data.add(name);	
		                            	data.add(cell2.getStringCellValue()); 
	                                }
		                            
	                                if (cell2.getColumnIndex() == 4) {                               	 
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 5) {
		                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
		                            		data.add(cell2.getStringCellValue());
		                            	} else {
		                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                            	}	                            	
	                                }
	                                
	                                if (cell2.getColumnIndex() == 6) {
		                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
		                            		data.add(cell2.getStringCellValue());
		                            	} else {
		                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                            	}	                            	
	                                }
	                                                                
	                                if (cell2.getColumnIndex() == 7) {
		                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
		                            		data.add(cell2.getStringCellValue());
		                            	} else {
		                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
		                            	}	                            	
	                                }
			                        
			                        if (cell2.getColumnIndex() == 8) {
			                        	
			                            if (cell2.getCellTypeEnum() != NUMERIC) {                                     	
			                                data.add(cell2.getStringCellValue());
			                            } else {                                       	
			                            	if (cell2.getDateCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
			                                    data.add(ft.format(cell2.getDateCellValue()));
			                            	}
			                            }                                  	
			                        }
	                                
	                                if (cell2.getColumnIndex() == 9) {                               	 
	                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
	                                		if (cell2.getStringCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                            		data.add(cell2.getStringCellValue());
			                            	}	
			                            } else {                                       	
			                            	if (cell2.getDateCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
							                    data.add(ft.format(cell2.getDateCellValue()));
			                            	}
			                            }
	                                }

									if (cell2.getColumnIndex() == 10) {
										if (cell2.getCellType() == Cell.CELL_TYPE_FORMULA) {
								            if (cell2.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC) {
								            	if (cell2.getDateCellValue() == null) {
								            		whiteСell.add((int) data.size()/19);
								            		data.add("");
								            	} else {
								            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
								                    data.add(ft.format(cell2.getDateCellValue()));
								            	}                                          	
								            } else {
								            	if (cell2.getStringCellValue().length() != 0) {
								            		if (cell2.getStringCellValue().equals("Не подлежит поверке")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								            		if (cell2.getStringCellValue().equals("Не поверяется")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								            		if (cell2.getStringCellValue().equals("-") | cell2.getStringCellValue().equals("???")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								    			} else {
								    				whiteСell.add((int) data.size()/19);
								    			}
								            	
								            	data.add(cell2.getStringCellValue());
								            }
										} else {  
											if (cell2.getCellTypeEnum() != NUMERIC) {     
												
												if (cell2.getStringCellValue().length() != 0) {
								            		if (cell2.getStringCellValue().equals("Не подлежит поверке")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								            		if (cell2.getStringCellValue().equals("Не поверяется")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								            		if (cell2.getStringCellValue().equals("-") | cell2.getStringCellValue().equals("???")) {
								            			whiteСell.add((int) data.size()/19);
								            		}
								    			} else {
								    				whiteСell.add((int) data.size()/19);
								    			}
								        		
								                data.add(cell2.getStringCellValue());
								        			
								            } else {                                       	
								            	if (cell2.getDateCellValue() == null) {
								            		whiteСell.add((int) data.size()/19);
								            		data.add("");
								            	} else {
								            		
								            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
								                    data.add(ft.format(cell2.getDateCellValue()));
								            	}
								            } 
										}                               	 
									}
	                                
	                                if (cell2.getColumnIndex() == 11) {
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 12) {                                      	
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 13) {                                      	
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 14) {                                      	
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 15) {                                      	
	                                	data.add(cell2.getStringCellValue());
	                                }
	                                
	                                if (cell2.getColumnIndex() == 16) { 
	                                	
	                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
	                                		if (cell2.getStringCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                            		data.add(cell2.getStringCellValue());
			                            	}	
			                            } else {                                       	
			                            	if (cell2.getDateCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                                    data.add((int) cell2.getNumericCellValue());
			                            	}
			                            }
	                                }	
	                                
	                                if (cell2.getColumnIndex() == 17) { 
	                                	
	                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
	                                		if (cell2.getStringCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                            		data.add(cell2.getStringCellValue());
			                            	}	
			                            } else {                                       	
			                            	if (cell2.getDateCellValue() == null) {
			                            		data.add("");
			                            	} else {
			                                    data.add((int) cell2.getNumericCellValue());
			                            	}
			                            }
	                                	data.add(cell2.getRowIndex()); // добавляем в список номер строки
	                                }	 
			                    break;
				        	default:
					        	if (sD == "true" && field2.isSelected() == false && column10DP == false) {
	                            	if (data.size() == 0) {
	                            		dateGreen.add(0);
	                            	} else {
	                            		dateGreen.add(data.size()/19);
	                            	}
	                            	sD = "false";
	                            }               

	                            if (cell2.getColumnIndex() == 3) {     
	                            	if (number.length() != 0) {
		                        		char chEnd = number.charAt(number.length()-1);		                        		
		                        		if (chEnd == '.') {
		                        			rZ.add(((int) data.size()/18) + 2);
		                        		}
		                            }
	                            	data.add(number);
	                            	data.add(c2String); 	                            	
	                            	data.add(name);	
	                            	data.add(cell2.getStringCellValue()); 
                                }
	                            
                                if (cell2.getColumnIndex() == 4) {                               	 
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 5) {
	                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
	                            		data.add(cell2.getStringCellValue());
	                            	} else {
	                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                            	}	                            	
                                }
                                
                                if (cell2.getColumnIndex() == 6) {
	                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
	                            		data.add(cell2.getStringCellValue());
	                            	} else {
	                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                            	}	                            	
                                }
                                                                
                                if (cell2.getColumnIndex() == 7) {
	                            	if (cell2.getCellTypeEnum() != NUMERIC) {  	                            		
	                            		data.add(cell2.getStringCellValue());
	                            	} else {
	                            		data.add(String.valueOf((int) cell2.getNumericCellValue()));
	                            	}	                            	
                                }
		                        
		                        if (cell2.getColumnIndex() == 8) {
		                        	
		                            if (cell2.getCellTypeEnum() != NUMERIC) {                                     	
		                                data.add(cell2.getStringCellValue());
		                            } else {                                       	
		                            	if (cell2.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
		                                    data.add(ft.format(cell2.getDateCellValue()));
		                            	}
		                            }                                  	
		                        }
                                
                                if (cell2.getColumnIndex() == 9) {                               	 
                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
                                		if (cell2.getStringCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		data.add(cell2.getStringCellValue());
		                            	}	
		                            } else {                                       	
		                            	if (cell2.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
						                    data.add(ft.format(cell2.getDateCellValue()));
		                            	}
		                            }
                                }

								if (cell2.getColumnIndex() == 10) {
									if (cell2.getCellType() == Cell.CELL_TYPE_FORMULA) {
							            if (cell2.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC) {
							            	if (cell2.getDateCellValue() == null) {
							            		whiteСell.add((int) data.size()/19);
							            		data.add("");
							            	} else {
							            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
							                    data.add(ft.format(cell2.getDateCellValue()));
							            	}                                          	
							            } else {
							            	if (cell2.getStringCellValue().length() != 0) {
							            		if (cell2.getStringCellValue().equals("Не подлежит поверке")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							            		if (cell2.getStringCellValue().equals("Не поверяется")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							            		if (cell2.getStringCellValue().equals("-") | cell2.getStringCellValue().equals("???")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							    			} else {
							    				whiteСell.add((int) data.size()/19);
							    			}
							            	
							            	data.add(cell2.getStringCellValue());
							            }
									} else {  
										if (cell2.getCellTypeEnum() != NUMERIC) {     
											
											if (cell2.getStringCellValue().length() != 0) {
							            		if (cell2.getStringCellValue().equals("Не подлежит поверке")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							            		if (cell2.getStringCellValue().equals("Не поверяется")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							            		if (cell2.getStringCellValue().equals("-") | cell2.getStringCellValue().equals("???")) {
							            			whiteСell.add((int) data.size()/19);
							            		}
							    			} else {
							    				whiteСell.add((int) data.size()/19);
							    			}
							        		
							                data.add(cell2.getStringCellValue());
							        			
							            } else {                                       	
							            	if (cell2.getDateCellValue() == null) {
							            		whiteСell.add((int) data.size()/19);
							            		data.add("");
							            	} else {
							            		
							            		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
							                    data.add(ft.format(cell2.getDateCellValue()));
							            	}
							            } 
									}                               	 
								}
                                
                                if (cell2.getColumnIndex() == 11) {
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 12) {                                      	
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 13) {                                      	
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 14) {                                      	
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 15) {                                      	
                                	data.add(cell2.getStringCellValue());
                                }
                                
                                if (cell2.getColumnIndex() == 16) { 
                                	
                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
                                		if (cell2.getStringCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		data.add(cell2.getStringCellValue());
		                            	}	
		                            } else {                                       	
		                            	if (cell2.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                                    data.add((int) cell2.getNumericCellValue());
		                            	}
		                            }
                                }	
                                
                                if (cell2.getColumnIndex() == 17) { 
                                	
                                	if (cell2.getCellTypeEnum() != NUMERIC) { 
                                		if (cell2.getStringCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                            		data.add(cell2.getStringCellValue());
		                            	}	
		                            } else {                                       	
		                            	if (cell2.getDateCellValue() == null) {
		                            		data.add("");
		                            	} else {
		                                    data.add((int) cell2.getNumericCellValue());
		                            	}
		                            }
                                	data.add(cell2.getRowIndex()); // добавляем в список номер строки
                                }	
			                    break;
				        	}		        	
			        	}
			        }   	
				    
	                    if (nRowYellow == true && b == true && nRowRed == false) {
                			
                            if (data.size() == 0) {
                            	dateYellow.add(0);
                            } else {
                            	dateYellow.add((data.size()-19)/19);
                            }  
                            
                			monthToFinish.add(0, data.get(data.size()-9));
                			monthToFinish.add(1, data.get(data.size()-7));
                			monthToFinish.add(2, data.get(data.size()-6));
                			arrayMonthToFinish.add(monthToFinish);	                    			
                			nRowYellow = false;
                			nRowRedWithoutYellow = true;
                		} 
                		
                		if (whiteСell.size() != 0) {
                        	if (nRowRed == true && b == true && 
                        	((int) whiteСell.get(whiteСell.size() - 1) != (int) (data.size()-19)/19)) {
                        		
						      	if (data.size() == 0) {
						      		dateRed.add(0);
						      	} else {
						      		dateRed.add((int)(data.size()-19)/19);
						      	}
						      	
                    			weekToFinish.add(data.get(data.size()-10));
                    			weekToFinish.add(data.get(data.size()-8));
                    			weekToFinish.add(data.get(data.size()-7));
                				nRowRed = false;	                                        			
                			}   
                		} else {
                			if (nRowRed == true && b == true) {
                        		
						      	if (data.size() == 0) {
						      		dateRed.add(0);
						      	} else {
						      		dateRed.add((int)(data.size()-19)/19);
						      	}
						      	
                    			weekToFinish.add(data.get(data.size()-10));
                    			weekToFinish.add(data.get(data.size()-8));
                    			weekToFinish.add(data.get(data.size()-7));
                				nRowRed = false;	                                        			
                			}
                		}
            			nRowRedWithoutYellow = false;    
				    }    	
				    }
				    
				    if (fieldSub.isSelected() == true) {
				    	
				    	for (int i = 0; i < data.size(); i++) {
				    		jointUpload.add(data.get(i));
				    	}
				    }
				    
				    HashSet set = new HashSet(dateGreen);
				    dateGreen.clear();
				    dateGreen = new ArrayList<Integer>(set);	    
				    
					int cl = 19;
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

				    dateGreen.sort(null);
		            dateYellow.sort(null);
		            dateRed.sort(null);
		            
				    JTable table1 = new JTable(dm1) {				      
				        private static final long serialVersionUID = 1L;
				        
				        // кнопку редактирования изменять можно
		                public boolean isCellEditable(int row, int column) {  
		                	
		                	if (column != 18) {
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
			                for (int i = 0; i < dateGreen.size(); i++) {
			                	if (row == ((int) dateGreen.get(i)) && column == 10) {
			                		//c.setBackground(new java.awt.Color(174, 212, 149));
			                		c.setBackground(Color.white);
			                	}
			                }
			                
			                for (int i = 0; i < whiteСell.size(); i++) {
			                	if (row == ((int) whiteСell.get(i)) && column == 10) {
			                		// c.setBackground(Color.white);
			                	}
			                } 				                
			                for (int i = 0; i < dateYellow.size(); i++) {
			                	if (row == ((int) dateYellow.get(i)) && column == 10) {
			                		c.setBackground(new java.awt.Color(247, 239, 162));
			                	}
			                }  				                
			                for (int i = 0; i < dateRed.size(); i++) {
			                	if (row == ((int) dateRed.get(i)) && column == 10) {
			                		c.setBackground(new java.awt.Color(213, 92, 95));
			                	}
			                }			                
			                return c;	               	                
			            }
				    };				    
				    
				    JTableHeader th1 = table1.getTableHeader();
		        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));     
		        	th1.setPreferredSize(new Dimension(100, 120)); 
		        	
		        	// горизонтальная прокрутка заголовков
		        	table1.getTableHeader().setPreferredSize(new Dimension(10000,120));
		        	
		            // кнопка "редактировать"    
		            table1.getColumn(" ").setCellRenderer(new ButtonRendererDC (frame));
		            table1.getColumn(" ").setCellEditor(new ButtonEditorDC (new JCheckBox(), frame));
		            table1.getColumnModel().getColumn(18).setPreferredWidth(30);
		            // кнопка "редактировать"
		        			        	
		        	table1.setPreferredScrollableViewportSize(table.getPreferredSize());
		            table1.changeSelection(0, 0, false, false);
		            JScrollPane scrollPane1 = new JScrollPane(table1);
		            getContentPane().add(scrollPane1);		            
		        	
		        	table1.getColumnModel().getColumn(0).setPreferredWidth(35);
		        	
		        	table1.getColumnModel().getColumn(1).setPreferredWidth(140); 
		        	table1.getColumnModel().getColumn(1).setMaxWidth(140);
		        	table1.getColumnModel().getColumn(1).setMinWidth(140);
		        	
		        	table1.getColumnModel().getColumn(2).setPreferredWidth(170);  
		        	table1.getColumnModel().getColumn(2).setMaxWidth(170);
		        	table1.getColumnModel().getColumn(2).setMinWidth(170);
		        	
		        	table1.getColumnModel().getColumn(3).setPreferredWidth(170);
		        	table1.getColumnModel().getColumn(3).setMaxWidth(170);
		        	table1.getColumnModel().getColumn(3).setMinWidth(170);
		        	
		        	table1.getColumnModel().getColumn(4).setPreferredWidth(170); 
		        	table1.getColumnModel().getColumn(4).setMaxWidth(170);
		        	table1.getColumnModel().getColumn(4).setMinWidth(170);
		        	
		        	table1.getColumnModel().getColumn(5).setPreferredWidth(180); 
		        	table1.getColumnModel().getColumn(5).setMaxWidth(180);
		        	table1.getColumnModel().getColumn(5).setMinWidth(180);
		        	
					table1.getColumnModel().getColumn(6).setPreferredWidth(100);
					table1.getColumnModel().getColumn(7).setPreferredWidth(100);
					
					table1.getColumnModel().getColumn(8).setPreferredWidth(70);       	
					table1.getColumnModel().getColumn(9).setPreferredWidth(150);	
					
					table1.getColumnModel().getColumn(10).setPreferredWidth(170); 
		        	table1.getColumnModel().getColumn(10).setMaxWidth(170);
		        	table1.getColumnModel().getColumn(10).setMinWidth(170);
		        	
					table1.getColumnModel().getColumn(11).setPreferredWidth(110);
					table1.getColumnModel().getColumn(12).setPreferredWidth(150);
					table1.getColumnModel().getColumn(13).setPreferredWidth(110);
					table1.getColumnModel().getColumn(14).setPreferredWidth(150);
					table1.getColumnModel().getColumn(15).setPreferredWidth(155);
					table1.getColumnModel().getColumn(16).setPreferredWidth(155);
					table1.getColumnModel().getColumn(17).setPreferredWidth(150);
					// table1.getColumnModel().getColumn(18).setPreferredWidth(30); // кнопка редактировать
									
					table1.setRowHeight(25);
					
					for (int i = 0; i <= mainHeaders.length - 2; i++) {
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        	}
		            
				    workbook.close();
				    
				    table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
				    
				    JPanel panel = new JPanel(new BorderLayout(10, 10));
				    JPanel panelTop = new JPanel();
				    JPanel panelBt1 = new JPanel(new GridBagLayout());
				    JPanel panel2 = new JPanel(new BorderLayout(0, 0));				   
				    JPanel panel1 = new JPanel(new GridBagLayout());
				    GridBagConstraints c = new GridBagConstraints();
				    
		            three.setPreferredSize(new Dimension(0, 10));
		            three.setMaximumSize(new Dimension(0, 10));
		            three.setMinimumSize(new Dimension(0, 10));
				    
				    panelTop.setPreferredSize(new Dimension(0, 1));
				    panelTop.setMaximumSize(new Dimension(0, 1));
				    panelTop.setMinimumSize(new Dimension(0, 1));
				   
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
                	appointment.add(0, "Все назначения");    
                	
                	if (fieldSub.isSelected() == true) {
                		new SelectingColumnsDC().selecting(jointUpload, appointment);
				    } else {
				    	jointUpload.clear();
	        			new SelectingColumnsDC().selecting(copyData, appointment);	        			
				    }                	
                	appointment.remove(0);	        			        		
	        		
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
    public static void main(String[] args) throws IOException, ParseException {    	
       new DeviceAndConsumables().start(frame, 0);
    }
}

//создание кнопки "редактировать"
class ButtonRendererDC extends JButton implements TableCellRenderer {
	
	JFrame frame;
	int i = 0;
	
	public ButtonRendererDC(JFrame frame) {
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
	
	     if (!isSelected) {   	
	     	setForeground(table.getSelectionForeground());
	        setBackground(table.getSelectionBackground());
	     }
	     
	     setBackground(Color.white);
	     setIcon(pencil);
	     return this;
	}
}

//кнопка "редактировать"
class ButtonEditorDC extends DefaultCellEditor {
	
	public JButton button;
	String label = "";
	JFrame frame;
	public boolean isPushed;
	
	 public ButtonEditorDC(JCheckBox checkBox, JFrame frame) {
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
	     	 frame.dispose();
	     	 frame.setVisible(true);
	     	 frame.revalidate();
	     	 button.setForeground(table.getSelectionForeground());
	         button.setBackground(table.getSelectionBackground());
	     }
	     
	     button.setBackground(Color.white);
	     button.setIcon(pencil);
	     label = (value == null) ? "" : value.toString();
	
	     isPushed = true;
	     TableModel tm = table.getModel();
	     String[] inputValue = new String[19];
	     
	     for (int i = 0; i < inputValue.length; i++) {
	         inputValue[i] = (String) tm.getValueAt(row, i);
	     }
	     
	     int currentRow = Integer.parseInt(inputValue[inputValue.length-1]);
	     
	     try {       				
	     	new EditButtonDC().windowDataChange(new InputEditing().inputValues(currentRow, 0, 19));	     	
		 } catch (IOException e) {
			e.printStackTrace();
		 }	
	     
	     isPushed = true;
	     return button;
	 }

	 public Object getCellEditorValue() {		 
		 label = "";
	     isPushed = false;
	     /*frame.dispose();
	 	 frame.setVisible(true);
	 	 frame.revalidate();*/
	     return label;
	 }

	 public boolean stopCellEditing() {
		 
	     isPushed = true;
	     return super.stopCellEditing();
	 }
}