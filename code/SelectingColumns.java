package net.codejava;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.Label;
import java.awt.List;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.ListCellRenderer;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdesktop.swingx.JXDatePicker;

public class SelectingColumns {
	Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 15);
	static String fileDictName = "";
	
    public void selecting(String[] inputValues, ArrayList dataSort) throws IOException {
    	
        JPanel jPanelL1 = new JPanel(new GridLayout(18, 1, 0, 10));
    	jPanelL1.add(new JLabel("Выберите нужные столбцы: ", SwingConstants.LEFT)).setFont(bigFontTR);
        String[] items = {"Все столбцы",
        		"Фамилия", "Имя", "Отчество", "Организация", "Область аттестации",
                "Категория", "Начало действия", "Окончание действия", "Номер удостоверения", "Контакты",
                "Дополнительная информация", "Цена"};

        JCheckBox l1[] = new JCheckBox[items.length];
        Boolean[] values = new Boolean[items.length];
        
        
        values[0] = Boolean.TRUE;
        for (int i = 1; i < values.length; i++) {
            values[i] = Boolean.FALSE;
        }
        
        CheckComboStore[] stores = new CheckComboStore[items.length];
        for (int jV = 0; jV < items.length; jV++)
            l1[jV] = new JCheckBox(items[jV], values[jV]);
        
        for (int i = 0; i < items.length; i++) {
            // new Main().scaleCheckBoxIcon(l1[i], 21);
            jPanelL1.add(l1[i]).setFont(bigFontTR);
            l1[i].addItemListener(new ItemListener() {
                public void itemStateChanged(ItemEvent e) {
                	for (int i = 1; i < items.length; i++) {
                        if (l1[i].isSelected() == true) {
                        	l1[0].setSelected(false);
                        }
                    }
                	
                }
            });
        }
        
        JPanel panel1 = new JPanel();
        jPanelL1.setPreferredSize(new Dimension(400, 650));
        panel1.add(new Label(" "), BorderLayout.WEST);
        panel1.add(jPanelL1, BorderLayout.CENTER);
        panel1.add(new Label(" "), BorderLayout.EAST);
        
        JButton buttonAdd = new JButton("Выгрузить в excel");
        buttonAdd.setPreferredSize(new Dimension(300, 30));
        buttonAdd.setFont(bigFontTR);

        JPanel panelNew = new JPanel(new BorderLayout(0, 0));
        panelNew.setPreferredSize(new Dimension(500, 520));
        panelNew.add(new Label(" "), BorderLayout.WEST);
        panelNew.add(panel1, BorderLayout.CENTER);
        panelNew.add(new Label(" "), BorderLayout.EAST);

        JPanel panelBt = new JPanel(new GridLayout(1, 1, 0, 0));
        panelBt.setPreferredSize(new Dimension(0, 30));
        panelBt.add(new Label(" "), BorderLayout.WEST);
        panelBt.add(buttonAdd, BorderLayout.CENTER);
        panelBt.add(new Label(" "), BorderLayout.EAST);

        
        JPanel jPanelAdd = new JPanel(new GridBagLayout());
        GridBagConstraints c2 = new GridBagConstraints();
        jPanelAdd.setPreferredSize(new Dimension(550, 570));
        c2.fill = GridBagConstraints.VERTICAL;
        c2.gridx = 1;
        c2.gridy = 0;
        
        c2.weightx = 0.5;
        c2.weighty = 2;
        c2.fill = GridBagConstraints.BOTH;
        jPanelAdd.add(panelNew, c2);
        
        c2.fill = GridBagConstraints.VERTICAL;
        c2.gridx = 1;
        c2.gridy = 1;
        
        c2.weightx = 0.5;
        c2.weighty = 0.07;
        c2.fill = GridBagConstraints.BOTH;
        
        jPanelAdd.add(panelBt, c2);
           
        JFrame frameSC = new JFrame();
        /*
        ImageIcon liderIcon = new ImageIcon(new SelectingColumns().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();
        frameSC.setIconImage(image);
        */
        
        frameSC.add(new JScrollPane(jPanelAdd));
        frameSC.setSize(850, 900);
        frameSC.pack();
        // frameSC.setLocationRelativeTo(null);
        frameSC.setVisible(true);
       
        buttonAdd.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {     
            	frameSC.setVisible(false);
            	 ArrayList<Integer> nColumns = new ArrayList<Integer>();
            	 int nL = 0;
            	// узнаем количество выделенных значений
                int count = 0;
                for (int i = 0; i < items.length; i++) {
                    if (l1[i].isSelected() == true) {
                        count++;
                    }
                }

                // записываем выделенные значения
                for (int i = 0; i < items.length; i++) {
                    if (l1[i].isSelected() == true) {
                        nColumns.add(i);
                    }
                }
                
                ArrayList data = new ArrayList();
                
                try {
	       			 InputStream inputStream = new FileInputStream("Database.xls");
	    	         Workbook workbook = new HSSFWorkbook(inputStream);
	    	         Sheet currentSheet = workbook.getSheetAt(0);
                    
                    Iterator<Row> rowIterator = currentSheet.iterator();
                    rowIterator.next(); // skip the header row
                    int last_name;
                    int ch = 0;
                    Row nextRow2 = rowIterator.next();
                    Iterator<Cell> cellIterator2 = nextRow2.cellIterator();
                    int nRows = 0;
                    ArrayList ar = new ArrayList();
                    int c = 0;
                    int c2 = 0;
                    int c3 = 0;
                    ArrayList dataSort3 = new ArrayList();
                    if (nColumns.get(0) == 0) {
                    	 for (int i = 0; i < dataSort.size(); i++) {
                    		 if (i != 0 && i % 13 != 0) {
                    			 dataSort3.add(dataSort.get(i - 1));
                    		 }                  		
                    	 }
                    	 
                    } else {
                    	for (int i = 0; i < dataSort.size(); i++) {
                   		    if (i != 0 && i % 13 != 0) {
                   			 ar.add(dataSort.get(i - 1));
                   		 }                  		
                   	    }
                    	for (int i = 0; i < ar.size(); i++) {
                    		double d = (double)nColumns.get(c3) / (double)(i - (12 * c2) + 1);
                    		if (d  == 1.0) {
                   			 	dataSort3.add(ar.get(i));
                   			 	if (nColumns.size() != c3 + 1) {
                   			 		c3++;
                   			 	} else {
                   			 		c3 = 0;
                   			 	}
                    		}  
                    		if (c != 11) {
                   			   c++;                   			
                    		} else {
                    			c2++;
                    			c = 0;
                    			
                    		}
                    	}
                    }
                    int nClms;
                    if (nColumns.get(0) == 0) {
                    	 nClms = 12;
                    } else {
                    	nClms = nColumns.size();
                    }
                    int nRow = dataSort3.size()/nClms;
                    
                    InputStream inputStream2 = new FileInputStream("Database.xls");
                    HSSFWorkbook sh = new HSSFWorkbook(inputStream2);
                    HSSFSheet worksheet = sh.getSheetAt(0);
                    
                    HSSFCell cell = null;
                    int lastRow = 3;
                    HSSFRow row ;
                    
                    for (int i = 3; i <  100000; i++) {
                    	if (worksheet.getRow(i) != null) {
                    		worksheet.removeRow(worksheet.getRow(i));
                    	}	
                    }
                    
                    CellStyle style2 = sh.createCellStyle();
                    HSSFFont fontZ = sh.createFont();
                    fontZ.setBold(true);
                    style2.setWrapText(true);
                    style2.setFont(fontZ);
                    style2.setAlignment(HorizontalAlignment.CENTER);
                    style2.setVerticalAlignment(VerticalAlignment  .CENTER);
                    style2.setBorderTop(BorderStyle.THIN);
                    style2.setBorderBottom(BorderStyle.THIN);
                    style2.setBorderLeft(BorderStyle.THIN);
                    style2.setBorderRight(BorderStyle.THIN);
                    
                    // Создать стиль шрифта
            		HSSFFont font= sh.createFont();
            		font.setFontName("Calibri");
            		font.setFontHeightInPoints ((short) 11); 

            		// Создать стиль ячейки
            		HSSFCellStyle style = sh.createCellStyle();

            		// Установить границу
            	    style.setFont(font);// Установить шрифт
                    style.setWrapText(true);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment  .CENTER);
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                    
                    int dataIndex = 0;
                    int n = 1;
                    for (int j = 0; j < nRow; j++) {
                    	
                    	if (nColumns.size() == 1 && nColumns.get(0) == 0) {
                       	 	nL = 0;
                    	} else {
                    		if (nColumns.size() == 1) { 
                    			nL = 0;
                    		} else {
                    			nL = 0;
                    		}
                    	}
                        row = worksheet.createRow(lastRow);
                        
                        if (nColumns.get(0) != 0) {
	                        for (int i = 1; i <= 5; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(nL+1);
		                            row.createCell(nL+1).setCellValue((String) dataSort3.get(dataIndex));
		                            if (i == 4) {
		                            	row.getCell(nL + 1).setCellStyle(style);		                            	 
		                            } else {
		                            	row.getCell(nL + 1).setCellStyle(style);		                            	
		                            }
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }
		                            
	                        	}    
	                        }
	                        for (int i = 6; i <= 6; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(nL+1);
		                            row.createCell(nL+1).setCellValue((Integer) dataSort3.get(dataIndex));
		                            row.getCell(nL + 1).setCellStyle(style);
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }		                           
	                        	}    
	                        }
	                        for (int i = 7; i <= 11; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(nL+1);
		                            row.createCell(nL+1).setCellValue((String) dataSort3.get(dataIndex).toString());
		                            if (i == 11) {
		                            	row.getCell(nL + 1).setCellStyle(style);
		                            } else {
		                            	row.getCell(nL + 1).setCellStyle(style);
		                            }
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }	
	                        	}    
	                        }
	                        
	                        for (int i = 12; i <= 12; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(nL + 1);
		                            row.createCell(nL + 1).setCellValue((String) dataSort3.get(dataIndex));
		                            row.getCell(nL + 1).setCellStyle(style);
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }
	                        	} 
	                        }
	                        
                        } else {
                        	for (int i = 1; i <= 5; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(i);
		                            row.createCell(i).setCellValue((String) dataSort3.get(dataIndex));
		                            if (i == 4) {
		                            	row.getCell(i).setCellStyle(style);
		                            } else {
		                            	row.getCell(i).setCellStyle(style);
		                            }
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }
	                        	}    
	                        }
	                        for (int i = 6; i <= 6; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(i);
		                            row.createCell(i).setCellValue(dataSort3.get(dataIndex).toString());
		                            row.getCell(i).setCellStyle(style);
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }	
	                        	}    
	                        }
	                        for (int i = 7; i <= 11; i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(i);
		                            row.createCell(i).setCellValue((String) dataSort3.get(dataIndex));
		                            if (i == 11) {
		                            	row.getCell(i).setCellStyle(style);
		                            } else {
		                            	row.getCell(i).setCellStyle(style);
		                            }
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }	
	                        	}    
	                        }
	                        
	                        for (int i = 12; i <= 12; i++) {                        	
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(i);
		                            row.createCell(i).setCellValue(dataSort3.get(dataIndex).toString());
		                            row.getCell(i).setCellStyle(style);
		                            
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }
	                        	}     
	                        }
                        }
                        lastRow++;
                    }
                   
                    ArrayList<Integer> nColumnsDel = new ArrayList<Integer>();
                    int nC = 0;
                    if (nColumns.get(0) != 0) {
	                    for (int i = 1; i <= 12; i++) {
	                    	if (nC < nColumns.size() && nColumns.get(nC).equals(i)) {
	                    		 nC++;
	                    	} else {
	                    		nColumnsDel.add(i);
	                    	}
	                    }
	                    
                    	for (int i = 0; i < nColumnsDel.size(); i++) {
		                    for (int rId = 1; rId <= 2; rId++) {
		                        Row rowC = worksheet.getRow(rId);
		                        for (int cID = nColumnsDel.get(i); cID < rowC.getLastCellNum(); cID++) {
		                            Cell cOld = rowC.getCell(cID);
		                            
		                            if (cOld != null) {
		                                rowC.removeCell(cOld);
		                            }
		                            Cell cNext = rowC.getCell(cID);
		                            if (cNext != null) {
		                                Cell cNew = rowC.createCell(cID, cNext.getCellType());
		                                cloneCell(cNew, cNext);
		                                worksheet.setColumnWidth(cID, worksheet.getColumnWidth(cID));
		                            }
		                        }
		                    }
                    	}   
                    }
                    nColumns.add(100);
                    worksheet.removeRow(worksheet.getRow(1));                   
                    worksheet.removeRow(worksheet.getRow(2));
                    Row newRow0 = worksheet.createRow(1);
                    Row newRow = worksheet.createRow(2);
                    
                    int nCD = 0;
                    String merge [] = {"ФИО Эксперта"};
                    int m = 0;
                    int m2 = 0;
                   
                    ArrayList<Integer> massM2 = new ArrayList<Integer>();
                    ArrayList<Integer> massM3 = new ArrayList<Integer>();
                    if (nColumns.get(0) != 0) {
                    	for (int j = 0; j <= 3; j++) {
	                    	for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
	                			worksheet.removeMergedRegion(i);
	                		}
                    	}	
                    	
                    	CellRangeAddress cra = new CellRangeAddress(1, 1, 1, 3);
                    	if (nColumns.get(0).equals(1) && nColumns.get(1).equals(2) && nColumns.get(2).equals(3)) {
                    			worksheet.addMergedRegion(cra);
                    			newRow0.createCell(1).setCellValue(merge[0]);
	                		 	newRow0.getCell(1).setCellStyle(style2);
	                		 	newRow0.createCell(2).setCellValue("");
	                		 	newRow0.getCell(2).setCellStyle(style2);
	                		 	newRow0.createCell(3).setCellValue("");
	                		 	newRow0.getCell(3).setCellStyle(style2);
                    	                    	
		                        for (int i = 1; i <= 12; i++) {
	                		 	// for (int i = 1; i <= 16; i++) {
		                    	if (nColumns.get(nCD).equals(i)) {
		                    		if ((i >= 4 && i <= 11) || i == 12 ) {
		                    			CellRangeAddress cra3 = new CellRangeAddress(1, 2, nCD+1, nCD+1);
		                    			worksheet.addMergedRegion(cra3);
		                    			newRow.createCell(nCD+1).setCellValue("");
			                    		newRow.getCell(nCD+1).setCellStyle(style2);
		                    		}
		                    		
		                    		if ((i >= 4 && i <= 11) || i == 12 ) {
		                    			
		                    			newRow0.createCell(nCD+1).setCellValue(items[i]);
			                    		newRow0.getCell(nCD+1).setCellStyle(style2);
			                    		
		                    		} else {
		                    			
			                    		newRow.createCell(nCD+1).setCellValue(items[i]);
			                    		newRow.getCell(nCD+1).setCellStyle(style2);
		                    		}	
		                    		nCD++;
		                    	}
		                    }	
		                    	

                    	} else {
                    		
                    		for (int i = 1; i <= 12; i++) {	
		                    	if (nCD < nColumns.size() && nColumns.get(nCD).equals(i)) {
		                    		if ((i >= 1 && i <= 11) || i == 12) {
		                    			CellRangeAddress cra3 = new CellRangeAddress(1, 2, nCD+1, nCD+1);
		                    			worksheet.addMergedRegion(cra3);
		                    			newRow.createCell(nCD+1).setCellValue("");
			                    		newRow.getCell(nCD+1).setCellStyle(style2);
		                    		}
		                    				
		                    			newRow0.createCell(nCD+1).setCellValue(items[i]);
			                    		newRow0.getCell(nCD+1).setCellStyle(style2);
			                    		
		                    		nCD++;
		                    	}
		                    }	                   				                    
		                    
                    	}
	                    
	                 } else {
	                	 
	                	 for (int i = 1; i <= 12; i++) {
	                		 if (i != 1 ) {
	                    		newRow0.createCell(i).setCellValue(items[i]);
	                    		newRow0.getCell(i).setCellStyle(style2);	                    			
		                    	newRow.createCell(i).setCellValue(items[i]);
		                    	newRow.getCell(i).setCellStyle(style2);
	                    	} else {	 
		                		 	newRow0.createCell(i).setCellValue(merge[m]);
		                		 	newRow0.getCell(i).setCellStyle(style2);               			
		                    		newRow.createCell(i).setCellValue(items[i]);
		                    		newRow.getCell(i).setCellStyle(style2);
		                    		m++;
	                    	}	
	                	}	
	                }
                    int nC2 = 0;
                    for (int nColumn = 1; nColumn <= 13; nColumn++) {
                    	if ((nColumns.get(nC2)-1) == nColumn) {
	                    	int width = (int) worksheet.getColumnWidthInPixels(nColumn);
	                    	width = width * 30;
	                    	worksheet.setColumnWidth(nC2, width);
	                    	
	                    	nC2++;
                    	} 
                    }
                    
                    int width = (int) worksheet.getColumnWidthInPixels(11);
                    width = width * 30;
                    for (int nColumn = nColumns.size(); nColumn <= 14; nColumn++) {
                    	worksheet.setColumnWidth(nColumn-1, width);
                    }	
                    worksheet.setColumnWidth(nC2, 360);

                    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
                    String filenameOut = "Список экспертов от " + timeStamp  + ".xls";
                    // String filenameOut = "N:\\Отдел ЭПБ\\17. Программы\\" + "Список экспертов <" + timeStamp + ">" + ".xlsx";
                    
                    JFileChooser fileChooser = new JFileChooser();
                	fileChooser.setCurrentDirectory(new File("N:\\Программы\\"));
                	FileNameExtensionFilter filter = new FileNameExtensionFilter("Files", ".xls");
                	fileChooser.addChoosableFileFilter(filter);
                	fileChooser.setAcceptAllFileFilterUsed(false);
                	fileChooser.setDialogTitle("Save the dictionary file"); 
                	fileChooser.setSelectedFile(new File(filenameOut));
                	int userSelection = fileChooser.showSaveDialog(fileChooser);
                	if (userSelection == JFileChooser.APPROVE_OPTION) {
                	    fileDictName = fileChooser.getSelectedFile().getAbsolutePath();
                	}
                	
                	File file = new File(fileDictName);
                	if (file.exists() == false) {
                		// FileOutputStream output_file =new FileOutputStream(filenameOut);
                		FileOutputStream output_file =new FileOutputStream(file);
                	
                		sh.write(output_file);
                		output_file.close();
                		sh.close();
                		workbook.close();
                		System.out.println("Is saved in Excel file.");
                	} else {
                        System.out.println("File already exist");
                    }
                } catch (IOException ex1) {
                    System.out.println("Error reading file!");
                    System.out.println(ex1);
                    ex1.printStackTrace();
                } 
                catch (NullPointerException n) {
                    System.out.println(n);
                }
            }
        });     
    }	
    
        private static void cloneCell( Cell cNew, Cell cOld ){
        cNew.setCellComment( cOld.getCellComment() );
        cNew.setCellStyle( cOld.getCellStyle() );
        cNew.setCellValue(1.7);
        switch ( cNew.getCellType() ){
            case Cell.CELL_TYPE_BOOLEAN:{
                cNew.setCellValue( cOld.getBooleanCellValue() );
                break;
            }
            case Cell.CELL_TYPE_NUMERIC:{
                cNew.setCellValue( cOld.getNumericCellValue() );
                break;
            }
            case Cell.CELL_TYPE_STRING:{
                cNew.setCellValue( cOld.getStringCellValue() );
                break;
            }
            case Cell.CELL_TYPE_ERROR:{
                cNew.setCellValue( cOld.getErrorCellValue() );
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                cNew.setCellFormula( cOld.getCellFormula() );
                break;
            }
        }

    }
   
}

