package net.codejava;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.Label;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class SelectingColumnsCar{
	
	Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
	static String fileDictName = "";
	
    public void selecting(ArrayList dataSortInput) throws IOException {
    	
    	// копируем данные, чтобы можно было использовать кнопку множество раз
    	ArrayList dataInput = new ArrayList(dataSortInput);
    	/*
    	ImageIcon liderIcon = new ImageIcon(new SelectingColumnsCar().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();
        */
        String[] items = {"Все столбцы",
        		"№", "Марка, модель", "VIN", "Регистрационный номер", "Год выпуска", "Пробег", "ПТС", "СТС",
    		    "Страховая компания", "№ полиса", "Срок действия", "Страховая компания", "№ полиса",
    		    "Срок действия", "Техническое состояние", "Форма собственности", "Владелец оборудавания", 
    		    "Местонахождение", "Ответственный владелец", "Примечание"};
        
        String[] names = {"Все столбцы", "№", "Марка, модель", "VIN", "Регистрационный номер", "Год выпуска", "Пробег", "ПТС", "СТС",
    		    "Страховая компания (КАСКО)", "№ полиса (КАСКО)", "Срок действия (КАСКО)", "Страховая компания (ОСАГО)", 
    		    "№ полиса (ОСАГО)", "Срок действия (ОСАГО)", "Техническое состояние", 
    		    "Форма собственности", "Владелец оборудавания", "Местонахождение", "Ответственный владелец", "Примечание"};
        
        JPanel jPanelL1 = new JPanel(new GridLayout(names.length + 1, 3, 5, 0));
    	jPanelL1.add(new JLabel("Выберите нужные столбцы: ", SwingConstants.LEFT)).setFont(bigFontTR);
    	
        JCheckBox l1[] = new JCheckBox[names.length];
        Boolean[] values = new Boolean[names.length];        
        values[0] = Boolean.FALSE;
        
        for (int i = 1; i < values.length; i++) {
            values[i] = Boolean.TRUE;
        }
        
        for (int jV1 = 0; jV1 < names.length; jV1++)
            l1[jV1] = new JCheckBox(names[jV1], values[jV1]);
        
        for (int i = 0; i < names.length; i++) {
            new WorkingCheckBox().scaleCheckBoxIcon(l1[i], 21);
            jPanelL1.add(l1[i]).setFont(bigFontTR);
            
            l1[i].addItemListener(new ItemListener() {
                public void itemStateChanged(ItemEvent e) {
                	for (int i = 1; i < names.length; i++) {
                        if (l1[i].isSelected() == true) {
                        	l1[0].setSelected(false);
                        }
                    }                	
                }
            });
        }  
        
        JPanel panel1 = new JPanel();
        jPanelL1.setPreferredSize( new Dimension(510, 730) );
        panel1.add(new Label(" "), BorderLayout.WEST);
        panel1.add(jPanelL1, BorderLayout.CENTER);
        panel1.add(new Label(" "), BorderLayout.EAST);
        
        JButton buttonAdd = new JButton("Выгрузить в excel");
        buttonAdd.setPreferredSize(new Dimension(200, 30));
        buttonAdd.setFont(bigFontTR);

        JPanel panelBt = new JPanel();
        panelBt.setPreferredSize(new Dimension(100, 30));
        panelBt.add(buttonAdd);

        JPanel jPanelSelecting = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();

        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 0;        
        c.weightx = 1;
        c.weighty = 0.9;
        c.fill = GridBagConstraints.BOTH;
        jPanelSelecting.add(panel1, c);
        
        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 1;        
        c.weightx = 1;
        c.weighty = 1;
        c.fill = GridBagConstraints.BOTH;
        
        jPanelSelecting.add(panelBt, c);
        jPanelSelecting.setPreferredSize(new Dimension(610, 790));
        
        JFrame frameSC = new JFrame();     
        frameSC.add(jPanelSelecting);
        // frameSC.setIconImage(image);
        frameSC.pack();
        frameSC.setLocationRelativeTo(null);
        frameSC.setLocation(0, 0);
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
	       			 InputStream inputStream = new FileInputStream("Reestr.xls");
	    	         Workbook workbook = new HSSFWorkbook(inputStream);
	    	         Sheet currentSheet = workbook.getSheetAt(1);
                    
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
                    	 for (int i = 0; i < dataInput.size(); i++) {
                    		 if (i != 0 && i % items.length != 0) {
                    			 dataSort3.add(dataInput.get(i - 1));
                    		 }                  		
                    	 }                    	 
                    } else {
                    	for (int i = 0; i < dataInput.size(); i++) {
                   		    if (i != 0 && i % items.length != 0) {
                   		    	ar.add(dataInput.get(i - 1));
                   		    }                  		
                   	    }
                    	for (int i = 0; i < ar.size(); i++) {
                    		double d = (double)nColumns.get(c3) / (double)(i - ((items.length - 1) * c2) + 1);
                    		if (d  == 1.0) {
                   			 	dataSort3.add(ar.get(i));
                   			 	if (nColumns.size() != c3 + 1) {
                   			 		c3++;
                   			 	} else {
                   			 		c3 = 0;
                   			 	}
                    		}  
                    		if (c != ( items.length - 2) ) {
                   			   c++;                   			
                    		} else {
                    			c2++;
                    			c = 0;
                    			
                    		}
                    	}
                    }
                    int nClms;
                    if (nColumns.get(0) == 0) {
                    	 nClms = (items.length - 1);
                    } else {
                    	nClms = nColumns.size();
                    }
                    int nRow = dataSort3.size()/nClms;
                    
                    
                    InputStream inputStream2 = new FileInputStream("Reestr.xls");
                    HSSFWorkbook sh = new HSSFWorkbook(inputStream2);
                    HSSFSheet worksheet = sh.getSheetAt(1);
                    
                    HSSFCell cell = null;
                    int lastRow = 3;
                    HSSFRow row ;
                    
                    // удаляем лист с девайсами и автомобилями
                    sh.removeSheetAt(0);
                    sh.removeSheetAt(1);
                    
                    for (int i = 3; i <  100000; i++) {
                    	if (worksheet.getRow(i) != null) {
                    		worksheet.removeRow(worksheet.getRow(i));
                    	}	
                    }
                    
                    // Создать стиль шрифта
            		HSSFFont fontHeader = sh.createFont();
            		fontHeader.setFontHeightInPoints ( (short) 11); 
            		fontHeader.setBold(true);
            		// Создать стиль шрифта
            		HSSFFont font = sh.createFont();
            		font.setFontHeightInPoints ( (short) 11);
            		
            		// Создать стиль ячейки
            		HSSFCellStyle styleHeader = sh.createCellStyle();
            		// Установить границу
            		styleHeader.setFont(fontHeader);
            		styleHeader.setWrapText(true);
            		styleHeader.setAlignment(HorizontalAlignment.CENTER);
            		styleHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            		styleHeader.setBorderTop( BorderStyle.THIN );
            		styleHeader.setBorderBottom( BorderStyle.THIN );
            		styleHeader.setBorderLeft( BorderStyle.THIN );
            		styleHeader.setBorderRight( BorderStyle.THIN );
                    
                    // Создать стиль ячейки
            		HSSFCellStyle style = sh.createCellStyle();
            		// Установить границу
            	    style.setFont(font);
                    style.setWrapText(true);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    style.setBorderTop( BorderStyle.THIN );
                    style.setBorderBottom( BorderStyle.THIN );
                    style.setBorderLeft( BorderStyle.THIN );
                    style.setBorderRight( BorderStyle.THIN );
                    
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
	                        for (int i = 0; i < dataSort3.size(); i++) {
	                        	if (nL < nColumns.size() && nColumns.get(nL) == i || nColumns.get(0) == 0) { 
		                            cell = row.createCell(nL);
		                            row.createCell(nL).setCellValue((String) dataSort3.get(dataIndex));
		                            row.getCell(nL).setCellStyle(style);
		                            dataIndex++;
		                            if (nColumns.get(0) != 0) {
	                                	nL ++;
	                                }
	                        	} 
	                        }
	                        
                        } else {
	                        for (int i = 0; i <= (items.length - 2); i++) {                        	
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
	                    for (int i = 0; i <= items.length - 1; i++) {
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
                    
                    // объединение ячеек
                    Row newRow0 = worksheet.createRow(1);
                    Row newRow = worksheet.createRow(2);
                    
                    int nCD = 0;
                    String merge [] = {"КАСКО", "ОСАГО"};
                    int m = 0;
                   
                    if (nColumns.get(0) != 0) {
                    	// убираем все соединения
                    	for (int j = 0; j <= ( items.length - 1 ); j++) {
	                    	for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
	                			worksheet.removeMergedRegion(i);
	                		}
                    	}	
                    	System.out.println( items[10] + " " + items[13] );
                    	for (int i = 1; i <= items.length - 1; i++) {	
	                    	if (nCD < nColumns.size() && nColumns.get(nCD).equals(i)) {
	                    		if ( nColumns.get(nCD).equals(9) | nColumns.get(nCD).equals(12) |
	                    			 nColumns.get(nCD).equals(10) | nColumns.get(nCD).equals(13) |
	                    			 nColumns.get(nCD).equals(11) | nColumns.get(nCD).equals(14) ) {
	                    			int countMerge = 0;
	                    			int indexM = 0;
	                    			
	                    			if ( (nColumns.get(nCD).equals(9) && nColumns.get(nCD + 1).equals(10) ) | 
	                    				 (nColumns.get(nCD).equals(9) && nColumns.get(nCD + 1).equals(11) ) |
	                    				 (nColumns.get(nCD).equals(10) && nColumns.get(nCD + 1).equals(11))) {
	                    				countMerge = 1;
	                    				indexM = 0;
	                    			}
	                    			if ( nColumns.get(nCD).equals(9) && nColumns.get(nCD + 1).equals(10) && nColumns.get(nCD + 2).equals(11) ) {
	                    				countMerge = 2;
	                    				indexM = 0;
	                    			}
	                    			if ( (nColumns.get(nCD).equals(12) && nColumns.get(nCD + 1).equals(13) ) |
                    					 (nColumns.get(nCD).equals(12) && nColumns.get(nCD + 1).equals(14) ) |
                    					 (nColumns.get(nCD).equals(13) && nColumns.get(nCD + 1).equals(14))) {
	                    				countMerge = 1;
	                    				indexM = 1;
	                    			}
	                    			if ( nColumns.get(nCD).equals(12) && nColumns.get(nCD + 1).equals(13) && nColumns.get(nCD + 2).equals(14) ) {
	                    				countMerge = 2;
	                    				indexM = 1;
	                    			}
	                    			 
	                    			int cMerged = nCD;
	                    			System.out.println("Out: " + ( cMerged + countMerge ) + " " + cMerged);	 
	                    			if ( countMerge == 0) {
                    					newRow0.createCell(cMerged).setCellValue("");	
                        				newRow0.createCell(cMerged).setCellValue( merge[indexM] ); 
                    					newRow0.getCell(cMerged).setCellStyle(styleHeader);	
                         				
	                    				newRow.createCell(cMerged).setCellValue("");	
	                    				newRow.createCell(cMerged).setCellValue( items[i] );
	                    				newRow.getCell(cMerged).setCellStyle(styleHeader);
	                    				
	                    			} else {
	                    				for (int index = 0; index <= countMerge; index ++ ) {
	                    					newRow0.createCell(cMerged + index).setCellValue("");	
	                        				newRow0.createCell(cMerged + index).setCellValue( merge[indexM] ); 
	                    					newRow0.getCell(cMerged + index).setCellStyle(styleHeader);	
	                    				}                 				                  				
		                    			CellRangeAddress cra3 = new CellRangeAddress(1, 1, cMerged, cMerged + countMerge );
		                    			worksheet.addMergedRegion(cra3);	                    				                    			
		                    			
		                    			for (int index = 0; index <= countMerge; index ++ ) {	                    				
		                    				newRow.createCell(cMerged  + index).setCellValue("");	
		                    				newRow.createCell(cMerged  + index).setCellValue( items[i + index] );
		                    				newRow.getCell(cMerged  + index).setCellStyle(styleHeader);
		                    			}
	                    			}
		                    		nCD += countMerge + 1;	                 		
	                    		}
	                    		else {
	                    			CellRangeAddress cra = new CellRangeAddress(1, 2, nCD, nCD);
	                    			worksheet.addMergedRegion(cra);
	                    			newRow.createCell(nCD).setCellValue("");
		                    		newRow.getCell(nCD).setCellStyle(styleHeader);
		                    		newRow0.createCell(nCD).setCellValue(items[i]);
		                    		newRow0.getCell(nCD).setCellStyle(styleHeader);
		                    		nCD++;
	                    		}                 		
	                    	}
	                    }
	                 } else {	                	 
	                	 for (int i = 1; i <= items.length - 1; i++) {
	                		 switch (i) {
	                		 	case 9:
	                		 		newRow0.createCell( i - 1 ).setCellValue(merge[m]);
		                		 	newRow0.getCell( i - 1 ).setCellStyle(styleHeader);               			
		                    		newRow.createCell( i - 1 ).setCellValue(items[i]);
		                    		newRow.getCell( i - 1 ).setCellStyle(styleHeader);
		                    		m++;
		                    		break;
	                		 	case 12:
	                		 		newRow0.createCell( i - 1 ).setCellValue(merge[m]);
		                		 	newRow0.getCell( i - 1 ).setCellStyle(styleHeader);               			
		                    		newRow.createCell( i - 1 ).setCellValue(items[i]);
		                    		newRow.getCell( i - 1 ).setCellStyle(styleHeader);
		                    		m++;
		                    		break;
	                    		default:
	                    			newRow0.createCell( i - 1 ).setCellValue(items[i]);
		                    		newRow0.getCell( i - 1 ).setCellStyle(styleHeader);	                    			
			                    	newRow.createCell( i - 1 ).setCellValue(items[i]);
			                    	newRow.getCell( i - 1 ).setCellStyle(styleHeader);
	                    			break;
	                		 }	
	                	}	
	                }
                    int nC2 = 0;
                    for (int nColumn = 0; nColumn <= items.length; nColumn++) {
                    	if ((nColumns.get(nC2)-1) == nColumn) {
	                    	int width = (int) worksheet.getColumnWidthInPixels(nColumn);
	                    	width = width * 30;
	                    	worksheet.setColumnWidth(nC2, width);	                    	
	                    	nC2++;
                    	} 
                    }
                    
                    int width = (int) worksheet.getColumnWidthInPixels(11);
                    width = width * 30;
                    for (int nColumn = nColumns.size() - 1; nColumn <= items.length + 1; nColumn++) {
                    	worksheet.setColumnWidth(nColumn-1, width);
                    }	

                    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
                    String filenameOut = "База МТЦ (Автомобили) от " + timeStamp  + ".xls";
                    
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
