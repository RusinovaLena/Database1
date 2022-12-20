package net.codejava;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.Insets;
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
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import net.codejava.Cars;
import net.codejava.WorkingCheckBox;

public class SelectingColumnsTechnics implements ActionListener {
    
	Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
	static String fileDictName = "";
	
    public void selecting( ArrayList inputData, int columnCount ) throws IOException {   
    	
    	// копируем данные, чтобы можно было использовать кнопку множество раз
    	ArrayList data = new ArrayList(inputData); 	    	
    	// удаляем номера строк
    	for (int i = columnCount; i < data.size(); i += columnCount) {
    		data.remove(i);
    	}    	
    	// устанавливаем иконку для панели с выгрузкой
    	/*
    	ImageIcon liderIcon = new ImageIcon(new Cars().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();              
    	*/
        String[] items = {"Все столбцы",
        		"№", "Вид оборудования", "Наименование прибора", "Тип, марка, модель", 
        		"<html><left>Производитель, страна производства,основные технические характеристики", "Зав.№", "Год выпуска",
        		"Комплектность", "Документы", "Техническое состояние", "Форма собственности","Владелец оборудавания",
        		"Местонахождения и текущие обязательства", "Офис", "Инвентарный номер",
                "Примечание"}; 
        
        String[] nameColumn = {"Все столбцы",
        		"№", "Вид оборудования", "Наименование прибора", "Тип, марка, модель", 
        		"Производитель, страна производства,основные технические характеристики", "Зав.№", "Год выпуска",
        		"Комплектность", "Документы", "Техническое состояние", "Форма собственности","Владелец оборудавания",
        		"Местонахождения и текущие обязательства", "Офис", "Инвентарный номер",
                "Примечание"}; // 17
        
        JPanel panelOutload = new JPanel( new GridLayout(items.length + 1, 3, 5, -8) );
        
        panelOutload.add(new JLabel("Выберите нужные столбцы: ", SwingConstants.LEFT)).setFont(bigFontTR);
        
        JCheckBox l1[] = new JCheckBox[items.length];
        Boolean[] values = new Boolean[items.length];        
        values[0] = Boolean.TRUE;
        
        for (int i = 1; i < values.length; i++) {
            values[i] = Boolean.FALSE;
        }
        
        for (int jV1 = 0; jV1 < items.length; jV1++)
            l1[jV1] = new JCheckBox(items[jV1], values[jV1]);
        
        for (int i = 0; i < items.length; i++) {
            new WorkingCheckBox().scaleCheckBoxIcon(l1[i], 21);
            panelOutload.add(l1[i]).setFont(bigFontTR);
            
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
        
        JButton buttonAdd = new JButton("Выгрузить в excel");
        buttonAdd.setFont(bigFontTR);
        buttonAdd.setPreferredSize(new Dimension(200, 30));

        JPanel panelBt = new JPanel();
        panelBt.setPreferredSize(new Dimension(100, 30));
        panelBt.add(buttonAdd);

        JPanel jPanelSelecting = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();

        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 0;
        
        c.weightx = 1;
        c.weighty = 1;
        c.insets = new Insets(0, 0, 0, 0);
        c.fill = GridBagConstraints.CENTER;
        jPanelSelecting.add(panelOutload, c);
        
        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 1;
        
        c.weightx = 1;
        c.weighty = 1;
        c.insets = new Insets(0, 0, 0, 0);
        c.fill = GridBagConstraints.BOTH;     
        jPanelSelecting.add(panelBt, c);
        jPanelSelecting.setPreferredSize(new Dimension(510, 730));
        
        JFrame frameSC = new JFrame();      
        frameSC.add(jPanelSelecting);
        // frameSC.setIconImage(image);
        frameSC.setLocationRelativeTo(null);
        frameSC.setSize( new Dimension(610, 790) );
        frameSC.isMaximumSizeSet();
        frameSC.setLocation(50, 50);
        frameSC.setVisible(true);
       
        buttonAdd.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {    
            	
            	frameSC.setVisible(false);
            	ArrayList<Integer> nColumns = new ArrayList<Integer>();
            	int nL = 0;
                // записываем выделенные значения
                for (int i = 0; i < items.length; i++) {
                    if (l1[i].isSelected() == true) {
                        nColumns.add(i);
                    }
                }             
                try {
                    InputStream inputStream = new FileInputStream("Reestr.xls");
                    HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
                    Sheet currentSheet = workbook.getSheetAt(2);
                    Iterator<Row> rowIterator = currentSheet.iterator();                   
                    rowIterator.next(); // skip the header row            
                    
                    ArrayList dataPart = new ArrayList();   
                    
                    int counter = 0;
                    int currentRow = 0; 
                    int countRow = 1;
                    int nClm = 0;
                    
                    // если выбраны не все столбцы
                    if (nColumns.get(0) != 0) {      
                    	countRow = 0;
                    	// определяем нужные строки
                    	for (int i = 0; i < data.size(); i++) {                   		
                    		double d = (double)nColumns.get(nClm) / (double)(i - ( columnCount * countRow) + 1);
                    		
                    		if (d  == 1.0) {                   				                    			
                    			dataPart.add(data.get(i));
                    			
                   			 	if (nColumns.size() != nClm + 1) {
                   			 		nClm++;
                   			 	} else {
                   			 		nClm = 0;
                   			 	}
                    		}                   		
                    		if (counter != (columnCount - 1) ) {
                    			counter++;                   			
                    		} else {
                    			countRow ++;
                    			counter = 0;	                    			
                    		}
                    	}                   	
                    	currentRow = currentRow + 2;
                    	counter = 0;
                    	countRow = 0;
                    	nClm = 0;                    	
                    } else {
                    	// определяем нужные строчки                    	                    		
                    	ArrayList newData = new ArrayList();           	
                    	
                    	for (int i = 0; i < data.size(); i++) {	
                			newData.add(data.get(i));		                				                			
                			counter++;
                			
                    		if (counter == (columnCount - 1) ) {	
                    			countRow ++;	
                    			counter = 0;
                    		}	                    		
                    	}
                    	counter = 0;
                    	countRow = 1;
                    	currentRow = currentRow + 2;
                    	data.clear();
                    	
                    	for (int i = 0; i < newData.size(); i++) {
                    		data.add(newData.get(i));
                    	}                         	
                    	newData.clear();                    	 	
                    	currentRow = 0;                    	
                    }       
                    
                    int nClms;
                    int nRow;
                    
                    if (nColumns.get(0) == 0) {
                    	 nClms = columnCount;
                    	 nRow = data.size()/nClms;
                    } else {
                    	nClms = nColumns.size();
                        nRow = dataPart.size()/nClms;
                    }
                    InputStream inputStream2 = new FileInputStream("Reestr.xls");
                    HSSFWorkbook sh = new HSSFWorkbook(inputStream2);
                    HSSFSheet worksheet = sh.getSheetAt(2);
                    HSSFCell cell = null;
                    int lastRow = 2;
                    Row row ;
                    Row row02;
                    row02 = worksheet.getRow(0);
                    row02.createCell((columnCount - 1) ).setCellValue("");
                    
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
                    
                    // удаляем лист с девайсами и автомобилями
                    sh.removeSheetAt(0);
                    sh.removeSheetAt(0);
                    
                    // удалили все строки кроме заголовка
                    for (int i = 2; i <  100000; i++) {
                    	if (worksheet.getRow(i) != null) {
                    		worksheet.removeRow(worksheet.getRow(i));
                    	}	
                    }
                    
                    int dataIndex = 0;
                    
                    for (int j = 0; j < nRow; j++) {
                    	nL = 0;
                        row = worksheet.createRow(lastRow);  
                        
                        if (nColumns.get(0) != 0) {
	                        for (int i = 0; i <= (columnCount - 1); i++) {
	                        	if ((nColumns.get(nL)-1) == i) {  	                      		
	                        		cell = (HSSFCell) row.createCell(nL);                        		 
	                        		if (dataPart.get(dataIndex).getClass().getTypeName() == "java.lang.String") {  	 	                        		
 	 	                        		row.createCell(nL).setCellValue((String) dataPart.get(dataIndex)); 	 	                        		
 	 	                        	} else {
 	 	                        		row.createCell(nL).setCellValue((int) dataPart.get(dataIndex));
 	 	                        	}	                        		
	                        		row.getCell(nL).setCellStyle(style);
		                            dataIndex++;
	                                nL ++;	                                
	                                if (nColumns.size() == nL) {
	                                	nL = 0;
	                                }	                                
	                        	}  
	                        }	                        
                        } else {                         		
		            		for (int i = 0; i <= (columnCount - 1); i++) {
		                        cell = (HSSFCell) row.createCell(i);	
		                        System.out.println ( data.size() ); 
		                    	if ( data.get(dataIndex).getClass().getTypeName() == "java.lang.String" ) {  	 	                        		
		                    		row.createCell(i).setCellValue( (String) data.get(dataIndex) ); 	 	                        		
		                    	} else {
		                    		row.createCell(i).setCellValue( (int) data.get(dataIndex) );
		                    	}
		                    	row.getCell(i).setCellStyle(style);
		                        dataIndex++;
		            		 } 
                        }	                        	                
                        lastRow++;
                    }                  
                    dataPart.clear();
                    data.clear();
                    
                    // удаляем не нужные столбцы
                    ArrayList<Integer> nColumnsDel = new ArrayList<Integer>();
                    int nC = 0;
                    nColumns.add(10000);                   
                    if (nColumns.get(0) != 0) {
                    	for (int i = 1; i <=columnCount; i++) {
	                    	if (nColumns.get(nC).equals(i)) {
	                    		 nC++;
	                    	} else {
	                    		nColumnsDel.add(i);
	                    	}
	                    }
                    	for (int i = 0; i < nColumnsDel.size(); i++) {
		                    for (int rId = 0; rId <= 1; rId++) {
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
                    nColumnsDel.clear();
                    
                    worksheet.removeRow(worksheet.getRow(1));          
                    Row newRow0 = worksheet.createRow(1);
              
                    int nCD = 0;
                    // даем название столбцам
                    if (nColumns.get(0) != 0) {
                		for (int i = 1; i <= columnCount; i++) {
                			
	                    	if (nColumns.get(nCD).equals(i)) {
                    			newRow0.createCell(nCD).setCellValue(nameColumn[i]);
                    			newRow0.getCell(nCD).setCellStyle(styleHeader);
                    			nCD++;	                    		
	                    	}
	                    }			                                        	                    
	                 } else {
	                	 for (int i = 1; i <=columnCount; i++) {
	                    	newRow0.createCell(i-1).setCellValue(nameColumn[i]);
	                    	newRow0.getCell(i-1).setCellStyle(styleHeader);	 
	                	}	
	                }
                    
                    int nC2 = 0;
                    for (int nColumn = 0; nColumn <= items.length; nColumn++) {
                    	if ((nColumns.get(nC2)-1) == nColumn) {
	                    	int width = (int) worksheet.getColumnWidthInPixels(nColumn);
	                    	if ( nColumns.get(nC2)-1 == 0 ) {
	                    		width = width * 30;
	                    	} else {
	                    		width = width * 45;
	                    	}
	                    	worksheet.setColumnWidth(nC2, width);	                    	
	                    	nC2++;
                    	} 
                    }

                    nColumns.clear();
                    
                    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
                    String filenameOut = "База МТЦ (Орг. техника) от " + timeStamp  + ".xls";
                    
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
                		FileOutputStream output_file = new FileOutputStream(file);
                	
                		sh.write(output_file);
                		output_file.close();
                		workbook.close();
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
            
			private int ParseInt(Object object) {
				// TODO Auto-generated method stub
				return 0;
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

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub			
	}  
}
