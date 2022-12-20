package net.codejava;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.image.BufferedImage;
import javax.swing.JList;
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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Iterator;
import javax.swing.DefaultComboBoxModel;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.ListCellRenderer;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SelectingColumnsDC  implements ActionListener {

    public class CustomerItem {

        public String label;
        public boolean status;

        public CustomerItem(String label, boolean status) {
            this.label = label;
            this.status = status;
        }
    }
    
    public class RenderCheckComboBox implements ListCellRenderer {

        //a JCheckBox is associated for one item
        JCheckBox checkBox;

        Color selectedBG = new Color(112, 146, 190);

        public RenderCheckComboBox() {
            this.checkBox = new JCheckBox();
        }

        @Override
        public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected,
                boolean cellHasFocus) {

            //recuperate the item value
            CustomerItem value_ = (CustomerItem) value;

            if (value_ != null) {
                checkBox.setText(value_.label);
                checkBox.setSelected(value_.status);
            }

            if (isSelected) {
                checkBox.setBackground(Color.GRAY);
            } else {
                checkBox.setBackground(Color.WHITE);
            }
            return checkBox;
        }

    }
	 // несколько JCheckBox 
    public void actionPerformed(ActionEvent e) {
        JComboBox cb = (JComboBox) e.getSource();
        CheckComboStore store = (CheckComboStore)cb.getSelectedItem();
        CheckComboRenderer ccr = (CheckComboRenderer)cb.getRenderer();
        ccr.checkBox.setSelected((store.state = !store.state));
    }
    
    // изменяем размер галочки в JCheckBox
    public void scaleCheckBoxIcon(JCheckBox checkbox, int heightWidth) throws IOException {
        boolean previousState = checkbox.isSelected();
        checkbox.setSelected(false);

        Icon boxIcon = UIManager.getIcon("CheckBox.icon");
        BufferedImage boxImage = new BufferedImage(
        boxIcon.getIconWidth(), boxIcon.getIconHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics graphics = boxImage.createGraphics();

        try{
            boxIcon.paintIcon(checkbox, graphics, 0, 0);
        }
        finally
        {
            graphics.dispose();
        }

        ImageIcon newBoxImage = new ImageIcon(boxImage);
        Image finalBoxImage = newBoxImage.getImage().getScaledInstance(
        boxImage.getWidth(), boxImage.getHeight(), Image.SCALE_SMOOTH);
        finalBoxImage = finalBoxImage.getScaledInstance(heightWidth, heightWidth, Image.SCALE_SMOOTH);

        checkbox.setIcon(new ImageIcon(finalBoxImage));
        checkbox.setSelected(true);

        Icon checkedBoxIcon = UIManager.getIcon("CheckBox.icon");
        BufferedImage checkedBoxImage = new BufferedImage(
                boxIcon.getIconWidth(), boxIcon.getIconHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics checkedGraphics = checkedBoxImage.createGraphics();

        try{
            checkedBoxIcon.paintIcon(checkbox, checkedGraphics, 0, 0);
        }
        finally{
            checkedGraphics.dispose();
        }

        ImageIcon newCheckedBoxImage = new ImageIcon(checkedBoxImage);
        Image finalCheckedBoxImage = newCheckedBoxImage.getImage().getScaledInstance(
                boxImage.getWidth(), boxImage.getHeight(), Image.SCALE_SMOOTH);
        finalCheckedBoxImage = finalCheckedBoxImage.getScaledInstance(heightWidth, heightWidth, Image.SCALE_SMOOTH);

        checkbox.setSelectedIcon(new ImageIcon(finalCheckedBoxImage));
        checkbox.setSelected(false);
        checkbox.setSelected(previousState);
    }
    
	Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
	static String fileDictName = "";
	
    public void selecting(ArrayList inputData, ArrayList inputHeaders) throws IOException {  
    	
    	// копируем данные, чтобы можно было использовать кнопку множество раз
    	ArrayList data = new ArrayList(inputData); 	 
    	
    	for (int i = 18; i < data.size(); i += 18) {
    		data.remove(i);
    	}
    	/*
    	ImageIcon liderIcon = new ImageIcon(new DeviceAndConsumables().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();
        */
        JPanel jPanelL1 = new JPanel(new GridLayout(22, 3, 5, 0));
    	jPanelL1.add(new JLabel("Выберите нужные назначения: ", SwingConstants.LEFT)).setFont(bigFontTR);
    	   	
        String[] items = {"Все столбцы",
        		"№", "Вид контроля",  "Назначение (область применения)", "Наименование прибора", "Тип, марка, модель", 
        		"<html><center>Производитель, страна производства, марка,<br> модель, основные технические характеристики</html>",
        		"Зав.№", "Количество", "Год выпуска", "Дата поверки (калибровки)","Дата окончания поверки (калибровки)", "Документы", 
        		"Техническое состояние", "Указание в поверке на принадлежность к организации", 
       		     "Форма собственности", "Владелец оборудавания", "Местонахождение", "Примечание"}; // 19
        
        String[] nameColumn = {"Все столбцы",
        		"№", "Вид контроля",  "Назначение (область применения)", "Наименование прибора", "Тип, марка, модель", 
        		"Производитель, страна производства, марка," + " модель, основные технические характеристики",       		
        		"Зав.№", "Количество","Год выпуска", "Дата поверки (калибровки)","Дата окончания поверки (калибровки)", "Документы", 
        		"Техническое состояние", "Указание в поверке на принадлежность к организации", 
       		     "Форма собственности", "Владелец оборудавания", "Местонахождение", "Примечание"}; 
        
        JComboBox<CustomerItem> combo = new JComboBox<CustomerItem>() {
            @Override
            public void setPopupVisible(boolean visible) {
                if (visible) {
                    super.setPopupVisible(visible);
                }
            }
        };
        
        CustomerItem[] headersString = new CustomerItem[inputHeaders.size()];
        headersString[0] = new CustomerItem(inputHeaders.get(0).toString(), true);
        
        for (int i = 1; i < inputHeaders.size(); i++) {
        	headersString[i] = new CustomerItem(inputHeaders.get(i).toString(), false);
        }	
        
        combo.setModel(new DefaultComboBoxModel<CustomerItem>(headersString));
        combo.setRenderer(new RenderCheckComboBox());
        
        combo.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent ae) {
                CustomerItem item = (CustomerItem) ((JComboBox) ae.getSource()).getSelectedItem();
                item.status = !item.status;
                // update the ui of combo
                combo.updateUI();
                //keep the popMenu of the combo as visible
                combo.setPopupVisible(true);
                
                for (int i = 1; i < headersString.length; i++) {
                	
                	if (headersString[i].status == true) {
                		headersString[0].status = false;
                	}
                }	
            }
        });
        
        jPanelL1.add(combo).setFont(bigFontTR);
        
        jPanelL1.add(new JLabel("Выберите нужные столбцы: ", SwingConstants.LEFT)).setFont(bigFontTR);
        
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
        jPanelL1.setPreferredSize(new Dimension(510, 730));
        panel1.add(jPanelL1, BorderLayout.CENTER);
        
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
        frameSC.setLocation(50, 50);
        frameSC.setVisible(true);
       
        buttonAdd.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {    
            	ArrayList headers = new ArrayList();
            	ArrayList headersName = new ArrayList();
            	for (int i = 0; i < headersString.length; i++) {
            		if (headersString[i].status == true) {
            			headers.add(i);
            			headersName.add(headersString[i].label.toString());
            			// System.out.println(headersString[i].label + " " + i);
            		}	
            	}
            	
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
                try {
                    InputStream inputStream = new FileInputStream("Reestr.xls");
                    HSSFWorkbook workbook = new HSSFWorkbook(inputStream);  
                    Sheet currentSheet = workbook.getSheetAt(0);
                    Iterator<Row> rowIterator = currentSheet.iterator();
                    
                    rowIterator.next(); // skip the header row

                    Row nextRow2 = rowIterator.next();
                    Iterator<Cell> cellIterator2 = nextRow2.cellIterator();                   
                    
                    ArrayList dataPart = new ArrayList();
                    ArrayList currentRows = new ArrayList<Integer>();
                    // int currentRows[] = new int[2];
                    int currentZ[] = new int[2];
                    int currentRowsEnd[] = new int[2];                   
                    
                    // получаем выделенные заголовки
                    if(headers.get(0).equals(0)) {
                    	// если выбраны все заголовки
                    	for (int i = 0; i <= currentSheet.getLastRowNum(); i++) {
                    		currentRows.add(i);
                    	}
                    } else {
                    	
                    	for (int j = 0; j < headers.size(); j++) {
                    		for (int i = 0; i < data.size() - 17; i+=18) {
                    			if (data.get(i + 2).equals(headersName.get(j))) {    
                    				if (data.size() < 18) {                   					
                    					currentRows.add(0);
                    				} else {
                    					currentRows.add(i/18);
                    				}
                    			}
                    		}
                    	} 	
                    }
                    
                    Collections.sort(currentRows);
                    
                    int counter = 0;
                    int currentRow = 0; 
                    int countRow = 1;
                    int nClm = 0;
                    // если выбраны не все столбцы
                    if (nColumns.get(0) != 0) {      
                    	countRow = 0;
                    	// определяем нужные строки
                    	if (!headers.get(0).equals(0)) {
	                    	while (currentRow < currentRows.size()) {
		                    	for (int i = 0; i < data.size(); i++) {
		                    		
		                    		double d = (double)nColumns.get(nClm) / (double)(i - (18 * countRow) + 1);
		                    		
		                    		if (d  == 1.0) {
		                    				                    			
		                    			if (countRow == (int)currentRows.get(currentRow)) {
		                    				dataPart.add(data.get(i));
		                    			}                  			
		                    			
		                   			 	if (nColumns.size() != nClm + 1) {
		                   			 	nClm++;
		                   			 	} else {
		                   			 	nClm = 0;
		                   			 	}
		                    		}  
		                    		
		                    		if (counter != 17) {
		                    			counter++;                   			
		                    		} else {
		                    			countRow ++;
		                    			counter = 0;	                    			
		                    		}
		                    	}
		                    	currentRow ++;
		                    	counter = 0;
		                    	countRow = 0;
		                    	nClm = 0;
	                    	}
                    	} else {
	                    	for (int i = 0; i < data.size(); i++) {
	                    		
	                    		double d = (double)nColumns.get(nClm) / (double)(i - (18 * countRow) + 1);
	                    		
	                    		if (d  == 1.0) {
	                    				                    			
	                    			dataPart.add(data.get(i));   			
	                    			
	                   			 	if (nColumns.size() != nClm + 1) {
	                   			 	nClm++;
	                   			 	} else {
	                   			 	nClm = 0;
	                   			 	}
	                    		}  
	                    		
	                    		if (counter != 17) {
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
	                    	
                    	}
                    	
                    } else {
                    	// определяем нужные строчки
                    	
                    	if (!headers.get(0).equals(0)) {
                    		
	                    	ArrayList newData = new ArrayList();           	
	                    	
	                    	while (currentRow < currentRows.size()) {
		                    	for (int i = 0; i < data.size(); i++) {
		                			if (countRow - 1 == (int)currentRows.get(currentRow)) {
		                				newData.add(data.get(i));
		                			}
		                			
		                			counter++;
		                			
		                    		if (counter == 18) {	
		                    			countRow ++;	
		                    			counter = 0;
		                    		}	                    		
		                    	}
		                    	counter = 0;
		                    	countRow = 1;
		                    	currentRow ++;
	                    	}	
	                    	
	                    	data.clear();
	                    	for (int i = 0; i < newData.size(); i++) {
	                    		data.add(newData.get(i));
	                    	}                                 	
	                    	newData.clear();
                    	} 	
                    	currentRow = 0;
                    	
                    }       
                    
                    currentRows.clear();
                    
                    int nClms;
                    int nRow;
                    
                    if (nColumns.get(0) == 0) {
                    	 nClms = 18;
                    	 nRow = data.size()/nClms;
                    } else {
                    	nClms = nColumns.size();
                        nRow = dataPart.size()/nClms;
                    }
                    InputStream inputStream2 = new FileInputStream("Reestr.xls");
                    HSSFWorkbook sh = new HSSFWorkbook(inputStream2);
                    HSSFSheet worksheet = sh.getSheetAt(0);
                    HSSFCell cell = null;
                    
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
                    
                    int lastRow = 2;
                    Row row ;
                    Row row02;
                    row02 = worksheet.getRow(0);
                    row02.createCell(14).setCellValue("");
                    
                    // удаляем лист с автомобилями
                    sh.removeSheetAt(1);
                    sh.removeSheetAt(1);
                    
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

	                        for (int i = 0; i <= 17; i++) {
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
                        	 int mrg = 0;
                        	 for (int i = 0; i <= 17; i++) {
                        		 
	                            if (i >= 2 && data.get(dataIndex).toString().length() == 0) {
	                            	mrg ++;
	                            }
	                            
	                            dataIndex++;
                        	 }
                        	 dataIndex = dataIndex - 18;
                        	 
                        	 if (mrg == 14) {
                        		 
	                        	 for (int i = 0; i <= 17; i++) {
	 	                        	cell = (HSSFCell) row.createCell(i);
	 	                            row.createCell(i).setCellValue((String) data.get(dataIndex));
	 	                            row.getCell(i).setCellStyle(style);
	 	                            dataIndex++;
	                         	 } 	                        	 
                        	 } else {          
                        		 
                        		 for (int i = 0; i <= 17; i++) {
 	 	                        	cell = (HSSFCell) row.createCell(i);
 	 	                        	if (data.get(dataIndex).getClass().getTypeName() == "java.lang.String") {  	 	                        		
 	 	                        		row.createCell(i).setCellValue((String) data.get(dataIndex)); 	 	                        		
 	 	                        	} else {
 	 	                        		row.createCell(i).setCellValue((int) data.get(dataIndex));
 	 	                        	}
 	 	                        	row.getCell(i).setCellStyle(style);
 	 	                            dataIndex++;
                        		 }   
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
                    	for (int i = 1; i <= 17; i++) {
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
                		for (int i = 1; i <= 18; i++) {                   			
	                    	if (nColumns.get(nCD).equals(i)) {
                    			newRow0.createCell(nCD).setCellValue(nameColumn[i]);
                    			newRow0.getCell(nCD).setCellStyle(styleHeader);
                    			nCD++;	                    		
	                    	}
	                    }			                                        	                    
	                 } else {
	                	 for (int i = 1; i <= 18; i++) {
	                    	newRow0.createCell(i-1).setCellValue(nameColumn[i]);
	                    	newRow0.getCell(i-1).setCellStyle(styleHeader); 
	                	}	
	                }
                    
                    if (nColumns.get(0) != 0) {
	                    int nC2 = 0;
	                    for (int nColumn = 0; nColumn <= 17; nColumn++) {
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
                    
	                    /*int width = (int) worksheet.getColumnWidthInPixels(19);
	                    width = width * 30;
	                    for (int nColumn = nColumns.size() + 1; nColumn <= 17; nColumn++) {
	                    	worksheet.setColumnWidth(nColumn-1, width);
	                    }*/
                    }    
                    
                    nColumns.clear();
                    
                    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
                    String filenameOut = "База МТЦ (Приборы и расходники) от " + timeStamp  + ".xls";
                    
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
}  

class CheckComboRenderer implements ListCellRenderer
{
    JCheckBox checkBox;

    public CheckComboRenderer() throws IOException {    	
        checkBox = new JCheckBox();
        new SelectingColumnsDC().scaleCheckBoxIcon(checkBox, 22);
    }
    
    public Component getListCellRendererComponent(JList list, Object value,
                                                  int index, boolean isSelected,
                                                  boolean cellHasFocus)
    {
        CheckComboStore store = (CheckComboStore)value;
        checkBox.setText(store.id);
        checkBox.setSelected(((Boolean)store.state).booleanValue());
        checkBox.setBackground(isSelected ? Color.gray : Color.white);
        checkBox.setForeground(isSelected ? Color.white : Color.black);

        return checkBox;
    }
}

class CheckComboStore
{
    String id;
    Boolean state;

    public CheckComboStore(String id, Boolean state)
    {
        this.id = id;
        this.state = state;
    }
}




