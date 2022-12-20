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

public class Technics extends JFrame {
	 // приборы и расходники 	
	final static Object[] mainHeaders = {"№", "Вид оборудования", "Наименование прибора", "Тип, марка, модель", 
			"<html> <center>Производитель, страна<br>производства,основные<br>технические характеристики<html>",
		    "Зав.№", "<html> <center>Год<br>выпуска", "Комплектность", "Документы", "<html> <center>Техническое<br>состояние", "Форма собственности",
		    "<html> <center>Владелец<br>оборудавания<html>",
		    "<html> <center>Местонахождения и<br>текущие обязательства", "Офис", "<html> <center>Инвентарный<br>номер", "Примечание",
			// };
		    " "}; // кнопка редактировать	
	
	int nColumn;
	
    String timeStamp = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss").format(Calendar.getInstance().getTime());
    
    int startSize = 0;
    JTextField nameSearch = new JTextField(15);
    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
    
    static JFrame frame = new JFrame();
    static ArrayList <String>copyData = new ArrayList<String>();
    static ArrayList jointUpload = new ArrayList();
    
    ArrayList namesSearchAll = new ArrayList();
    ArrayList typeEquipment = new ArrayList();
    ArrayList technicalCondition = new ArrayList();
    ArrayList locations = new ArrayList();
    ArrayList allLocations = new ArrayList();
    ArrayList offices = new ArrayList();
    
    static int cH = 0;
    static int cW = 0;
    
    public void start(JFrame frame, int hThree, JTable tableStart) throws IOException, ParseException {    	
    	System.out.println( "mainHeaders: " + mainHeaders.length );
    	// установка другой иконки для JFrame
    	/*
    	ImageIcon liderIcon = new ImageIcon(new Technics().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();
        frame.setIconImage(image);
        */
    	frame.getContentPane().setLayout(new BorderLayout());
    	Font myFont = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 15);
    	JButton buttonSearch = new JButton("Поиск");
        JButton buttonStart = new JButton("<html><center>" + "Сбросить параметры поиска" + "<center><html>"); 
      
        JPanel panelMain = new JPanel();
   	 	panelMain.add(nameSearch).setFont(myFont);
   	    
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
            HSSFSheet currentSheet = workbook.getSheetAt(2);
            
		    for (Row row: currentSheet) {
		    	if (row.getRowNum() >= 2) {		        	
		    		for (Cell cell: row) {		        		
			        	switch ( cell.getCellTypeEnum() ) {
			        		case STRING:
			        			data.add( cell.getStringCellValue() );
			        			if (cell.getStringCellValue().length() != 0) {
				        			if ( cell.getColumnIndex() == 1 ) {
				        				typeEquipment.add( cell.getStringCellValue() );
				        			}
				        			if ( cell.getColumnIndex() >= 2 && cell.getColumnIndex() <= 4 ) {
				        				namesSearchAll.add( cell.getStringCellValue() );
				        			}
				        			if ( cell.getColumnIndex() == 9 ) {
				        				technicalCondition.add( cell.getStringCellValue() );
				        			}
				        			if (cell.getColumnIndex() == 12) {
				        				locations.add( cell.getStringCellValue() );
				        			}
				        			if (cell.getColumnIndex() == 13) {
				        				offices.add( cell.getStringCellValue() );
				        			}
			        			}
			        			break;
			        		case NUMERIC:
			        			data.add( (int) cell.getNumericCellValue() );
			        			if ( cell.getColumnIndex() >= 2 && cell.getColumnIndex() <= 4 ) {
			        				namesSearchAll.add( cell.getNumericCellValue() );
			        			}
			        			if (cell.getColumnIndex() == 13) {
			        				offices.add( (int) cell.getNumericCellValue() );
			        			}
			        			break;
			        		default:
			        			if ( cell.getColumnIndex() <= (mainHeaders.length - 2) ) {
			        				data.add("");
			        			}
			        			break;
			        	}	
			        	if ( cell.getColumnIndex() == (mainHeaders.length - 2) ) {
			        		data.add( cell.getRowIndex() );
			        		break;
			        	}
		        	}
		        }   
		    }    
		    inputStream.close();
		    
            int cl = mainHeaders.length;
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
						new AddDataTechnics().inputValues();
					} catch (IOException e1) {
						e1.printStackTrace();
					} 
                }
            });
							           	
        	frame.getContentPane().removeAll();            
            data.clear();
            data = new ArrayList();    
            
            JTable table1 = new JTable(dm) {                
            	// запрет на редактирование ячеек в таблице
                private static final long serialVersionUID = 1L;                
                // кнопку редактирования изменять можно
                public boolean isCellEditable( int row, int column ) {                
                	if ( column != mainHeaders.length - 1) {
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
 	                return c;
 		        }
            };
                       
            table1.changeSelection(0, 0, false, false);
            JScrollPane scrollPane1 = new JScrollPane(table1);
            getContentPane().add(scrollPane1);
            
            // кнопка "редактировать"    
            table1.getColumn(" ").setCellRenderer( new ButtonRendererTechnics(frame) );
            table1.getColumn(" ").setCellEditor( new ButtonEditorTechnics(new JCheckBox(), frame) );
            table1.getColumnModel().getColumn( mainHeaders.length - 1 ).setPreferredWidth(30); // кнопка редактировать           
            // кнопка "редактировать"   
            
            JTableHeader th1 = table1.getTableHeader();
        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));              	
        	
        	// горизонтальная прокрутка заголовков
        	table1.getTableHeader().setPreferredSize(new Dimension(10000, 120));
        	
        	table1.setRowHeight(25);
        	
        	for (int i = 0; i <= mainHeaders.length - 2; i++) {
		        switch (i) {
		        	case 0:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(35);
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 4:
		        	case 7:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(190); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 5:
		        	case 8:		        	
		        	case 15:
		        	case 9:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(110); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 6:
		        	case 13:
		        	case 14:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(70); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        	case 1:
		        	case 2:
		        	case 3:
		        	case 10:	
		        	case 11:
		        	case 12:
		        		table1.getColumnModel().getColumn(i).setPreferredWidth(150); 
		        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
		        		break;
		        }
        	}
            
            JPanel panel12 = new JPanel(new BorderLayout(10, 10));
            JPanel panel01 = new JPanel(new BorderLayout(0, 25));
            
            table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);                 
            table1.setPreferredScrollableViewportSize(tableStart.getPreferredSize());            
            panel01.setPreferredSize(new Dimension(0, 267)); // was 300
            
            three.setPreferredSize(new Dimension(0, hThree));
            three.setMaximumSize(new Dimension(0, hThree));
            three.setMinimumSize(new Dimension(0, 10));
            
            JPanel panelL = new JPanel( new GridLayout(5, 0, 10, 10) );          
            Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);                                    	    
            
            Set setKC = new HashSet( typeEquipment );
            typeEquipment.clear();
            typeEquipment = new ArrayList(setKC);
            Set setTO = new HashSet( technicalCondition );
       	    technicalCondition.clear();
       	    technicalCondition = new ArrayList(setTO);
	       	Set setEO = new HashSet( locations );
	        locations.clear();
	        locations = new ArrayList(setEO);
	        Set setO = new HashSet( offices );
            offices.clear();
            offices = new ArrayList(setO);
            
    	    panelL.add(new JLabel("Вид оборудования:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceTypeEquipment().outputPanel(typeEquipment, "Все виды оборудования")).setFont(bigFontTR);           
    	    panelL.add(new JLabel("", SwingConstants.RIGHT)).setFont(bigFontTR);
                  
            panelL.add(new JLabel("Техническое состояние:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add( new ChoiceTechnicalCondition().outputPanel( technicalCondition, "Все технические состояния") ).setFont(bigFontTR); 
    	    panelL.add(new JLabel("", SwingConstants.RIGHT)).setFont(bigFontTR);                  
            
            panelL.add(new JLabel("Нынешнее состояние:", SwingConstants.RIGHT)).setFont(bigFontTR);
            panelL.add(new ChoiceLocation().outputPanel(locations, "Все состояния")).setFont(bigFontTR);            
    	    panelL.add(new JLabel("" , SwingConstants.RIGHT)).setFont(bigFontTR);	    
            
            panelL.add(new JLabel( "Офис:", SwingConstants.RIGHT )).setFont(bigFontTR);
            panelL.add(new ChoiceOffice().outputPanel( offices, "Все офисы" )).setFont(bigFontTR);
            panelL.add(new JLabel("" , SwingConstants.RIGHT)).setFont(bigFontTR);
           
            panelL.add(new JLabel( "Наименование:", SwingConstants.RIGHT) ).setFont(bigFontTR);
            nameSearch.setMaximumSize( new Dimension(10, 14) );
            nameSearch.setMinimumSize( new Dimension(10, 14) );
            panelL.add( nameSearch );
            
            JPanel panelFF22 = new JPanel(new GridBagLayout());
            panelFF22.setAlignmentX(CENTER_ALIGNMENT);
            
            GridBagConstraints c22 = new GridBagConstraints();
            c22.fill = GridBagConstraints.VERTICAL;
            c22.gridx = 1;
            c22.gridy = 0;            
            c22.weightx = 0.95;
            c22.weighty = 2;
            c22.fill = GridBagConstraints.BOTH;
            panelFF22.add(new JScrollPane( table1 ), c22);
            
            c22.fill = GridBagConstraints.VERTICAL;
            c22.gridx = 1;
            c22.gridy = 1;           
            c22.weightx = 1;
            c22.weighty = 0.1;
            c22.fill = GridBagConstraints.BOTH;         
            panelFF22.add(three, c22);

            JButton buttonDC = new JButton("<html><center>" + "Приборы и<br> расходники"  + "<center><html>");
            buttonDC.setBackground(Color.LIGHT_GRAY);           
            
            JButton buttonCars = new JButton( "Автомобили" );
            buttonCars.setBackground(Color.LIGHT_GRAY);
            
            JButton buttonTechnology = new JButton( "Орг. техника" );
            buttonTechnology.setForeground( Color.WHITE );
            buttonTechnology.setBackground( Color.LIGHT_GRAY );
                                   
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
            
            cH = panelL.getHeight() - 200;
            cW = 100;
            
            buttonCars.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                	nameSearch.setText("");
                    nameSearch = new JTextField(14);
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
            
            buttonDC.addActionListener(new ActionListener() {
            	@Override
                public void actionPerformed(ActionEvent e) {
                	nameSearch.setText("");
                    nameSearch = new JTextField(14);
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
            	nameSearch.setText("");
                nameSearch = new JTextField(15);  
			    jointUpload.clear();
			    copyData.clear();
                frame.setPreferredSize(frame.getSize());
                
                namesSearchAll.clear();
                typeEquipment.clear();
                technicalCondition.clear();
                locations.clear();
                allLocations.clear();
                offices.clear();
                
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

        Set setYears = new HashSet(namesSearchAll);
        namesSearchAll.clear();
        namesSearchAll = new ArrayList(setYears);
        
		new AutoSuggestor(nameSearch, frame, null, Color.WHITE.brighter(), Color.DARK_GRAY, Color.RED, 0.8f, cH, cW) {		 	
			protected boolean wordTyped(String typedWord)  {	     	             		
				setDictionary(namesSearchAll);
		 		return super.wordTyped(typedWord);
		    }                                       
		};		
		
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		buttonSearch.addActionListener(new ActionListener() {			
			@Override
			public void actionPerformed(ActionEvent e) {
				
			    ArrayList data = new ArrayList();			   		    
			    ArrayList selectedTypesEquipment = new ChoiceTypeEquipment().selectedTypes();
			    ArrayList selectedTypesCondition = new ChoiceTechnicalCondition().selectedTypes();
			    ArrayList selectedLocations = new ChoiceLocation().selectedLocations();
			    ArrayList selectedOffices = new ChoiceOffice().selectedLocations();
			    
				try {			
				    InputStream inputStream = new FileInputStream("Reestr.xls");
				    Workbook workbook = new HSSFWorkbook(inputStream);
				    Sheet currentSheet = workbook.getSheetAt(2);
				    Iterator<Row> rowIterator = currentSheet.iterator();
				    rowIterator.next(); 
				    rowIterator.next();
				    
				    boolean column1 = false;
				    boolean column7 = false;
				    boolean column9 = false;
				    boolean column10 = false;
				    boolean column234 = false;
				    
				    for (Row row: currentSheet) {	
				    	
				    	column1 = false;
				    	column7 = false;
				    	column9 = false;
				    	column10 = false;
				    	column234 = false;
				    	
				        for (Cell cell: row) {				            
					    	if (row.getRowNum() > 1) {
						    	int columnIndex = cell.getColumnIndex();						    	
						        switch (cell.getCellType()) {
						        	case Cell.CELL_TYPE_STRING:
						        		if ( cell.getColumnIndex() == 1 ) {
						        			if ( selectedTypesEquipment.get(0).equals("Все виды оборудования") ) {
						        				column1 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedTypesEquipment.size(); i++) {
						                        	if (cell.getStringCellValue().equals( selectedTypesEquipment.get(i) )) {				                        		
						                        		column1 = true;
						                        	} 
					                        	}
				                        	}
						        		}
			                        	if ( cell.getColumnIndex() >= 2 && cell.getColumnIndex() <= 4 ) {
			                        		
			                        		String str = cell.getStringCellValue();				                        
			                        		int indexM = str.indexOf( nameSearch.getText() );
			                        		
				                        	if (indexM != -1) {
				                        		column234 = true;
				                        	}
				                        	if (str.replaceAll("\\s","").indexOf( nameSearch.getText()) != -1 ) {
				                        		column234 = true;
				                        	}
			                        	}
						        		if ( cell.getColumnIndex() == 9 ) {
						        			if ( selectedTypesCondition.get(0).equals("Все технические состояния") ) {
						        				column7 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedTypesCondition.size(); i++) {
						                        	if (cell.getStringCellValue().equals( selectedTypesCondition.get(i) )) {				                        		
						                        		column7 = true;
						                        	} 
					                        	}
				                        	}
						        		}
						        		if ( cell.getColumnIndex() == 12 ) {
						        			if ( selectedLocations.get(0).equals("Все состояния") ) {
						        				column9 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedLocations.size(); i++) {
						                        	if (cell.getStringCellValue().equals( selectedLocations.get(i) )) {				                        		
						                        		column9 = true;
						                        	} 
					                        	}
				                        	}
						        		}
						        		if ( cell.getColumnIndex() == 13 ) {
						        			if ( selectedOffices.get(0).equals("Все офисы") ) {
						        				column10 = true;
				                        	} else {				                        		
					                        	for (int i = 0; i < selectedOffices.size(); i++) {
						                        	if (cell.getStringCellValue().equals( selectedOffices.get(i) )) {				                        		
						                        		column10 = true;
						                        	} 
					                        	}
				                        	}
						        		}	
						        		break;
						        	case Cell.CELL_TYPE_NUMERIC:
						        		if ( cell.getColumnIndex() == 13 ) {
						        			if ( selectedOffices.get(0).equals("Все офисы") ) {
						        				column10 = true;
				                        	} else {		
					                        	for (int i = 0; i < selectedOffices.size(); i++) {
						                        	if ( String.valueOf( cell.getNumericCellValue() ) == selectedOffices.get(i).toString() ) {				                        		
						                        		column10 = true;
						                        	} 
					                        	}
				                        	}
						        		}
						        		break;	
					        		default:
					        			break;
						        }
					        }
			            }			        
				        if ( column1 == true && column7 == true && column9 == true && column10 == true && column234 == true) {
					        for (Cell cell: row) {
					        	if (row.getRowNum() > 1) {				        		
						        	switch ( cell.getCellTypeEnum() ) {
						        		case STRING:
						        			data.add( cell.getStringCellValue() );
						        			break;
						        		case NUMERIC:
						        			data.add( (int) cell.getNumericCellValue() );
						        			break;
						        		default:
						        			if ( cell.getColumnIndex() <= (mainHeaders.length - 2) ) {
						        				data.add("");
						        			}
						        			break;
						        	}	
						        	if ( cell.getColumnIndex() == (mainHeaders.length - 2) ) {
						        		data.add( cell.getRowIndex() );
						        		break;
						        	}			        			        	
		        				}
				        	} 
				        }
                    } 
					int cl = mainHeaders.length;
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
		            	// запрет на редактирование ячеек в таблице
		                private static final long serialVersionUID = 1L;                
		                // кнопку редактирования изменять можно
		                public boolean isCellEditable( int row, int column ) {                
		                	if ( column != mainHeaders.length - 1 ) {
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
		 	                return c;
		 		        }
		            };				    
				    
				    JTableHeader th1 = table1.getTableHeader();
		        	th1.setFont(new Font("Times New Roman", Font.BOLD, 12));     
		        	th1.setPreferredSize(new Dimension(50, 120)); 
		        	
		        	// горизонтальная прокрутка заголовков
		        	table1.getTableHeader().setPreferredSize(new Dimension(10000,120));		
		        	
		            // кнопка "редактировать"    
		            table1.getColumn(" ").setCellRenderer(new ButtonRendererTechnics (frame));
		            table1.getColumn(" ").setCellEditor(new ButtonEditorTechnics (new JCheckBox(), frame));
		            table1.getColumnModel().getColumn(mainHeaders.length - 1).setPreferredWidth(30);		            
		            // кнопка "редактировать"
		            
		            table1.changeSelection(0, 0, false, false);
		            JScrollPane scrollPane1 = new JScrollPane( table1 );
		            getContentPane().add(scrollPane1);		            
		        	
		            table1.setRowHeight(25);
		            
		        	for (int i = 0; i <= mainHeaders.length - 2; i++) {
				        switch (i) {
				        	case 0:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(35);
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 4:
				        	case 7:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(190); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 5:
				        	case 8:
				        	case 14:
				        	case 15:
				        	case 9:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(110); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 6:
				        	case 13:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(70); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        	case 1:
				        	case 2:
				        	case 3:
				        	case 10:	
				        	case 11:
				        	case 12:
				        		table1.getColumnModel().getColumn(i).setPreferredWidth(150); 
				        		table1.getColumnModel().getColumn(i).setCellRenderer( new MultilineTableCellRenderer() );
				        		break;
				        }
		        	}
		            
				    workbook.close();
				    
				    table1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
				    
				    table1.setDefaultRenderer(String.class, new MultilineTableCellRenderer());
		            
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
                	int columnCount = 0;
                	if ( mainHeaders[ mainHeaders.length - 1 ].equals(" ") ) {
                		columnCount = table1.getColumnCount() - 1;
                	} else {
                		columnCount = table1.getColumnCount();
                	}
                	if (fieldSub.isSelected() == true) {
                		new SelectingColumnsTechnics().selecting(jointUpload, columnCount );
				    } else {
				    	jointUpload.clear();
				    	new SelectingColumnsTechnics().selecting(copyData, columnCount );
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
    public static void main ( ) {
    	
    }
}

//создание кнопки "редактировать"
class ButtonRendererTechnics extends JButton implements TableCellRenderer {
	
	JFrame frame;
	int i = 0;
	
	public ButtonRendererTechnics (JFrame frame) {
	 	this.frame = frame;	
	    setOpaque(true);
	}

	public Component getTableCellRendererComponent(JTable table, Object value,
       boolean isSelected, boolean hasFocus, int row, int column) {
	
	     GridBagConstraints gbc = new GridBagConstraints();
	     gbc.gridwidth = GridBagConstraints.REMAINDER;
	     gbc.fill = GridBagConstraints.HORIZONTAL;
	
	     ImageIcon pencil = null;
	     pencil = new ImageIcon(new Technics().getClass().getClassLoader().getResource("pencil.png"));
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
class ButtonEditorTechnics extends DefaultCellEditor {
	
	public JButton button;
	String label = "";
	JFrame frame;
	public boolean isPushed;
	
	 public ButtonEditorTechnics(JCheckBox checkBox, JFrame frame) {
	     super(checkBox);
	     this.frame = frame;
	     button = new JButton(label);
	     button.setOpaque(true);
	     
	     button.addActionListener(new ActionListener() {
	         public void actionPerformed(ActionEvent e) {
	         	 fireEditingStopped();  
	         	 frame.dispose();
	          	 frame.setVisible(true);
	          	 frame.revalidate();
	         }
	     });
	 }

	 public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
	 	
		 //frame.dispose();
	 	 //frame.setVisible(true);
	 	 // frame.revalidate();
	 	 label = "";
	     button = new JButton(label);
	     GridBagConstraints gbc = new GridBagConstraints();
	     gbc.gridwidth = GridBagConstraints.REMAINDER;
	     gbc.fill = GridBagConstraints.HORIZONTAL;
	
	     ImageIcon pencil = null;
	
	     pencil = new ImageIcon( new Technics().getClass().getClassLoader().getResource("pencil.png") );
	     Image image = pencil.getImage();
	     Image newimg = image.getScaledInstance(23, 23, java.awt.Image.SCALE_SMOOTH);
	     pencil = new ImageIcon(newimg);
	
	     button.setBorderPainted(false);
	     button.setBorder(new LineBorder(Color.BLACK));
	
	     if (isSelected) {
	         button.setForeground(table.getSelectionForeground());
	         button.setBackground(table.getSelectionBackground());
	     } else {
	     	 //frame.dispose();
	     	 //frame.setVisible(true);
	     	 //frame.revalidate();
	     	 button.setForeground(table.getSelectionForeground());
	         button.setBackground(table.getSelectionBackground());
	     }
	     
	     button.setBackground(Color.white);
	     button.setIcon(pencil);
	     label = (value == null) ? "" : value.toString();	
	     isPushed = true;
	     TableModel tm = table.getModel();
	     	     
	     String[] inputValue = new String[table.getColumnCount()];
	     
	     for (int i = 0; i < inputValue.length; i++) {
	         inputValue[i] = (String) tm.getValueAt(row, i);
	     }	       	     	     
		 new EditButtonTechnics().windowDataChange( inputValue );	     	
		     
	     isPushed = true;
	     return button;
	 }
	 public Object getCellEditorValue() {		 
		 label = "";
	     isPushed = false;
	     //frame.dispose();
	 	 //frame.setVisible(true);
	 	 //frame.revalidate();
	     return label;
	 }
	 public boolean stopCellEditing() {		 
	     isPushed = true;
	     return super.stopCellEditing();
	 }
}