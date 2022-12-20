package net.codejava;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdesktop.swingx.JXDatePicker;
import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.net.URI;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;

public class AddDataDC {
    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
    private static final byte[] BUFFER = new byte[4096 * 1024];
    /*
	ImageIcon liderIcon = new ImageIcon(new Cars().getClass().getClassLoader().getResource(".png"));
    Image image = liderIcon.getImage();   
    */
    
    public ArrayList inputValues(ArrayList appointment) throws IOException {
    	ArrayList data = new ArrayList();
        JPanel panel = new JPanel(new GridLayout(18, 2, 5, 5));
        panel.setPreferredSize(new Dimension(700, 700));
        
        String[] items = {"№: ", "Вид контроля: ",  "Назначение (область применения): ", "Наименование прибора: ", "Тип, марка, модель: ", 
        		"<html><center>Производитель, страна производства, марка,<br> модель, основные технические характеристики: </html>",
        		"Зав.№: ", "Количество: ", "Год выпуска: ", "Дата поверки (калибровки): ","Дата окончания поверки (калибровки): ", "Документы: ", 
        		"Техническое состояние: ", "Указание в поверке на принадлежность к организации: ", 
       		     "Форма собственности: ", "Владелец оборудавания: ", "Местонахождение: ", "Примечание: "};

        JTextField[] fields = new JTextField[items.length];        
        String[] headersString = new String[appointment.size()];
        
        for (int i = 0; i < headersString.length; i++) {
        	headersString[i] = appointment.get(i).toString();
        }
        
        JComboBox fieldZ = new JComboBox(headersString);       
        JXDatePicker picker = new JXDatePicker();
        JXDatePicker picker2 = new JXDatePicker();
        SimpleDateFormat sDT = new SimpleDateFormat("dd.MM.yyyy");
        
        for (int i = 0; i < items.length; i++) { 
        	// собираем панель для добавления строки
        	switch(i) {
                default:
    	        	fields[i] = new JTextField(20);	          			
        			panel.add(new JLabel(items[i], SwingConstants.RIGHT)).setFont(bigFontTR);
        			panel.add(fields[i]).setFont(bigFontTR);
        			break;
        	}        		
        }        
        
        JButton buttonAdd = new JButton("Добавить");
        buttonAdd.setFont(bigFontTR);                
        buttonAdd.setPreferredSize(new Dimension(200, 30));
        
        JPanel panelNew = new JPanel(new BorderLayout(1, 1));
        panelNew.add(new Label(" "), BorderLayout.WEST);
        panelNew.add(panel, BorderLayout.CENTER);

        JPanel panelBt = new JPanel();
        panelBt.add(new Label(" "), BorderLayout.WEST);
        panelBt.add(buttonAdd, BorderLayout.CENTER);
        panelBt.add(new Label(" "), BorderLayout.EAST);
        panelBt.setPreferredSize(new Dimension(80, 30));
        
        JPanel jPanelAdd = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();

        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 0;
        
        c.weightx = 1;
        c.weighty = 1;
        c.fill = GridBagConstraints.BOTH;
        jPanelAdd.add(panelNew, c);
        
        c.fill = GridBagConstraints.VERTICAL;
        c.gridx = 1;
        c.gridy = 1;
        
        c.weightx = 0.1;
        c.weighty = 1;
        c.fill = GridBagConstraints.BOTH;
        
        jPanelAdd.add(panelBt, c);
        
        JFrame frameAdd = new JFrame();
        
        frameAdd.add(jPanelAdd);
        // frameAdd.setIconImage(image);
        frameAdd.setPreferredSize(new Dimension(900, 790));   
        frameAdd.pack();    
        frameAdd.setLocationRelativeTo(null);
        frameAdd.setVisible(true);       
        
        buttonAdd.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	frameAdd.setVisible(false);
                
                for (int i = 0; i < items.length; i++) { 
                	// собираем панель для добавления строки
                	switch(i) {
            	        default:
            	        	data.add(fields[i].getText());
            	        	break;
                	}       		
                }           
                
                try {
                    writeValues(data);
                } catch (FileNotFoundException e1) {
                    e1.printStackTrace();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }            
        });
        return data;
    }

    public void writeValues(ArrayList data) throws IOException {
        
        try {        	           
            int nRow = 1;
            int nClm = data.size() - 1;
            Object[][] dt = new Object[nRow][nClm];
            int sizeData = 0;
            for (int i = 0; i < nRow; i++) {
                for (int j = 0; j < nClm; j++) {
                    dt[i][j] = data.get(sizeData);
                    sizeData++;
                }
            }
            
            InputStream myxls = new FileInputStream("Reestr.xls");
            HSSFWorkbook sh = new HSSFWorkbook(myxls);
            Sheet worksheet = sh.getSheetAt(0);
            
            HSSFCell cell = null;
                       
            int newRow = new SearchHeaders().search(data.get(2).toString());
            
            Row row ;
            if (newRow <=  worksheet.getLastRowNum()) { 
            	worksheet.shiftRows(newRow, worksheet.getLastRowNum(), 1, true, false);
            }	
            
            row = worksheet.createRow(newRow);
            row = worksheet.getRow(newRow);
            worksheet.removeRow(worksheet.getRow(newRow));

            HSSFCellStyle style = sh.createCellStyle();
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment  .CENTER);
            style.setBorderBottom(style.getBorderBottom());
            style.setBorderTop(style.getBorderRight());
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            
            row = worksheet.createRow(newRow);
            
            for (int i = 0; i <= 8; i++) {
                cell = (HSSFCell) row.createCell(i);
                row.createCell(i).setCellValue(data.get(i).toString());
                row.getCell(i).setCellStyle(style);
            }
            
            for (int i = 9; i <= 10; i++) {
            	
            	String regex = "(\\d{2}.\\d{2}.\\d{4})";
        		Matcher m = Pattern.compile(regex).matcher(data.get(i).toString());
        		
        		if (m.find()) {
	                cell = (HSSFCell) row.createCell(i);
	                row.getCell(i).setCellStyle(style);
	                String s = data.get(i).toString();
	                
	                if (s.length() != 0) {
		                SimpleDateFormat format = new SimpleDateFormat();
		                format.applyPattern("dd.MM.yyyy");
		                Date docDate= format.parse(s);
		                row.createCell(i).setCellValue(format.format(docDate));
		                row.getCell(i).setCellStyle(style);
	                } else {		                
			           row.createCell(i).setCellValue("");
			           row.getCell(i).setCellStyle(style);
		            }
        		} else {
        			cell = (HSSFCell) row.createCell(i);
                    row.createCell(i).setCellValue((String) data.get(i));
                    row.getCell(i).setCellStyle(style);
        		}
            }
            
            for (int i = 11; i < data.size(); i++) {
                cell = (HSSFCell) row.createCell(i);
                row.createCell(i).setCellValue((String) data.get(i));
                row.getCell(i).setCellStyle(style);
            }
            myxls.close();
            
            java.nio.file.Path path1 = FileSystems.getDefault().getPath("Reestr.xls");
            // убрали скрытие файла - изменили значения - скрыли файл
            Files.setAttribute(path1, "dos:hidden", false);
            FileOutputStream outputStream;
                try {
                    outputStream = new FileOutputStream("Reestr.xls");
                    try {
                        sh.write(outputStream);
                        sh.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    outputStream.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
           Files.setAttribute(path1, "dos:hidden", true);
           

        } catch (IOException ex1) {
            System.out.println("Error reading file");
            System.out.println(ex1);
            ex1.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }

        System.out.println("Is saved in Excel file.");
    }
    public static void copy(InputStream input, OutputStream output) throws IOException {
        int bytesRead;
        while ((bytesRead = input.read(BUFFER))!= -1) {
            output.write(BUFFER, 0, bytesRead);
        }
    }   
}
