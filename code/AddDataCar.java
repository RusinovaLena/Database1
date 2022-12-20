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
import java.util.zip.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;

public class AddDataCar {
	// устанавливаем фрифт
    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);
    private static final byte[] BUFFER = new byte[4096 * 1024];
    
    // устанавливаем иконку панели
	/*
	ImageIcon liderIcon = new ImageIcon(new Cars().getClass().getClassLoader().getResource(".png"));
    Image image = liderIcon.getImage();
    */   

    public ArrayList inputValues() throws IOException {
    	ArrayList data = new ArrayList();
    	
        String[] items = {"№: ", "Марка, модель: ", "VIN: ", "Регистрационный номер: ", "Год выпуска: ", "Пробег: ", "ПТС: ", "СТС: ",
    		    "Страховая компания (КАСКО): ", "№ полиса (КАСКО): ", "Срок действия (КАСКО): ", "Страховая компания (ОСАГО): ", 
    		    "№ полиса (ОСАГО): ", "Срок действия (ОСАГО): ", "Техническое состояние: ", 
    		    "Форма собственности: ", "Владелец оборудавания: ", "Местонахождение: ", "Ответственный владелец: ", "Примечание: "};
        
        JPanel panel = new JPanel(new GridLayout(items.length, 2, 5, 5));
        panel.setPreferredSize(new Dimension(700, 700));                
        JTextField[] fields = new JTextField[items.length];
        
        for (int i = 0; i < items.length; i++) { 
        	fields[i] = new JTextField(20);	  
	    	panel.add(new JLabel(items[i], SwingConstants.RIGHT)).setFont(bigFontTR);
	    	panel.add(fields[i]).setFont(bigFontTR);	
        }    
        
        JButton buttonAdd = new JButton("Добавить");
        buttonAdd.setFont(bigFontTR);               
        buttonAdd.setPreferredSize(new Dimension(200, 30));
              
        JPanel panelNew = new JPanel( new BorderLayout(1, 1) );
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
                	data.add(fields[i].getText());
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
            int nClm = data.size();
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
            Sheet worksheet = sh.getSheetAt(1);
            
            HSSFCell cell = null;                       
            int newRow = worksheet.getLastRowNum() + 1;
            
            Row row ;            
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
            
            int dataIndex = 0;
            row = worksheet.createRow(newRow);
            for (int i = 0; i < data.size(); i++) {
                cell = (HSSFCell) row.createCell(i);
                row.createCell(i).setCellValue((String) data.get(dataIndex));
                row.getCell(i).setCellStyle(style);
                dataIndex++;
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
            System.out.println(ex1);
            ex1.printStackTrace();
        }
    }
    
    public static void copy(InputStream input, OutputStream output) throws IOException {
        int bytesRead;
        
        while ((bytesRead = input.read(BUFFER))!= -1) {
            output.write(BUFFER, 0, bytesRead);
        }
    }   
}

