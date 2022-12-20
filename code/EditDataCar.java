package net.codejava;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;
import org.jdesktop.swingx.JXDatePicker;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

public class EditDataCar extends JFrame {

    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);     
    private static final byte[] BUFFER = new byte[4096 * 1024];
    
    @SuppressWarnings("unused")
	public ArrayList windowDataChange(String[] initialData) {    	
    	ArrayList data = new ArrayList();
    	
        String[] items = {"№: ", "Марка, модель: ", "VIN: ", "Регистрационный номер: ", "Год выпуска: ", "Пробег: ", "ПТС: ", "СТС: ",
    		    "Страховая компания (КАСКО): ", "№ полиса (КАСКО): ", "Срок действия (КАСКО): ", "Страховая компания (ОСАГО): ", 
    		    "№ полиса (ОСАГО): ", "Срок действия (ОСАГО): ", "Техническое состояние: ", 
    		    "Форма собственности: ", "Владелец оборудавания: ", "Местонахождение: ", "Ответственный владелец: ", "Примечание: "};
    	
        JPanel panel = new JPanel(new GridLayout(items.length, 2, 5, 5));
        
        JTextField[] fields = new JTextField[items.length];
        
        // собираем панель для изменения строки
        for (int i = 0; i < items.length; i++) { 
	        	fields[i] = new JTextField(initialData[i], 20);	          	
		    	panel.add(new JLabel(items[i], SwingConstants.RIGHT)).setFont(bigFontTR);
		    	panel.add(fields[i]).setFont(bigFontTR);	       	
        }
               
        panel.setPreferredSize(new Dimension(700, 540));
        
        JPanel panelNew = new JPanel(new BorderLayout(1, 1));
        panelNew.add(new Label(" "), BorderLayout.WEST);
        panelNew.add(panel, BorderLayout.CENTER);
        
        JButton jButtonEditing = new JButton("Изменить");
        JButton jButtonDelete= new JButton("Удалить");
        
        jButtonEditing.setPreferredSize(new Dimension(200, 30));       
        jButtonDelete.setPreferredSize(new Dimension(200, 30));

        JPanel panelBt = new JPanel();
        panelBt.add(jButtonEditing).setFont(bigFontTR);
        panelBt.add(new Label(" "));
        panelBt.add(new Label(" "));
        panelBt.add(new Label(" "));
        panelBt.add(jButtonDelete).setFont(bigFontTR);       
        panelBt.setPreferredSize(new Dimension(700,30));
        
        JPanel jPanelEdit = new JPanel(new GridBagLayout());
        
        GridBagConstraints c2 = new GridBagConstraints();
        c2.fill = GridBagConstraints.VERTICAL;
        c2.gridx = 1;
        c2.gridy = 0;        
        c2.weightx = 1;
        c2.weighty = 1;
        c2.fill = GridBagConstraints.BOTH;
        jPanelEdit.add(panelNew, c2);
        
        c2.fill = GridBagConstraints.VERTICAL;
        c2.gridx = 1;
        c2.gridy = 1;      
        c2.weightx = 0.1;
        c2.weighty = 0.1;
        c2.fill = GridBagConstraints.BOTH;       
        jPanelEdit.add(panelBt, c2);
        
        JFrame frame = new JFrame();       
        /*
        ImageIcon liderIcon = new ImageIcon(new EditDataCar().getClass().getClassLoader().getResource("logoLider.png"));
        Image image = liderIcon.getImage();  
        frame.setIconImage(image);     
        */
       
        frame.setPreferredSize(new Dimension(900, 730));
        frame.add(jPanelEdit);
        frame.pack();
        frame.setVisible(true);

        jButtonEditing.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	
            	frame.setVisible(false);
                
            	 for (int i = 0; i < items.length; i++) {       
                 	data.add(fields[i].getText());
                 } 
            	 
            	 data.add(initialData[initialData.length - 1]);
                
                try {
                    new EditDataCar().writeChangeData(data);
                }
                catch (IOException ex) {
                    System.out.println(ex);
                } catch (ParseException e1) {
					e1.printStackTrace();
				}
            }
        });
        
        jButtonDelete.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	frame.setVisible(false);
            	String nRow = initialData[20];
            	try {
					new EditDataCar().deleteChangeData(nRow);
				} catch (IOException e1) {
					e1.printStackTrace();
				}
            }
        });               
        return data;
    }

    public void writeChangeData(ArrayList data) throws IOException, ParseException {
         
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
            HSSFSheet worksheet = sh.getSheetAt(1);
            HSSFCell cell = null;
            int newRow = Integer.parseInt(data.get(data.size() - 1).toString());
            Row row ;
            worksheet.shiftRows(newRow, worksheet.getLastRowNum(), 1, true, false);
            row = worksheet.createRow(newRow);
            row = worksheet.getRow(newRow);
            worksheet.removeRow(worksheet.getRow(newRow));
            removeRow(worksheet, newRow + 1);

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

            for (int i = 0; i <= data.size() - 2; i++) {
            	 cell = (HSSFCell) row.createCell(i);
                 row.createCell(i).setCellValue( (String) data.get(dataIndex) );
                 row.getCell(i).setCellStyle(style);
                 dataIndex++;
            }
            
            myxls.close();

            java.nio.file.Path path1 = FileSystems.getDefault().getPath("Reestr.xls");
            // делаем файл не скрытым и вновь скрываем его
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
    }
    
    public void deleteChangeData(String values) throws IOException {
    	
            InputStream myxls = new FileInputStream("Reestr.xls");
            HSSFWorkbook sh = new HSSFWorkbook(myxls);
            HSSFSheet worksheet = sh.getSheetAt(1);
            HSSFCell cell = null;
            int newRow = Integer.parseInt(values);
            Row row ;
            
            row = worksheet.createRow(newRow);
            row = worksheet.getRow(newRow);
            removeRow(worksheet, newRow+1);
            
            myxls.close();
            
            java.nio.file.Path path1 = FileSystems.getDefault().getPath("Reestr.xls");
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
    }

    public static void removeRow(HSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex, lastRowNum, -1);
        }       
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
    
    public static void copy(InputStream input, OutputStream output) throws IOException {
        int bytesRead;
        
        while ((bytesRead = input.read(BUFFER))!= -1) {
            output.write(BUFFER, 0, bytesRead);
        }
    }
}


