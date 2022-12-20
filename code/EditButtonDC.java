package net.codejava;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.jdesktop.swingx.JXDatePicker;

import com.mysql.cj.protocol.a.NativeConstants.IntegerDataType;

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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EditButtonDC extends JFrame{

    Font bigFontTR = new Font("TimesRoman", Font.BOLD + Font.ITALIC, 14);     
    private static final byte[] BUFFER = new byte[4096 * 1024];
    
    @SuppressWarnings("unused")
	public ArrayList windowDataChange(String[] initialData) {   
    	
    	ArrayList data = new ArrayList();
    	
        String[] items = {"№: ", "Вид контроля: ",  "Назначение (область применения): ", "Наименование прибора: ", "Тип, марка, модель: ", 
        		"<html><center>Производитель, страна производства, марка,<br> модель, основные технические характеристики: </html>",
        		"Зав.№: ", "Количество: ", "Год выпуска: ", "Дата поверки (калибровки): ","Дата окончания поверки (калибровки): ", "Документы: ", 
        		"Техническое состояние: ", "Указание в поверке на принадлежность к организации: ", 
       		     "Форма собственности: ", "Владелец оборудавания: ", "Местонахождение: ", "Примечание: "};
    	
        JPanel panel = new JPanel(new GridLayout(18, 2, 5, 5));        
        JTextField[] fields = new JTextField[items.length];       
        JXDatePicker picker = new JXDatePicker();
        JXDatePicker picker2 = new JXDatePicker();
        SimpleDateFormat sDT = new SimpleDateFormat("dd.MM.yyyy");
        
        for (int i = 0; i < items.length; i++) { 
        	// собираем панель для добавления строки
        	switch(i) {
    				
        		case(9):
        			fields[i] = new JTextField(initialData[i], 20);	 
                	picker.setDate(Calendar.getInstance().getTime());
                	picker.setFormats(sDT);
                	panel.add(new JLabel(items[9], SwingConstants.RIGHT)).setFont(bigFontTR);
                	panel.add(fields[i]).setFont(bigFontTR);
                	break;
                	
        		case(10):
        			fields[i] = new JTextField(initialData[i], 20);	 
    	        	int x = 5;
    	        	Calendar cal = Calendar.getInstance();
    	        	cal.add( Calendar.YEAR, x);
    	        	Date dateNew = cal.getTime();
    	        	picker2.setDate(dateNew);
    	        	picker2.setFormats(sDT);
    	        	cal.add( Calendar.YEAR, -x);
    	        	panel.add(new JLabel(items[10], SwingConstants.RIGHT)).setFont(bigFontTR);
                	panel.add(fields[i]).setFont(bigFontTR);
                	break;
    	        
    	        default:
    	        	fields[i] = new JTextField(initialData[i], 20);	          			
        			panel.add(new JLabel(items[i], SwingConstants.RIGHT)).setFont(bigFontTR);
        			panel.add(fields[i]).setFont(bigFontTR);
        			break;
        	}        		
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
        
        panelBt.setPreferredSize(new Dimension(700, 30));
        
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
        ImageIcon liderIcon = new ImageIcon(new EditButtonDC().getClass().getClassLoader().getResource(".png"));
        Image image = liderIcon.getImage();   
        frame.setIconImage(image);   
        */   
        frame.setPreferredSize(new Dimension(900, 740));
        frame.add(jPanelEdit);
        frame.pack();
        frame.setVisible(true);

        jButtonEditing.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	frame.setVisible(false);
            	
                for (int i = 0; i < items.length; i++) { 
                	// собираем панель для добавления строки
                	switch(i) {
            	        default:
            	        	data.add(fields[i].getText());
                	}        		
                }  
                data.add(initialData[initialData.length - 1]);
                
                try {
                    new EditButtonDC().writeChangeData(data);
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
            	String nRow = initialData[18];
            	try {
					new EditButtonDC().deleteChangeData(nRow);
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
            HSSFSheet worksheet = sh.getSheetAt(0);
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

            row = worksheet.createRow(newRow);
            
            for (int i = 0; i <= 8; i++) {
                cell = (HSSFCell) row.createCell(i);
                row.createCell(i).setCellValue((String) data.get(i));
                row.getCell(i).setCellStyle(style);
            }
            
            for (int i = 9; i <= 10; i++) {    
            	String regex = "(\\d{2}.\\d{2}.\\d{4})";
        		Matcher m = Pattern.compile(regex).matcher(data.get(i).toString());
        		
        		if (m.find()) {
	                cell = (HSSFCell) row.createCell(i);
	                String s = data.get(i).toString();
	                row.getCell(i).setCellStyle(style);
	                
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

            for (int i = 11; i <= data.size() - 2; i++) {
            	 cell = (HSSFCell) row.createCell(i);
                 row.createCell(i).setCellValue((String) data.get(i));
                 row.getCell(i).setCellStyle(style);
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
            HSSFSheet worksheet = sh.getSheetAt(0);
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

