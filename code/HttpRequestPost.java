package net.codejava;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.Charset;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.TimeUnit;

import javax.servlet.ServletRequest;
import javax.servlet.http.HttpServletRequest;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

// ����� ��� �������� �������� � ���������� ������� ��������
class ReportGenerator extends TimerTask {
	 JSONObject jsonObject;
	
	 static String fileDictName = "";
	 ReportGenerator (JSONObject jsonObject) {
		 this.jsonObject = jsonObject;
	 }
	
	  public void run() {
			URL url;
			HttpURLConnection conn;
			OutputStream os = null;
			InputStreamReader isR = null;
			BufferedReader bfR = null;
			StringBuilder sB = new StringBuilder();
			try {
	            byte[] out = jsonObject.toString().getBytes("UTF-8");
				url = new URL("http://Officele.ru/pages/JSONeq.php");
				conn = (HttpURLConnection)url.openConnection();
	            conn.setRequestMethod("POST");
	            conn.setDoOutput(true);
	            conn.setDoInput(true);
	            conn.addRequestProperty("User-Agent", "Mozilla/94.0.2");
	            conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
	            // conn.setConnectTimeout(200);
	            // conn.setReadTimeout(200);
	            conn.connect();
	            try {
	            	os = conn.getOutputStream();
	                os.write(out);                
	            } catch (Exception e) {
	            	System.err.println(e.getMessage());
	            }
	            
	            if (conn.HTTP_OK == conn.getResponseCode()) {      
	            	// ����������
	            	isR = new InputStreamReader(conn.getInputStream(), "UTF-8");
	            	bfR = new BufferedReader(isR);
	            	String line;     	         	 
	          	  	while ((line = bfR.readLine())!=null) {
	          	  		sB.append(line);
	          	  	}  
	                isR.close();
	            }	            
			} catch (MalformedURLException e) {
				e.printStackTrace();
			} catch (IOException e) {
	        	System.err.println(e);
	        }  finally {
	        	// ��������� ��
	        	try {
	        		isR.close();
	        	} catch (IOException e) {
	            	System.err.println(e);
	            } try {
	        		bfR.close();
	        	} catch (IOException e) {
	            	System.err.println(e);
	            } try {
	        		os.close();
	        	} catch (IOException e) {
	            	System.err.println(e);
	            }       	
	        }
	  }
	      
}
public class HttpRequestPost {
	
	public void givenUsingTimer_whenSchedulingDailyTask_thenCorrect(JSONObject jsonObject, int count) throws IOException, ParseException {

		try {
			InputStream myxls = new FileInputStream("Reestr.xls");
			HSSFWorkbook  wb = new HSSFWorkbook(myxls);
			String dateCurrent;
			dateCurrent = wb.getSheetAt(0).getRow(0).getCell(14).getStringCellValue();
			
			String dataNow = new SimpleDateFormat("dd.MM.yyyy").format(Calendar.getInstance().getTime());
			if (dateCurrent.equals(dataNow)) {
			} else {
				new ReportGenerator(jsonObject).run();
				if (count == 2) { 
					HSSFWorkbook sh = new HSSFWorkbook(myxls);
					Sheet worksheet = sh.getSheetAt(0);
		            Row row;
		            row = worksheet.getRow(0);
		            row.createCell(2).setCellValue(dataNow);
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
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}             
        catch(IOException ex){             
            System.out.println(ex.getMessage());
        } 
	}
}
