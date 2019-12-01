package com.ils.driver;

import java.awt.AWTException;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import javax.script.ScriptException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Driver_Master {

	public static void main(String[] args) throws InvalidFormatException, IOException, AWTException, UnsupportedFlavorException, ScriptException {

		//Find user name and get path for Driver_Excel file
		String userName = System.getProperty("user.name");
		//String driverMasterPath = "T:\\IPOD\\Selenium\\1_UserBatch\\" +userName + "\\Driver_Excel.xlsx";
		
		String current_path = System.getProperty("user.dir");
		File Excel_File = new File(current_path + "\\Driver_Excel.xlsx");
		FileInputStream fis = new FileInputStream(Excel_File);
		Workbook excelbook = WorkbookFactory.create(fis);
		Sheet sheet = excelbook.getSheet("Sheet1");
		int lst_row = sheet.getLastRowNum();
		String releaseNbr = "";
		String xlsEnv = "";
		String[] retArray = null;
		
		
		for (int i = 0; i <= lst_row; i++) {
			String flag = sheet.getRow(i).getCell(0).getStringCellValue();
			if (flag.equalsIgnoreCase("Y") || flag.equalsIgnoreCase("B")) {
				String EXL_NAME = sheet.getRow(i).getCell(1).getStringCellValue();
				
				try {
					xlsEnv = sheet.getRow(i).getCell(2).getStringCellValue();
				} catch (Exception e) {
					xlsEnv = "NULL";
				}
				
				try {
				releaseNbr = sheet.getRow(i).getCell(3).getStringCellValue();
				} catch (Exception e) {
					releaseNbr = "NoRelNbr";
				}
				
				
				System.out.println("Now we are going to work on Excel ----->>>>" + EXL_NAME);
				
				
				try {
					retArray = Test_Cases.Web_call(EXL_NAME, xlsEnv, releaseNbr, userName);
				} catch (InterruptedException e) {
					
					e.printStackTrace();
				}
				
				if (flag.equalsIgnoreCase("Y")) {
					//Code to update driver sheet with status
					DateTimeFormatter formatter =  DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
					String dateStamp = LocalDateTime.now().format(formatter);
					
							
					//passed
					if (retArray[0] == "0") {
						sheet.getRow(i).createCell(5).setCellValue("PASSED");
					} else if (retArray[1].length() > 0 && retArray[2].length() == 0) { //failed without defects	
						sheet.getRow(i).createCell(5).setCellValue("FAILED");
					} else { //needs review, has failed with defects
						sheet.getRow(i).createCell(5).setCellValue("REVIEW");
					}
					//Write Failed steps
					sheet.getRow(i).createCell(6).setCellValue(retArray[0]);
					
					//Write lines failed with defects
					String failSteps = "Fail Steps : "+retArray[1];
					String bugSteps  = "Bug Steps  : "+retArray[2];
					String failBug = failSteps + "\r\n" + bugSteps;
					sheet.getRow(i).createCell(7).setCellValue(failBug);
					
					//Write Bug List
					String bugList = "Jira bug(s) :" + retArray[3];
					sheet.getRow(i).createCell(8).setCellValue(bugList);
					
					if ( i == lst_row) { //last test close input stream and write output file
						fis.close();
						try {
							FileOutputStream fos = new FileOutputStream(Excel_File);
							excelbook.write(fos);
							fos.close();
						} 
				        catch (Exception e) { 
				            e.printStackTrace(); 
				        }
					}    
				}
			}

		}
		
	
	}

}
