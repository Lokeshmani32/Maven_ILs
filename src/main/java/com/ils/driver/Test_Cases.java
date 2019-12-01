package com.ils.driver;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.KeyEvent;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.script.*;
import javafx.scene.web.*;

import com.ils.genericmethods.Generic_Methods;





public class Test_Cases extends Generic_Methods {

	static String FCA = "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fh/ilss/";
	static String FCB = "http://www.newtours.demoaut.com/";
	static String FCD = "http://www.newtours.demoaut.com/mercuryregister.php";
	static String Backend=" https://ess-external-alb-1667722264.us-gov-west-1.elb.amazonaws.com/ess-backend/";
	static String Frontend=" https://ess-external-alb-1667722264.us-gov-west-1.elb.amazonaws.com/ess/";

	static String F0B= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/ess-backend/";
    static String F0F = " https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/ess/";
	static String FA="https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fa/ess-backend/";
	static String FB= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fb/ess-backend/";
	static String FC= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fc/ess-backend/";
	static String FD= " https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fd/ess-backend/";
	static String FE= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fe/ess-backend/";
	static String FF= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/ff/ess-backend/";
	static String FG = "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fg/ess-backend/";
	static String FH = "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fh/ilss/";
	static String FI= "https://ess-f0-f-1-alb-995328410.us-east-1.elb.amazonaws.com/fi/ess-backend/";
	static String fn_env = null;
	static String Browsername = "CH";
	
	public Test_Cases() throws IOException {
		super();

	}
	
		
	public static String[] Web_call(String Exl_Name, String xlsEnv, String releaseNbr, String userName) throws InvalidFormatException, IOException, InterruptedException, UnsupportedFlavorException, AWTException, ScriptException
	{
		String[][] varArray = new String[200][2];
		String varEnv = null;
		WebElement elementVar = null;
		int rowNbr = 0;
		String varName;
		int lineNbr;
		int failedSteps = 0;
		String xlsLineFailed = "";
		String xlsLineBug = "";
		String bugList = "Jira #(s) - ";
		 
		
		//Initialize varArray
		for (int iArr = 0; iArr < 200; iArr++) {;
		    for (int jArr = 0; jArr < 2; jArr++) {
			
		        varArray[iArr][jArr] = "";
		    }
		}
		
		//Get path for output xlsx file
		int indexLast = Exl_Name.lastIndexOf("\\");
		String xlsPath = Exl_Name.substring(0, indexLast+1);
		
		
		//Create file name for result file
		releaseNbr = releaseNbr.replaceAll("\\.", "_");
		DateTimeFormatter formatter =  DateTimeFormatter.ofPattern("yyyy-MM-dd HH_mm_ss");
		String fileDate = LocalDateTime.now().format(formatter);
		String resultXlsxFile = releaseNbr + "-" + xlsEnv + "-Result_" + fileDate + ".xls";
		
		//full path for result file
		String resultPathFile = xlsPath + resultXlsxFile;
		//String resultPathFile = "c:\\Olaf\\" + resultXlsxFile;
		
		//Test result file xls file
		//Workbook resultWorkbook = new XSSFWorkbook();
		Workbook resultWorkbook = new HSSFWorkbook();
		CreationHelper createHelper = resultWorkbook.getCreationHelper();
		
		//Create Sheets
		Sheet resultSheet = resultWorkbook.createSheet("Results");
		Sheet failsSheet = resultWorkbook.createSheet("Fails");
		Sheet varSheet = resultWorkbook.createSheet("Variables");
		Sheet[] sheetNamesArr = {resultSheet,failsSheet,varSheet};
		
		
		//Create Cell styles for result workbook
		Short col_num;
		CellStyle HeaderStyle = resultWorkbook.createCellStyle();
		col_num = IndexedColors.GREY_25_PERCENT.getIndex();
		HeaderStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		HeaderStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
		HeaderStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
		HeaderStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
		HeaderStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
		HeaderStyle.setFillForegroundColor(col_num);
	
	CellStyle NormalStyle = resultWorkbook.createCellStyle();
		NormalStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		NormalStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NormalStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NormalStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NormalStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		col_num = IndexedColors.WHITE.getIndex();
		NormalStyle.setFillForegroundColor(col_num);
		
	CellStyle PassStyle = resultWorkbook.createCellStyle();
		PassStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		PassStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		PassStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		PassStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		PassStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		col_num = IndexedColors.BRIGHT_GREEN.getIndex();
		PassStyle.setFillForegroundColor(col_num);	
		
	CellStyle FailStyle = resultWorkbook.createCellStyle();
		FailStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		FailStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		FailStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		FailStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		FailStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		col_num = IndexedColors.RED.getIndex();
		FailStyle.setFillForegroundColor(col_num);	
	
	CellStyle SkipStyle = resultWorkbook.createCellStyle();
		SkipStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		SkipStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		SkipStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		SkipStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		SkipStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		col_num = IndexedColors.LIGHT_YELLOW.getIndex();
		SkipStyle.setFillForegroundColor(col_num);	
		
		
		//Create Headers for Sheets (resultSheet and failsSheet)
 		for (int k = 0; k < 2; k++) {
			Row headerRow = sheetNamesArr[k].createRow(0);
			//Column A
			Cell cell = headerRow.createCell(0);
			cell.setCellValue("Test_Description");
			cell.setCellStyle(HeaderStyle);
			//Column B
			cell = headerRow.createCell(1);
			cell.setCellValue("Step Status");
			cell.setCellStyle(HeaderStyle);
			//Column B
			cell = headerRow.createCell(2);
			cell.setCellValue("Test Sheet Row");
			cell.setCellStyle(HeaderStyle);
			//Column C
			cell = headerRow.createCell(3);
			cell.setCellValue("Action");
			cell.setCellStyle(HeaderStyle);
			//Column D
			cell = headerRow.createCell(4);
			cell.setCellValue("Locator");
			cell.setCellStyle(HeaderStyle);
			//Column E
			cell = headerRow.createCell(5);
			cell.setCellValue("Value");
			cell.setCellStyle(HeaderStyle);
			//Column F
			cell = headerRow.createCell(6);
			cell.setCellValue("Execution Time");
			cell.setCellStyle(HeaderStyle);
			//Column G
			cell = headerRow.createCell(7);
			cell.setCellValue("Comments");
			cell.setCellStyle(HeaderStyle);
		}
		
		//Set headers for varSheet
		Row headerRow = varSheet.createRow(0);
		//Column A
		Cell cell = headerRow.createCell(0);
		cell.setCellValue("Variable Name");
		cell.setCellStyle(HeaderStyle);
		//Column B
		cell = headerRow.createCell(1);
		cell.setCellValue("Variable Value");
		cell.setCellStyle(HeaderStyle);
		
		//Test driver xlsx file
		File Excel_File = new File(Exl_Name);
		FileInputStream fis = new FileInputStream(Excel_File);
		Workbook excelbook = WorkbookFactory.create(fis);
		
		Sheet sheet = excelbook.getSheet("Sheet1");
		int lst_row = sheet.getLastRowNum();
		int lst_Cell = sheet.getRow(0).getLastCellNum();
		for (int rw = 0; rw <= lst_row; rw++) {
			for (int cl = 0; cl < lst_Cell; cl++) {
				if (sheet.getRow(rw).getCell(cl) == null) {
					sheet.getRow(rw).createCell(cl);
				}
				sheet.getRow(rw).getCell(cl).setCellType(1);
			}
		}
		
		
		
		for (int i = 0; i <= lst_row; i++) {
			String flag = sheet.getRow(i).getCell(1).getStringCellValue();
			rowNbr = rowNbr + 1;
			if (flag.equalsIgnoreCase("Y")) {
				String Test_Description = sheet.getRow(i).getCell(0).getStringCellValue();		
				String Action = sheet.getRow(i).getCell(2).getStringCellValue();
				String Locator = sheet.getRow(i).getCell(3).getStringCellValue();
				String Value = sheet.getRow(i).getCell(4).getStringCellValue();
				
				
				//Set lineNbr to be the same as the line in the test worksheet
				lineNbr = i + 1;	
				

				System.out.println("Action---" + Action);

				int startInd;
				int endInd;
				switch (Action) {

				case "Close":
					driver.quit();
					
					//Write passed line to result sheet
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' closed the Browser");
					
					//Write variables to result worksheet
					for (int k = 0; k < varArray.length; k++) {
						if (varArray[k][0].isEmpty()) {
							break;
						} else {
							Row varRow = varSheet.createRow(k+1);
							
							Cell varCell = varRow.createCell(0);
							varCell.setCellValue(varArray[k][0]);
							
							varCell = varRow.createCell(1);
							varCell.setCellValue(varArray[k][1]);
							
						}	
					}
				 		
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, "", "Test Complete", lineNbr, "", "Run By: " +userName, "Exe Env: " + xlsEnv, "Test Name: " + Exl_Name );
					WriteResults(NormalStyle, PassStyle, resultWorkbook, failsSheet, "", "Test Complete", lineNbr, "", "Run By: " +userName, "Exe Env: " + xlsEnv, "Test Name: " + Exl_Name );
					
					
					//Auto size all columns for all sheets
					for (int k = 0; k < 3; k++) {
						int totalCols = sheetNamesArr[k].getRow(0).getLastCellNum();
						for(int j = 0; j < totalCols; j++) {
							//set last column width is 80 
							if (j == totalCols-1 ) {
								sheetNamesArr[k].setColumnWidth(j, 80*256);
							} else if (j == 5) {
								sheetNamesArr[k].setColumnWidth(j, 40*256);
							} else {
								sheetNamesArr[k].autoSizeColumn(j);
							}
				        }	
					}
					
					//Write Result Spreadsheet
					FileOutputStream fileOut = new FileOutputStream(resultPathFile);
					resultWorkbook.write(fileOut);
			        fileOut.close();
			        
			        
					break;

				case "BrowserName":
					driver(Value);
					break;
					
				case "Url":
					/*used to select an environment and user for the test 
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action	Locator 	Value 
					 * y	URL		FH			essat.user12
					 * y	URL		env token	user name for test
					 * Var	N		N			N
					 */
					
					//Override work sheet enviroment with enviroment from xls sheet if populated
					if (xlsEnv != "NULL" ) {
						Locator = xlsEnv;
					}
					url(env(Locator));
					
					//Click Accept button
					driver.findElement(By.xpath("//button[@class='m-1 btn btn-primary']")).click();
					driver.manage().window().maximize();
					
					Thread.sleep(1000);
					
					//Switch to user specified  
					if (Value!="ess.user12") {
						driver.findElement(By.xpath("//li[@class='px-3 d-md-down-none nav-item']//button[@class='btn btn-link'][contains(text(),'ess.user12@SPO')]")).click();
						String xpath = "//button[contains(text(),'" + Value + "')]";
						driver.findElement(By.xpath(xpath)).click();
					}
					
					Thread.sleep(1000);
					
					//Set varEnv for use when looking environment variables 
					if (Locator.length() > 2) {
						varEnv = Locator.substring(0,1);
					} else {
						varEnv = Locator;
					}
					
					//Write passed line to result sheet
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' sucessfully navigated to the "+Locator+" enviroment");
					
					break;
					
				case "InputValue":
					
					//Parse locator and value from searchName
					String[] LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					String locator = LocatorArray[0];
					String locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for Value
					if (Value.substring(0,1).contentEquals("^")) {
						varName = Value.substring(1, Value.length());
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					
					try {
						elementVar = FindWebElement(driver, locator, locatorVal);
						//elementVar.clear();
						elementVar.sendKeys(Keys.CONTROL + "a");
						elementVar.sendKeys(Keys.DELETE);
						elementVar.sendKeys(Value);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' entered "+Value+" in the page object identified by " + Locator);
					break;
					
				case "ListItemSelect":
					/*used to select a drop down Item from a list 
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action			Locator 			Value 
					 * y	ListItemSelect	attribute#value		list item or items to be chosen
					 * y	ListItemSelect	html id#dodaacs		FB0488 - POPE NC
					 * Var	N				N:Y					Y
					 */
					
					//Parse locator and value from searchName
					LocatorArray  = Locator.split("#", 2);
					
					//User did not enter # separator for the Locator
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locator = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					//Lookup variable if first position is "^" for Value
					if (Value.substring(0,1).contentEquals("^")) {
						varName = Value.substring(1, Value.length());
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					
					Thread.sleep(1000);
					
					//Set page element based on locator and locatorVal
					try {
						elementVar = FindWebElement(driver, locator, locatorVal);
						//send keys to elementVar
						elementVar.sendKeys(Value);
						elementVar.sendKeys(Keys.RETURN);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' selected '"+Value+"' in the list identified by: " + Locator);
					
					break;
					
				case "Click":
					/*used to click on browser objects that allow the click operation
					 * e.g. buttons
					 * 
					 * Spreadsheet format:
					 * Run	Action		Locator (searchName)	inputString
					 * y	Click		attribute#value			N/A
					 */
					//System.out.println("Click steps");
					
					//Parse locator and value from searchName
					Thread.sleep(1000);
					
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locator = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					//Set page element based on locator and locatorVal
					try {
						elementVar = FindWebElement(driver, locator, locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					//Scroll to elementVar on page
					ScrollToView(driver, elementVar);
				
					//Click web page element
					elementVar.click();
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' clicked the button identified by: " + Locator);
					
					break;
					
				case "TabSwitch":
					//Used to switch to a newly opened tab 
					ArrayList<String> newTab = null;
					
					String currTab=driver.getWindowHandle();
					
					for (int jnewTab = 0; jnewTab < 11; jnewTab++) {
						Thread.sleep(1000);
						newTab = new ArrayList<String>(driver.getWindowHandles());
						if (newTab.size() > 1) {
							jnewTab = 11;
						}
					}
					
					
				    newTab.remove(currTab);
				    
				    try {
				    driver.switchTo().window(newTab.get(0));
				    } catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not find a page to switch to");
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not find a page to switch to");
						System.out.println("Tabswitch could not find a tab to switch to may need wait");
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
				    }
				    
				    String curUrl = driver.getCurrentUrl();
				    
				    WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' switched to the new tab with the following url: "+ curUrl);
					break;
					
				case "TabClose":
					//Used to close current tab and go back to original
					currTab=driver.getWindowHandle();
					newTab = new ArrayList<String>(driver.getWindowHandles());
				    newTab.remove(currTab);
				    
				    try {
				    driver.switchTo().window(currTab).close();
				    driver.switchTo().window(newTab.get(0));
				    } catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not find a page to switch to");
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not find a page to switch to");
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
				    }
				    
				    WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' closed the browser tab and returned to the original");
					break;
					
				case "Submit":
					try {
						getelement(Locator).submit();
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					Thread.sleep(1000);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' Submitted the page object identified by: " + Locator);
					break;
					
				
				case "Alert":
					Thread.sleep(1000);
					Alert AR = driver.switchTo().alert();
					AR.accept();
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' Closed the alert");
					break;
					
				case "Wait":
					
					int waitSecs = Integer.parseInt(Value)*1000;      
					Thread.sleep(waitSecs);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' waited for :" + Value + " second(s)");
					break;	
					
				case "Get":
					/*Used to get data from a screen element and store it in a variable    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	Get			xpath#//bla...		variable name
					 * Var	N			N					Y
					 */
					//Parse locator and value from searchName
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locator = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					//Set page element based on locator and locatorVal
					try {
						elementVar = FindWebElement(driver, locator, locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					String varValue = elementVar.getText();
					
					//Call SaveVar to store data
					SaveVar(Value,varValue, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' got the text from object identified by: " + Locator + " and stored it in the variable named : " + Value + " as the value: "+ varValue);
					
					break;

				case "PdfGet2":
					/*Used to get PDF data from a screen element and store it in a variable    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	Get			N/A					variable name
					 * Var	N			N					Y
					 */
					String result = "";
					//Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					
					WebDriverWait wait1 = new WebDriverWait(driver, 20);
					
					//elementVar = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//embed")));
					elementVar = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//html")));
					
					
					elementVar.sendKeys(Keys.CONTROL + "A");
					Thread.sleep(1000);
					
					
					elementVar.sendKeys(Keys.CONTROL + "c");
					Thread.sleep(7000);
					
					//Transferable contents = clipboard.getContents(null);
					//result = (String)contents.getTransferData(DataFlavor.stringFlavor);
					
					System.out.println("Result :"+result);
					break;
					
					
				case "PdfGet":
					/*Used to get PDF data from a screen element and store it in a variable    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	Get			N/A					variable name
					 * Var	N			N					Y
					 */
					//String result = "";
					JavascriptExecutor js = (JavascriptExecutor)driver;
					
					wait1 = new WebDriverWait(driver, 20);
					
					elementVar = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//embed")));
					
					elementVar.sendKeys(Keys.CONTROL + "A");
					Thread.sleep(1000);	
					
					
					  String script = 
					  "var activeEl = document.activeElement;"+
					  "var selTxt = activeEl.selection.text;"+
					  "return selTxt;";
					 					
								
					result = (String) js.executeScript(script);
					
					
					
					System.out.println("Result :"+result);
					//Get current URL
					/*
					 * String strURL = driver.getCurrentUrl();
					 * 
					 * URL url = new URL(strURL);
					 * 
					 * InputStream is = url.openStream(); BufferedInputStream fileToParse = new
					 * BufferedInputStream(is);
					 * 
					 * PDDocument document = PDDocument.load(fileToParse);
					 * 
					 * String output = new PDFTextStripper().getText(document);
					 */
					
					/*
					 * PDFTextStripper pdfStripper = null; PDDocument pdDoc = null; COSDocument
					 * cosDoc = null; String parsedText = null;
					 * 
					 * String getURL = driver.getCurrentUrl();
					 * 
					 * URL url = new URL(getURL); BufferedInputStream file = new
					 * BufferedInputStream(url.openStream()); PDFParser parser = new
					 * PDFParser((RandomAccessRead) file);
					 * 
					 * parser.parse(); cosDoc = parser.getDocument(); pdfStripper = new
					 * PDFTextStripper();
					 * 
					 * 
					 * parser.parse(); cosDoc = parser.getDocument(); pdfStripper = new
					 * PDFTextStripper(); pdfStripper.setStartPage(1); pdfStripper.setEndPage(1);
					 * 
					 * pdDoc = new PDDocument(cosDoc); result = pdfStripper.getText(pdDoc);
					 * 
					 */					
					
					
					
					
					
					
					
					//WebDriverWait wait1 = new WebDriverWait(driver, 20);
					
					//elementVar = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//embed")));
					
					//elementVar.sendKeys(Keys.CONTROL + "A");
					//Thread.sleep(1000);
					
					
					//robot = new Robot();
		            //robot.setAutoDelay(250);
		            //robot.keyPress(KeyEvent.VK_CONTROL);
		            //robot.keyPress(KeyEvent.VK_C);
		            //robot.keyRelease(KeyEvent.VK_C);
		            //robot.keyRelease(KeyEvent.VK_CONTROL);
		            
					//elementVar.sendKeys(Keys.CONTROL + "c");
					
					//Thread.sleep(3000);
					
		            //result = (String)Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
					
		            
		            //Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					//Transferable contents = clipboard.getContents(null);
					
					
					//result = (String) contents.getTransferData(DataFlavor.stringFlavor);
					
					//Call SaveVar to store data
					SaveVar(Value, result, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' got the text from object identified by: " + Locator + " and stored it in the variable named : " + Value + " as the value: "+ result);
					
					break;
					
					
				case "Defect":
					/*Used to get data from a screen element and store it in a variable    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				 	Value
					 * y	Defect		123456789  				NA
					 * Var	N			N						N
					 */
					//Parse locator and value from searchName
					
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The previous step '"
								+ (lineNbr-1) + "' has a identified defect, the defect is : "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The previous step '"
								+ (lineNbr-1) + "' has a identified defect, the defect is : "+Locator);
						xlsLineBug = xlsLineBug + String.format("%04d", lineNbr)+":";
						bugList = bugList + Locator + ", ";
					break;
					
				case "SwitchUser":
					/*Used to get data from a screen element and store it in a variable    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				 	Value
					 * y	SwitchUser	essat.user12@SPO		ess.user12
					 * Var	N			N						N
					 */
					//Parse locator and value from searchName
					
					locatorVal = "//button[contains(text(),'" + Locator + "')]";
					
					try {
						elementVar = FindWebElement(driver, "xpath" , locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					elementVar.click();
					
					locatorVal = "//button[contains(text(),'" + Value + "')]";
					
					try {
						elementVar = FindWebElement(driver, "xpath" , locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch to user identified by: "+Value);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch to user identified by: "+Value);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					elementVar.click();
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' changed the currently logged in usert from "+Locator+ "to "+Value);
					
					
					break;
	

				
				case "RandDoc":
					/*Generates 15 days old random document number (to allow for ordinal date of status to be old)    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator (searchDesc)		Value (inputString)
					 * y	RanDoc		<first 6 pos of DocNbr>		variable name
					 * Var	N			N					Y
					 */
					//
					
					//Get first 6 positions of random doc nbr
					//Lookup variable if first position is "^" for inputString
					if (Locator.length() > 0) {
						if (Locator.substring(0,1).contentEquals("^")) {
							varName = Locator.substring(1, Locator.length());
							Locator = LookupVar(varName,varArray,varEnv);
						}
					}	
					
					String randomDoc = Locator;
					
					//subtract 15 from current date
					formatter =  DateTimeFormatter.ofPattern("yDDD");
					LocalDateTime dateVar = LocalDateTime.now().minusDays(15);
					String jDate = dateVar.format(formatter);
					
					//Get date in yddd format
					jDate = jDate.substring(3, 7);
					
					//System.out.println(jDate);
					
					randomDoc = randomDoc+jDate;
					
					//Create a random 4 position serial number 
					
					int iRnd = 0;
					Random rand = new Random();
					
					while (iRnd < 1000) {
						String iStr = String.format("%04d", rand.nextInt(8000));
						iRnd = Integer.parseInt(iStr);
					}
					String serNbr = Integer.toString(iRnd);
					
					randomDoc = randomDoc+serNbr;
					//System.out.println(randomDoc);
					
					//Save in array
					SaveVar(Value,randomDoc, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' created a variable: " + Value + " and stored the following random document number: " + randomDoc );
					
		           break;
					
				case "RandDocCur":
					/*Generates current date document number     
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator (searchDesc)		Value (inputString)
					 * y	RanDoc		<first 6 pos of DocNbr>		variable name
					 * Var	N			N					Y
					 */
					//
					
					//Get first 6 positions of random doc nbr
					//Lookup variable if first position is "^" for inputString
					if (Locator.length() > 0) {
						if (Locator.substring(0,1).contentEquals("^")) {
							varName = Locator.substring(1, Locator.length());
							Locator = LookupVar(varName,varArray,varEnv);
						}
					}
					
					randomDoc = Locator;
					
					//subtract 15 from current date
					formatter =  DateTimeFormatter.ofPattern("yDDD");
					dateVar = LocalDateTime.now();
					jDate = dateVar.format(formatter);
					
					//Get date in yddd format
					jDate = jDate.substring(3, 7);
					
					//System.out.println(jDate);
					
					randomDoc = randomDoc+jDate;
					
					//Create a random 4 position serial number 
					
					iRnd = 0;
					rand = new Random();
					
					while (iRnd < 1000) {
						String iStr = String.format("%04d", rand.nextInt(8000));
						iRnd = Integer.parseInt(iStr);
					}
					serNbr = Integer.toString(iRnd);
					
					randomDoc = randomDoc+serNbr;
					//System.out.println(randomDoc);
					
					//Save in array
					SaveVar(Value,randomDoc, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' created a variable: " + Value + " and stored the following random document number: " + randomDoc );
					
		           break;
					
				
				case "Join":
					/*used to create a variable that can store a string or concatenate strings  
					 * and other variables using the "|" as a delimiter (variable names will 
					 * be prefixed with ^ to indicate they are variables) creating variable value    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	Join		variable name	 	string|^variable...
					 * y	Join		codeRollRespVar		r|^codeRollDodaac
					 * Var	N			N					Y
					 */
					varValue = ""; 
					String varStoreName = Locator;
					
					//Check for concatenation character "|" in inputString
					if(Value.indexOf("|")>-1) {
						//Steps concatenation
						//System.out.println("Concat Req'd");
						
						String workingStr = Value;
						
						//Using regex patternString escaped by "\\Q" for the regex meta "|" followed by "\\E"  
						String patternString = "\\Q|\\E";
						Pattern pattern = Pattern.compile(patternString);
						String[] strParts = pattern.split(workingStr);
						
						int m = strParts.length;
						int z = 0;
						Value = "";
						
						//Work individual parts of the string
						while (z < m) {
							
							if (strParts[z].substring(0,1).contentEquals("^")) {
								//strPart is a variable (starts with "^") so lookup variable and append
							
								varName = strParts[z].substring(1, strParts[z].length());
								
								Value = Value + LookupVar(varName,varArray,varEnv);
							} else {
								//strPart is not a variable so append it
								Value = Value + strParts[z];
							}
							
							z = z + 1;
							
							varValue = Value;
						}
						
					} else {
					   //Steps for no concatenation
					   //System.out.println("No Concat Req'd");
						
						//Check to see if variable and lookup
						if (Value.substring(0,1).contentEquals("^")) {
							varName = Value.substring(1, Value.length());
							Value = LookupVar(varName,varArray,varEnv);
						}
						
						varValue = Value;
						
					}
					
					//Call SaveVar to store data
					SaveVar(varStoreName,varValue, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' generated a variable: " + varStoreName + " with the value of: " + varValue);

					
					break;

				case "Append":
					/*used to create a variable that can store a string or concatenate two strings  
					 * or other variables using the ":" as a delimiter (variable names will 
					 * be prefixed with ^ to indicate they are variables) creating variable value    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	Append		variable name	 	string:^variable...
					 * y	Append		codeRollRespVar		r:^codeRollDodaac
					 * Var	N			N					Y:Y
					 */
					
					varValue = ""; 
					varStoreName = Locator;
					
					//Check for concatenation character "|" in inputString
					if(Value.indexOf(":")>-1) {
						//Steps concatenation
						//System.out.println("Concat Req'd");
						
						String workingStr = Value;
						
						//Using regex patternString escaped by "\\Q" for the regex meta "|" followed by "\\E"  
						String patternString = ":";
						Pattern pattern = Pattern.compile(patternString);
						String[] strParts = pattern.split(workingStr);
						
						int m = strParts.length;
						int z = 0;
						Value = "";
						
						//Work individual parts of the string
						while (z < m) {
							
							if (strParts[z].substring(0,1).contentEquals("^")) {
								varName = strParts[z].substring(1, strParts[z].length());
								Value = Value + LookupVar(varName,varArray,varEnv);
							} else {
								Value = Value + strParts[z];
							}
							
							z = z + 1;
							
							varValue = Value;
						}
						
					} else {
					   //Steps for no concatenation
					   //System.out.println("No Concat Req'd");
						
						//Check to see if variable and lookup
						if (Value.substring(0,1).contentEquals("^")) {
							varName = Value.substring(1, Value.length());
							Value = LookupVar(varName,varArray,varEnv);
						}
						
						varValue = Value;
						
					}
					
					//Call SaveVar to store data
					SaveVar(varStoreName,varValue, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' generated a variable: " + varStoreName + " with the value of: " + varValue);

					
					break;
				
				case "Parse":
					/*Used to parse data from a string using regex or string values 
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator (searchName)					Value (inputString)
					 * y	Join		ReSearchString							\$[0-9]*,*[0-9]*,*[0-9]*\.[0-9]{2}[- ]          CREDIT RETURNS
					 * y	Parse		^InqMacrData#inqMacrDataChk1			BC  FC  SD,1,NET DEMANDS,1		
					 * y	Parse		^InqMacrData#inqMacrDataChk2			^ReSearchString,1,PROCESSING COMPLETE,1											
					 * Var	N			Y#N							        	Y              ,N,Y                  ,N
					 * 					<text to be parsed>#<var save name>		begStr,occurrence,endStr,occurrence				 */
					startInd = 0;
					endInd = 0;
					
					//Parse locator and value from Locator
					LocatorArray  = Locator.split("#", 2);
					String parseStr = LocatorArray[0];
					String varNameStore = LocatorArray[1];
					
					
					//Lookup variable if first position is "^" for parseStr
					if (parseStr.substring(0,1).contentEquals("^")) {
						varName = parseStr.substring(1, parseStr.length());
						parseStr = LookupVar(varName,varArray,varEnv);
					}
					
					//Parse parts of inputString
					String[] ValueSplit = Value.split(",");
					String begStr = ValueSplit[0];
					int begOcc = Integer.parseInt(ValueSplit[1]);
					String endStr = ValueSplit[2];
					int endOcc = Integer.parseInt(ValueSplit[3]);
					
					//Get beginning string (begStr) can be a variable
					//Lookup variable if first position is "^" for inputString
					if (begStr.substring(0,1).contentEquals("^")) {
						varName = begStr.substring(1, begStr.length());
						begStr = LookupVar(varName,varArray,varEnv);								
					}
					
					
					//Get ending string (endStr) can be a variable
					//Lookup variable if first position is "^" for inputString
					if (endStr.substring(0,1).contentEquals("^")) {
						varName = endStr.substring(1, endStr.length());
						endStr = LookupVar(varName,varArray,varEnv);
					}
					
					
					//find startInd of string to be parsed
					Pattern pattern = Pattern.compile(begStr);
			        Matcher matcher = pattern.matcher(parseStr);

			        int count = 0;
			        while(matcher.find() && count < begOcc) {
			            count++;
			            startInd = matcher.start();
			        }
			        
			        //find endInd of string to be parsed
			        pattern = Pattern.compile(endStr);
			        matcher = pattern.matcher(parseStr);
			        
			        count = 0;
			        while(matcher.find() && count < endOcc) {
			            count++;
			            endInd = matcher.end();
			        }
			        
					//Parse searchName string
			        varValue = parseStr.substring(startInd,endInd);
			        
			        //Store parse string as searchDesc variable
			        SaveVar(varNameStore,varValue, varArray);
			        
			        //System.out.println("The variable " + searchDesc + " was saved with the following value: " + parseStr);

			        WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' parsed data and generated a variable: " + varNameStore + " with the value of: " + varValue);

			        break;

				case "SearchData":
					/*Used to search for a string of data and then store the data after search string + n positions 
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator 										Value 
					 * y	SearchData	7,OrdDate,ORDINAL DATE,1						^StatusRespTxt											
					 * Var	N			N,N,N,N											Y
					 * 					positions,<var save name>,string to locate,occ	variable or string to be searched
					 */
					
					//add occurrence of 1 if not in sheet
					int cntComma = StringUtils.countMatches(Locator, ",");
					if (cntComma == 2) {
						Locator = Locator + ",1";
					}
					
					//Parse nbrPos, varNameStore and searchString from Locator
					LocatorArray  = Locator.split(",", 4);
					int nbrPos = Integer.parseInt(LocatorArray[0]);
					varNameStore = LocatorArray[1];
					String searchString = LocatorArray[2];
					int occur = Integer.parseInt(LocatorArray[3]);
					
					//Lookup variable if first position is "^" for Value
					if (Value.substring(0,1).contentEquals("^")) {
						varName = Value.substring(1, Value.length());
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					//find startInd of string to be parsed
					pattern = Pattern.compile(searchString);
			        matcher = pattern.matcher(Value);

			        count = 0;
			        startInd = 0;
			        while(matcher.find() && count < occur) {
			            count++;
			            startInd = matcher.end();
			        }
			        
			        //Move the startInd to the right the length of searchString + 2 pos for " :"
			        startInd = startInd + 2;
			        
			        endInd = startInd + nbrPos;
			        
					//Parse searchName string
			        varValue = Value.substring(startInd,endInd);
			        
			        //Store parse string as searchDesc variable
			        SaveVar(varNameStore,varValue, varArray);
			        
			        //System.out.println("The variable " + searchDesc + " was saved with the following value: " + parseStr);

			        WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' parsed data and generated a variable: " + varNameStore + " with the value of: " + varValue);

			        break;

					
				case "StoreString":
					/*used to create a variable that is a substring of  (variable names will 
					 * be prefixed with ^ to indicate they are variables) creating variable value    
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action			Locator (searchDesc)			Value (inputString)
					 * y	StoreString		<s pos>,<# char>,varName		string or ^variable
					 * y	StoreString		1,6,codeRollDodaac				^codeRollBaseName
					 * Var	N				N								Y
					 */
					//insert steps
					
					//Parse the searchDesc using ","
					String[] partsSearchDesc = Locator.split(",");
					startInd = Integer.parseInt(partsSearchDesc[0])-1;
					endInd = Integer.parseInt(partsSearchDesc[1]);
					varNameStore = partsSearchDesc[2];
					
					//Get string to parse
					int m = Value.length();
					
					//Lookup variable if first position is "^"
					if (Value.substring(0,1).contentEquals("^")) {
						varName = Value.substring(1, Value.length());
						Value = LookupVar(varName,varArray, varEnv);
					} 
						
					varValue = Value;
					
					//If endInt > varValue.length then set endInt = 
					if(startInd + endInd >= varValue.length()) {
						endInd = varValue.length();
					}
					else {
							endInd = startInd + endInd;
					}
					
					
					//Parse VarValue based on startInd and endInd
					varValue = varValue.substring(startInd, endInd);
					
					//Call SaveVar to store data
					SaveVar(varNameStore,varValue, varArray); 
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' substringed data and generated a variable: " + varNameStore + " with the value of: " + varValue);

							
				break;
				
				case "FindByRowGet":
					/*Used to find a specific element on a dynamic table based on a unique value in that table 
					 * and store it in a variable
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Location (searchName)					Value (inputString)
					 * y	FindByRow	xpath#//div[@class='rt-tr-group']		^BatchId:4:someVariable											
					 * Var	N			N#N										Y:       N:N
					 * 					locator   								TextToFind:ColumnToSaveinfo:variable name 
					 * 
					 * ^BatchId:4 is the TextToFind:ColumnToLookUp (in this case ^BatchId is the "Total Submitted" Column and 
					 * ColumnToLookup is Column 4, stores text in someVariable (text in the specified column) 
					 * 
					 */
					WebDriverWait wait = new WebDriverWait(driver, 10);
					
					//Parse locator and value from Locator
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locatorVal = LocatorArray[0];
					String xpathStr = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (xpathStr.substring(0,1).contentEquals("^")) {
						varName = xpathStr.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locatorVal + "#" + xpathStr;
					}
					
					//split Value into parts (searchForStr, tableColumn and varNameStore)
					String[] ValueArr = Value.split(",",3); 
					String searchForStr = ValueArr[0];
					int tableColumn = Integer.parseInt(ValueArr[1]);
					varNameStore =  ValueArr[2];
					
					String newXpathStr = xpathStr+"[1]";
					
					try {
					//wait up to 10 seconds from table to show up
					elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(newXpathStr)));
					//elementVar = driver.findElement(By.xpath(newXpathStr));
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+newXpathStr);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the page object identified by: "+newXpathStr);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					boolean found = false;
					//Find first row that has the TextToFind in it 
					List<WebElement> tableRows = driver.findElements(By.xpath(xpathStr));
					for (int j = 1; j == tableRows.size(); j++) {
						newXpathStr = xpathStr+"["+j+"]";
						
						WebElement tableRow = driver.findElement(By.xpath(newXpathStr));
						if (tableRow.getText().contains(searchForStr)) { 
							//System.out.println(tableRow.getText());
							found = true;
							break;
						}
					}
					
					if (found = false) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the text : "+searchForStr+" in the table identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the text : "+searchForStr+" in the table identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
						
					}
					
					//Use column entered to select cell from above selected row 
					String cellXpathStr = newXpathStr + "/div/div["+tableColumn+"]";
					elementVar = driver.findElement(By.xpath(cellXpathStr));
					
					//System.out.println(elementVar.getText());
					//Set variable info
					varValue = elementVar.getText();
					
					//Save variable
					SaveVar(varNameStore,varValue, varArray);

					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' located the text: "+searchForStr+" in the table and generated a variable: " + varNameStore + " with the value of: " + varValue);
					
		           break;
				
				case "FindByRowSelect":
					/*Used to find a specific element on a dynamic table based on a unique value in that table 
					 * and store it in a variable
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Location 								Value 
					 * y	FindByRow	xpath#//div[@class='rt-tr-group']		^BatchId:1											
					 * Var	N			N#Y
					 * 					locator   								TextToFind:occurance
					 * 
					 * ^BatchId:4 is the TextToFind:ColumnToLookUp (in this case ^BatchId is the "Total Submitted" Column and 
					 * ColumnToLookup is Column 4, stores text in someVariable (text in the specified column) 
					 * 
					 */
					wait = new WebDriverWait(driver, 10);
					int k = 0;
					
					//Parse locator and value from Locator
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locatorVal = LocatorArray[0];
					xpathStr = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (xpathStr.substring(0,1).contentEquals("^")) {
						varName = xpathStr.substring(1, locatorVal.length());
						xpathStr = LookupVar(varName,varArray,varEnv);
						Locator = locatorVal + "#" + xpathStr;
					}
					
					//split Value into parts (searchForStr, tableColumn and varNameStore)
					ValueArr = Value.split(",",2); 
					searchForStr = ValueArr[0];
					int intOccurrance = Integer.parseInt(ValueArr[1]);
					
					newXpathStr = xpathStr+"[1]";
					
					try {
						//wait up to 10 seconds from table to show up
						elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(newXpathStr)));
						//elementVar = driver.findElement(By.xpath(newXpathStr));
						} catch (Exception e) {
							WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
									+ Action + "' could not locate the page object identified by: "+newXpathStr);
							WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
									+ Action + "' could not locate the page object identified by: "+newXpathStr);
							failedSteps = failedSteps + 1;
							xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
							break;
						}
					
					found = false;
					//Find first row that has the TextToFind in it 
					tableRows = driver.findElements(By.xpath(xpathStr));
					for (int j = 1; j <= tableRows.size(); j++) {
						newXpathStr = xpathStr+"["+j+"]";
						WebElement tableRow = driver.findElement(By.xpath(newXpathStr));
						if (tableRow.getText().contains(searchForStr)) {
							//found searchForStr now insure correct occurrence
							k = k + 1;
							if (k == intOccurrance) {
								found = true;
								break;
							}
						}
					}
					
					if (found = false) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the "+intOccurrance+" occurance of text : "+searchForStr+" in the table identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the "+intOccurrance+" occurance of text : "+searchForStr+" in the table identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
						
					}
					
					//add to xpath to select checkbox cell from above selected row 
					cellXpathStr = newXpathStr + "/div/div/input";
					elementVar = driver.findElement(By.xpath(cellXpathStr));
					
					elementVar.click();

					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' selected the row that contained the "+intOccurrance+" occurance of the text: "+searchForStr+" in the table and selected the checkbox");

					
		           break;
					
					
				case "CheckRe":
					/*Used for regular expression check to validate information reflected on a web 
					 * element 
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator				Value
					 * y	ReCheck		attribute#value		String to use for Check
					 * y	ReCheck		html id#content		[0-9]{3}0001REJ INPUT COLUMNS WITH X BELOW ARE INVALID - INITIATOR                   SD: 01 DATE [0-9]{5} TIME [0-9]{4}    000000 TR NR [0 ]{5} NGV431
					 * Var	N			N:Y				 	Y
					 */
					//insert  steps

					
					//Parse locator and value from searchName
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locator = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					try {
						elementVar = FindWebElement(driver, locator , locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					String valText = elementVar.getText();
					
					//Replace line separators in valText
					valText = valText.replaceAll("\\r\\n|\\r|\\n", "");
					//System.out.println(valText);
					
					
					//Lookup variable if first position is "^" for inputString
					if (Value.substring(0,1).contentEquals("^")) {
						//strPart is a variable (starts with "^") so lookup variable and return
					
						varName = Value.substring(1, Value.length());
						
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					//Add .* to front and back of regular expression because Java matches to entire string
					Value = ".*" + Value + ".*";
					
					//Regular Expression check
					if (Pattern.matches(Value,valText) == true) {
						System.out.println("PASS---'" + valText + "'");
						System.out.println("Matched as a Regular Expression to:");
						System.out.println("'" + Value + "'");

						WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " matched to the :\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");

						
					}else {
						System.out.println("Fail---'" + valText + "'");
						System.out.println("Did not contain Regular Expression:");
						System.out.println("'" + Value + "'");
						
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " did not match the :\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " did not match the :\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
					}		

					break;
					
				case "CheckNs":
					/*Used to validate information reflected on a web element with all spaces removed
					 * from both elements, checks two ways: in string and exact match
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator(searchName)		Value(inputString)
					 * y	NsCheck		attribute:value			value
					 * y	NsCheck		id#content				Rejects: 1
					 * Var	N			N:Y				 		Y
					 */
					
					//Parse locator and value from searchName
					LocatorArray  = Locator.split("#", 2);
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					locator = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					try {
						elementVar = FindWebElement(driver, locator , locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					valText = elementVar.getText();
					
					//Replace line separators in valText
					valText = valText.replaceAll("\\r\\n|\\r|\\n", "");
					valText = valText.replace(" ", "");
					//System.out.println(valText);
					
					
					//Lookup variable if first position is "^" for inputString
					if (Value.substring(0,1).contentEquals("^")) {
						//strPart is a variable (starts with "^") so lookup variable and return
					
						varName = Value.substring(1, Value.length());
						
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					Value = Value.replace(" ", "");
					
					if (valText == Value) {
						System.out.println("PASS---'"+valText + "' was equal to '" + Value+"'");
						WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " fully matched the test data to the page element using the:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
					}else if (valText.contains(Value) == true) {
						System.out.println("PASS---'"+valText + "' contained the string '" + Value+"'");
						WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " found the test data in the page element using the:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
					}else {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' was not a partial or full match using the string:\r\n\r\n" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' was not a partial or full match using the string:\r\n\r\n" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
					}
					
					break;
					
	
				case "Check":
					/*Used to validate information reflected on a web element, checks two ways:     
					 * in string and exact match
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Location				Value
					 * y	Check		Locator#Locator Value	value
					 * y	Check		html id#content			Reject
					 * Var	N			N:Y				 	Y
					 */
					//System.out.println("Check Steps");
					
					LocatorArray  = Locator.split("#", 2);
					
					
					if (Locator.indexOf("#") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a '#' in it to divide the locator and locator value. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					//Parse locator and value from searchName
					locator  = LocatorArray[0];
					locatorVal = LocatorArray[1];
					
					//Lookup variable if first position is "^" for locatorVal
					if (locatorVal.substring(0,1).contentEquals("^")) {
						varName = locatorVal.substring(1, locatorVal.length());
						locatorVal = LookupVar(varName,varArray,varEnv);
						Locator = locator + "#" + locatorVal;
					}
					
					try {
						elementVar = FindWebElement(driver, locator , locatorVal);
					} catch (Exception e) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
								+ Action + "' could not locate the switch from user identified by: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					valText = elementVar.getText();
					
					//Replace line separators in valText
					valText = valText.replaceAll("\\r\\n|\\r|\\n", "");
					//System.out.println(valText);
					
					
					//Lookup variable if first position is "^" for inputString
					if (Value.substring(0,1).contentEquals("^")) {
						//strPart is a variable (starts with "^") so lookup variable and return
					
						varName = Value.substring(1, Value.length());
						
						Value = LookupVar(varName,varArray,varEnv);
					}
					
					if (valText == Value) {
						System.out.println("PASS---'"+valText + "' was equal to '" + Value+"'");
						WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " fully matched the test data to the page element using the:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
					}else if (valText.contains(Value) == true) {
						System.out.println("PASS---'"+valText + "' contained the string '" + Value+"'");
						WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action "
								+ Action + " found the test data in the page element using the:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
					}else {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' was not a partial or full match using the string:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' was not a partial or full match using the string:\r\n\r\n'" + Value + "'\r\n\r\ncaptured from the " + Locator + " where the actual browser text was : \r\n\r\n'" + valText + "'");
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
					}
					break;
					
				case "MathFunct":
					/*used to add two numbers and output to a variable in a specific format  
					 *     
					 * 			  			  
					 * Spreadsheet format:
					 * Run	Action		Locator					Value
					 * y	MathAdd		7-8,format of output	varName
					 * Var	N			N						N
					 */
					
					LocatorArray  = Locator.split(",", 2);
					
					
					if (Locator.indexOf(",") < 0) {
						WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a ',' in it to seperate the equation from the format. The locator was: "+Locator);
						WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "For the Action '"
								+ Action + "' the Locator must have a ',' in it to seperate the equation from the format. The locator was: "+Locator);
						failedSteps = failedSteps + 1;
						xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
						break;
					}
					
					ScriptEngineManager manager = new ScriptEngineManager();
					ScriptEngine engine = manager.getEngineByName("js");
					Object expResult = engine.eval(LocatorArray[0]);
					
					String mathRes = String.format(LocatorArray[1], expResult);
					
					//System.out.println(mathRes);
					
					
					//Call SaveVar to store data
					SaveVar(Value, mathRes, varArray);
					
					WriteResults(NormalStyle, PassStyle, resultWorkbook, resultSheet, Test_Description, "Pass", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' generated a variable: " + Value + " with the value of: " + mathRes);
					
					break;
				    
				default:
					if (Action == "") {
						Action = "<blank>"; 
					}
					System.out.println("Invalid Action!---" + Action);
					WriteResults(NormalStyle, FailStyle, resultWorkbook, resultSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' is not a recognized keyword");
					WriteResults(NormalStyle, FailStyle, resultWorkbook, failsSheet, Test_Description, "Fail", lineNbr, Action, Locator, Value, "The Action '"
							+ Action + "' is not a recognized keyword");
					failedSteps = failedSteps + 1;
					xlsLineFailed = xlsLineFailed + String.format("%04d", lineNbr)+":";
					break;
					
				
					

				}
				
				
				
				System.out.println(">>>>>>>>>>>>>>>>>>>>> Line number -> " + lineNbr + " is completed");
				System.out.println();

			} else {
				//System.out.println("Line number not executed - " + i);
			}
		}
		
		System.out.println("................................. TEST CASE END .......................................");
		
		String failSteps = String.format("%d", failedSteps);
		
		String[] retArray = {failSteps, xlsLineFailed, xlsLineBug, bugList};
		
		return retArray;
			

	}


		
	public static void ScrollToView(WebDriver driver, WebElement elementVar) {
		//Scroll into view
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("arguments[0].scrollIntoView(true);",elementVar);
	}

	public static String env(String val) {
		if (val.equalsIgnoreCase("FCA")) {
			fn_env = FCA;
		} else if (val.equalsIgnoreCase("FCB")) {
			fn_env = FCB;
		} else if (val.equalsIgnoreCase("FCD")) {
			fn_env = FCD;
		} else if (val.equalsIgnoreCase("FH")) {
			fn_env = FH;
		} else if (val.equalsIgnoreCase("F0B")) {
			fn_env = F0B;
		}
		
		else {
			fn_env = val;
		}
			
		return fn_env;
	}

	 public static final String SELECT_TEXT
     = "(function getSelectionText() {\n"
     + "    var text = \"\";\n"
     + "    if (window.getSelection) {\n"
     + "        text = window.getSelection().toString();\n"
     + "    } else if (document.selection && document.selection.type != \"Control\") {\n"
     + "        text = document.selection.createRange().text;\n"
     + "    }\n"
     //            + "    if (window.getSelection) {\n"
     //            + "      if (window.getSelection().empty) {  // Chrome\n"
     //            + "        window.getSelection().empty();\n"
     //            + "      } else if (window.getSelection().removeAllRanges) {  // Firefox\n"
     //            + "        window.getSelection().removeAllRanges();\n"
     //            + "      }\n"
     //            + "    } else if (document.selection) {  // IE?\n"
     //            + "      document.selection.empty();\n"
     //            + "    }"
     + "    return text;\n"
     + "})()";
	
	
}
