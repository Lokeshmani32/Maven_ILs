package com.ils.genericmethods;

// package declared according to java standards
 // next step will be to import drivers for compiling code

import java.io.IOException;
import java.sql.Date;
import java.text.DateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.HSSFBorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
//import org.apache.poi.xssf.usermodel.examples.IterateCells;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

// import option is default availabe at every class or method we use
public class Generic_Methods {
 
	// class being declared
	public static WebDriver driver;

	public static WebElement getelement(String elementname) throws IOException {
		
		// main method with string name elementname is declared for further use

		String[] arr = elementname.split("#");  //   this method is used to split driver.findelement.by to make it as ( driver.findelement(by"locater).sendkeys("value") for excel use.
		WebElement we = null;
		if (arr[0].equalsIgnoreCase("name") == true) {
			we = driver.findElement(By.name(arr[1]));
			// condition by name set

		} else if (arr[0].equalsIgnoreCase("linktext")) {
			// condition 2 for linklist and further go on for six locater

			we = driver.findElement(By.linkText(arr[1]));
		} else if (arr[0].equalsIgnoreCase("xpath")) {

			we = driver.findElement(By.xpath(arr[1]));
		}
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		return we;
	}

	
	public static void click(String elementname) throws IOException { WebElement
	  we = getelement(elementname); we.click();
	 
	}

	public static void openapp(String brname, String url) {
// drivers are being set with conditions..please use google in case of doubt
		if (brname.equalsIgnoreCase("FF") == true) {
			System.setProperty("webdriver.gecko.driver","\\driver\\geckodriver.exe");
			driver = new FirefoxDriver();

		} else if (brname.equalsIgnoreCase("CH") == true) {
			System.setProperty("webdriver.chrome.driver", "\\driver\\ChromeVer78\\chromedriver.exe");
			ChromeOptions options = new ChromeOptions();
			options.addArguments("start-maximized");
			options.setExperimentalOption("useAutomationExtension", false);
			driver = new ChromeDriver();

		} else if (brname.equalsIgnoreCase("IE") == true) {
			System.setProperty("webdriver.ie.driver", "drivers/IEDriverServer.exe");
			driver = new InternetExplorerDriver();
		}

		driver.get(url);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}

	public static WebDriver driver(String brname) {
		String current_path = System.getProperty("user.dir");

		if (brname.equalsIgnoreCase("firefox") == true) {
			driver = new FirefoxDriver();

		} else if (brname.equalsIgnoreCase("Chrome") == true) {
			//System.setProperty("webdriver.chrome.driver", current_path + "\\driver\\chromedriver.exe");
			System.setProperty("webdriver.chrome.driver", current_path + "\\driver\\chromedriver.exe");
			//Disable error
			ChromeOptions options = new ChromeOptions();
		    options.setExperimentalOption("useAutomationExtension", false);
			driver = new ChromeDriver(options);
			
		} else if (brname.equalsIgnoreCase("IE") == true) {
			System.setProperty("webdriver.ie.driver", "drivers/IEDriverServer.exe");
			driver = new InternetExplorerDriver();
		}

		return driver;
	}

	public static void url(String url) {

		driver.get(url);
		// driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		// driver.manage().window().maximize();

	}
	
	public static String[] SplitStr(String splitStr, int nbrSplit) {
		//splits based on the splitStr to a maximum number of splits based on nbrSplit
		String[] splitArray = splitStr.split("#", nbrSplit);
		return splitArray;
	}
	
	public static void WriteResults(CellStyle normalStyle, CellStyle statusStyle, Workbook resultWorkbook, Sheet writeSheet, String Test_Description ,String stepStatus, int testSheetrow, String Action, String Locator, String Value, String Comments) {
		//find last row in writeSheet
		int lastRowNbr = writeSheet.getLastRowNum();
		short col_num;
		
		//Convert testSheetrow to string
		String tstSheetRowStr = String.valueOf(testSheetrow);
		
		//Get timestamp
		DateTimeFormatter formatter =  DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
		String dateStamp = LocalDateTime.now().format(formatter);
		
		//Create an array for data
		String[] stepInfoArr = {Test_Description, stepStatus, tstSheetRowStr, Action, Locator, Value, dateStamp, Comments};
		
		//Add a row
		Row writeRow = writeSheet.createRow(lastRowNbr +1);
		
		//Write Row to sheet
		for (int w = 0; w < stepInfoArr.length; w++) {
			Cell writeCell = writeRow.createCell(w);
			writeCell.setCellValue(stepInfoArr[w]);
						if (w == 1){
				writeCell.setCellStyle(statusStyle);
			} else {
				writeCell.setCellStyle(normalStyle);
			}
		}
	}
	
	public static void SaveVar(String varName, String varValue, String[][] varArray) {
		int m = varArray.length;
		int i = 0;
		int arrayLoc = 0;
		boolean found = false;
		
		/* Get array location for variable varName if a variable with the same name
		 * exists update the value for the variable otherwise create a new entery 
		 * in the varArray to store variable name and value
		 */
		while (i < m) {
			if (varArray[i][0].equals(varName)) {
				arrayLoc = i;
				found = true;
				break;
			} else if (varArray[i][0].isEmpty()) {
				arrayLoc = i;
				break;
			} else {
				i = i + 1;
			}	
		}
		
		if (found == true) {
			//update existing variable value
			varArray[arrayLoc][1] = varValue;
			System.out.println("The array value " + varName + " was changed to " +  varArray[arrayLoc][0] + " : " + varArray[arrayLoc][1]);
		} else {
 			//Add variable name and value to array
        	varArray[arrayLoc][0] = varName;
			varArray[arrayLoc][1] = varValue;
			System.out.println("The array was populated with " +  varArray[arrayLoc][0] + " : " + varArray[arrayLoc][1]);
		
		}
		
	}

	public static String LookupVar(String varName, String[][] varArray, String envVar) {
		String varValue = null;
		//looks up varName in varArray and returns string
		
		String percentChk = varName.substring(0, 2);
		
		//replace %% with first 2 positions of envVar if varName begins with %%
		if (percentChk.contains("%%")) {
			String env = envVar.substring(0, 2);
			varName = varName.substring(2, varName.length());
			varName = env + varName;
		}
	
		int m = varArray.length;
		int i = 0;
		int arrayLoc = 0;
		boolean found;
		
		found = false;
		//Get array location for variable varName
		while (i < m && found == false) {
			if (varArray[i][0].equals(varName)) {
				arrayLoc = i;
				//found empty location in array exit by setting i = m
				found = true;
			} else {
				i = i + 1;
			}	
		}
		//need to code error if varName is not found in varArray
		if (found == true) {
			return varArray[arrayLoc][1];
		} else {
			return "Variable Not Found";
		}
		
	}
	
	public static WebElement FindWebElement(WebDriver driver, String locator, String locatorVal) {
		WebElement elementVar = null;
		WebDriverWait wait = new WebDriverWait(driver, 10);
		
		switch (locator) {
		case "name" :
			//elementVar = driver.findElement(By.name(locatorVal));
			elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorVal)));
		break;
		case "id" :
		case "html id" :	
			//elementVar = driver.findElement(By.id(locatorVal));
			elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorVal)));
		break;
		case "xpath" :
			//elementVar = driver.findElement(By.xpath(locatorVal));
			elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorVal)));
		break;
		case "linkText" :
			//elementVar = driver.findElement(By.linkText(locatorVal));
			elementVar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorVal)));
		break;
		
		}
		
		return elementVar;
	}
}
	

