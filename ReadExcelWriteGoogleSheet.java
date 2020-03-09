//This code reads from the excel sheet and uses loops to write on google sheets
package utilities;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.util.Iterator;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadExcelWriteGoogleSheet {
 public static void main (String [] args) throws IOException, InterruptedException{
	 
	 System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
	 WebDriver driver = new ChromeDriver();
	 driver.manage().window().maximize();	 
	        String baseUrl = "https://docs.google.com/spreadsheets/u/0/?tgif=d";
	        String actualTitle = "";
	        driver.get(baseUrl);	     
	       Thread.sleep(3000);
	       Webelement.email(driver).sendKeys("abadbushra@gmail.com");
	       Thread.sleep(1000);
	       Webelement.next(driver).click();
	       Thread.sleep(1000);
	       Webelement.password(driver).sendKeys("purple1122");
	       Thread.sleep(1000);
	       Webelement.next(driver).click();
	       Thread.sleep(1000);
	       Webelement.new_spreadsheet(driver).click();
	       Webelement.SS_title(driver).sendKeys(Keys.CONTROL,"a");
	       Webelement.SS_title(driver).sendKeys(Keys.DELETE);
	       Webelement.SS_title(driver).sendKeys("Testing Excel");
	       Thread.sleep(1000);
	       
	       File src=new File("C:\\Users\\Emumba\\eclipse-workspace\\POIdaraz\\src\\TestingExcel.xlsx");
	       FileInputStream fis=new FileInputStream(src);
	       XSSFWorkbook wb=new XSSFWorkbook(fis);
	       
	       XSSFSheet sheet1=wb.getSheetAt(0);
	       int rowcount=sheet1.getLastRowNum();

	       for(int i=0;i<rowcount+1;i++) {
	           String data0= sheet1.getRow(i).getCell(0).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data0 +"\n");	          
	       }
	       Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           
	       for(int i=0;i<rowcount+1;i++) {
	           String data01= sheet1.getRow(i).getCell(1).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data01 +"\n");
	       }
	       Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data02 = sheet1.getRow(i).getCell(2).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data02 +"\n");
	       }
           Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data03= sheet1.getRow(i).getCell(3).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data03 +"\n");
	       }
           
           Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data04 = sheet1.getRow(i).getCell(4).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data04 +"\n");
	       }
           
           Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data05= sheet1.getRow(i).getCell(5).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data05 +"\n");
	       }
           Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data06= sheet1.getRow(i).getCell(6).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data06 +"\n");
	       }
           Webelement.cell(driver).sendKeys(Keys.RIGHT);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
           
           for(int i=0;i<rowcount+1;i++) {
	           String data07= sheet1.getRow(i).getCell(7).getStringCellValue();
	           Webelement.cell1(driver).sendKeys(data07 +"\n");
	           
	       }
           //Thread.sleep(4000);
           //driver.close();
	       } 	 
 }
  

/*
//this code copies full workbook with all sheets into a new excel file looping sheets and iterating cells
package utilities;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelWriteGoogleSheet {

	public static void main(String[] args) throws IOException 
	{
		
		//Provide the Path of excel file which we want to copy
		File inputFile=new File("C:\\Users\\Emumba\\eclipse-workspace\\POIdaraz\\src\\TestingExcel.xlsx");
		FileInputStream fis=new FileInputStream(inputFile);
		XSSFWorkbook inputWorkbook=new XSSFWorkbook(fis);
		int inputSheetCount=inputWorkbook.getNumberOfSheets();
		System.out.println("Input sheetCount: "+inputSheetCount);
		
		// Provide the path of excel file in which we wanted to copy the data
		File outputFile=new File("C:\\Users\\Emumba\\eclipse-workspace\\POIdaraz\\src\\TestingExcel(1).xlsx");
		FileOutputStream fos=new FileOutputStream(outputFile);
		
		// Creating workbook for output
		XSSFWorkbook outputWorkbook=new XSSFWorkbook();
		
	
		for(int i=0;i<inputSheetCount;i++) 
                { 
                  XSSFSheet inputSheet=inputWorkbook.getSheetAt(i); 
                  String inputSheetName=inputWorkbook.getSheetName(i); 
                  XSSFSheet outputSheet=outputWorkbook.createSheet(inputSheetName); 

                 // Create and call method to copy the sheet and content in new workbook. 
                 copySheet(inputSheet,outputSheet); 
                }

               // Write all the sheets in the new Workbook(testData_Copy.xlsx) using FileOutStream Object
               outputWorkbook.write(fos); 
              // At the end of the Program close the FileOutputStream object. 
               fos.close(); 
          }

           public static void copySheet(XSSFSheet inputSheet,XSSFSheet outputSheet) 
           { 
              int rowCount=inputSheet.getLastRowNum(); 
              System.out.println(rowCount+" rows in inputsheet "+inputSheet.getSheetName()); 
               
                int currentRowIndex=0; if(rowCount>0)
		{
			Iterator rowIterator=inputSheet.iterator();
			while(rowIterator.hasNext())
			{
				int currentCellIndex=0;
				Iterator cellIterator=((Row) rowIterator.next()).cellIterator();
				while(cellIterator.hasNext())
				{
				//  Creating new Row, Cell and Input value in the newly created sheet. 
					String cellData=cellIterator.next().toString();
					if(currentCellIndex==0)
						outputSheet.createRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					else
						outputSheet.getRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					
					currentCellIndex++;
				}
				currentRowIndex++;
			}
			System.out.println((currentRowIndex-1)+" rows added to outputsheet "+outputSheet.getSheetName());
			System.out.println();
		}
	}
}
*/
