//This code reads from the excel sheet and uses loops to write on google sheets
package utilities;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;

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
	       Webelement.SS_title(driver).sendKeys(Keys.ENTER);
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
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }       
	       for(int i=0;i<rowcount+1;i++) {
	           String data01= sheet1.getRow(i).getCell(1).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data01 +"\n");
	       }	       
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }         
	       for(int i=0;i<rowcount+1;i++) {
	           String data02= sheet1.getRow(i).getCell(2).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data02 +"\n");
	       }
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }        
	       for(int i=0;i<rowcount+1;i++) {
	           String data03= sheet1.getRow(i).getCell(3).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data03 +"\n");
	       }
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }         
	       for(int i=0;i<rowcount+1;i++) {
	           String data04= sheet1.getRow(i).getCell(4).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data04 +"\n");
	       }
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }         
	       for(int i=0;i<rowcount+1;i++) {
	           String data05= sheet1.getRow(i).getCell(5).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data05 +"\n");
	       }
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }        
	       for(int i=0;i<rowcount+1;i++) {
	           String data06= sheet1.getRow(i).getCell(6).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data06 +"\n");
	       }
	       Webelement.cellB1(driver).sendKeys(Keys.RIGHT);
	       for(int a = 0; a < 60; a++) {
	    	    Webelement.cell(driver).sendKeys(Keys.ARROW_UP);
	       }        
	       for(int i=0;i<rowcount+1;i++) {
	           String data07= sheet1.getRow(i).getCell(7).getStringCellValue();
	           Webelement.cell(driver).sendKeys(data07 +"\n");
	       }
	       Thread.sleep(4000);
           driver.close();
	       } 	
}