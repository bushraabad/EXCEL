package xpath;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.interactions.Actions;
import xpath.WebElements;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class eMumba {
 public static void main (String [] args) throws IOException, InterruptedException{
	 
	 System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
	 WebDriver driver = new ChromeDriver();
	 driver.manage().window().maximize();	 
	        String baseUrl = "http://emumba.com/";
	        String actualTitle = "";
	        driver.get(baseUrl);	     
	       Thread.sleep(3000);
	       WebElements.solutions(driver).click();
	       System.out.println("element when part of the visible text is static syntax3 xpath is successful");
	       WebElements.company(driver).click();
	       System.out.println("element with static visible text syntax2 xpath is successful");
	       WebElements.customers(driver).click();
	       System.out.println("element with static visible text syntax1 xpath is successful");
	       WebElements.services(driver).click();
	       System.out.println("element when part of the visible text is static syntax2 xpath is successful");
	       WebElements.home(driver).click();
	       System.out.println("element when part of the visible text is static syntax1 xpath is successful");
	       WebElements.emumba(driver).click();
	       System.out.println("element with Locating preceding element xpath is successful");
	       WebElements.data(driver).click();
	       System.out.println("element with Locating preceding element xpath is successful");
	       WebElements.form_name(driver).click();
	       
	       System.out.println("element with Locating following sibling xpath  is successful");
	       //WebElements.form_name(driver).sendKeys("Bushra");
	       WebElements.form_email(driver).click();
	       System.out.println("element with Locating following sibling xpath  is successful");
	       //WebElements.form_name(driver).sendKeys("Bushra");
	       //Thread.sleep(3000);	       
	       //driver.close();
 }
}