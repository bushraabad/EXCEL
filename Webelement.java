package utilities;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class Webelement {
private static WebElement element = null;
	
	public static WebElement email(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div[1]/div/div[1]/div/div[1]/input"));
		return element;
	

}
private static WebElement element1 = null;
	
	public static WebElement next(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/span/span"));
		return element;
	
}
private static WebElement element2 = null;
	
	public static WebElement password(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div[1]/div[1]/div/div/div/div/div[1]/div/div[1]/input"));
		return element;
	
}
private static WebElement element3 = null;
	
	public static WebElement next1(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/span/span"));
		return element;
	
}
private static WebElement element4 = null;
	
	public static WebElement new_spreadsheet(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[2]/div[3]/div[1]/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[2]/div[1]/div[1]/img"));
		return element;
	
}
private static WebElement element5 = null;
	
	public static WebElement SS_title(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[2]/div[1]/div[2]/div[1]/div/div[1]/input"));
		return element;
	}
private static WebElement element6 = null;
	
	public static WebElement cell(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[9]/div/div"));
		return element;
	}
private static WebElement element7 = null;
	
	public static WebElement cell1(WebDriver driver){
		
		element = driver.findElement(By.xpath("/html/body/div[9]/div/div"));
		return element;
	}
private static WebElement element8 = null;
	
	public static WebElement file(WebDriver driver){
		
		element = driver.findElement(By.id("docs-file-menu"));
		return element;
	}
}
