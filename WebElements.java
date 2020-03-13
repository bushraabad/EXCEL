package xpath;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class WebElements {

	private static WebElement element = null;
		
		public static WebElement home(WebDriver driver){
			
			element = driver.findElement(By.xpath("//a[contains(text(),'Home')]")); //element when part of the visible text is static syntax1
			return element;
	}
	private static WebElement element1 = null;
		
		public static WebElement services(WebDriver driver){
			
			element = driver.findElement(By.xpath("//a[contains(.,'Services')]")); //element when part of the visible text is static syntax2
			return element;
	}
    private static WebElement element2 = null;
		
		public static WebElement solutions (WebDriver driver){
			
			element = driver.findElement(By.xpath("//a[starts-with(text(),'Sol')]")); //element when prefix of the inner text is static
			//*[contains(text(),'Solutions')]
			return element;
		}
private static WebElement element3 = null;
		
		public static WebElement customers(WebDriver driver){
			
			element = driver.findElement(By.xpath("//a[text()='Customers']"));  //element with static visible text syntax1
			return element;
	}
private static WebElement element4 = null;
		
		public static WebElement company(WebDriver driver){
			
			element = driver.findElement(By.xpath("//*[text()='Company']")); //element with static visible text syntax2
			return element;
	}
private static WebElement element5 = null;
		
		public static WebElement emumba(WebDriver driver){
			
			element = driver.findElement(By.xpath("//ul[@class='navbar-menu hide-nav']/preceding::img")); //element with Locating preceding element 
			return element;
	}
private static WebElement element6 = null;
		
		public static WebElement data(WebDriver driver){
			
			element = driver.findElement(By.xpath("//h3[@class='card-heading']/preceding::figcaption")); //element with Locating preceding element 
			return element;
	}
private static WebElement element7 = null;
		
		public static WebElement form_name(WebDriver driver){
			
			element = driver.findElement(By.xpath("//div[@name='contact-form']/following-sibling::input")); //element with Locating following sibling xpath 
			return element;
	}
private static WebElement element8 = null;
		
		public static WebElement form_email(WebDriver driver){
			
			element = driver.findElement(By.xpath("//input[@class='username']/following-sibling::input")); //element with Locating following sibling xpath 
			return element;
	}
private static WebElement element9 = null;
		
		public static WebElement form_message(WebDriver driver){
			
			element = driver.findElement(By.xpath("//input[@class='username']/following-sibling::*")); //element with Locating following sibling xpath 
			return element;
	}
private static WebElement element10 = null;
		
		public static WebElement footer_logo(WebDriver driver){
			
			element = driver.findElement(By.xpath("//article/figure/picture")); //element with Locating grand children 
			return element;
	}
}