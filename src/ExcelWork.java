import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class ExcelWork {
      WebDriver driver = new ChromeDriver();
      
      String filePath = "/Users/apple/Downloads/download.xlsx";
      
      @Test()
      public void Test1() {
  		driver.get("https://rahulshettyacademy.com/upload-download-test/");
  		
  		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
  		// download
  		driver.findElement(By.id("downloadButton")).click();
  		
  		//upload
  	    WebElement upload =	driver.findElement(By.xpath("//input[@type='file']"));
  	    upload.sendKeys(filePath);
  	    
  	    // Check Success Message
  	    By toastMessage = By.cssSelector(".Toastify__toast-body div:nth-child(2)");
  	    
  	    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
  	    wait.until(ExpectedConditions.visibilityOfElementLocated(toastMessage));
  	    String successMessage =  driver.findElement(toastMessage).getText();
  	    System.out.print(successMessage);
  	   
  	    //wait.until(ExpectedConditions.invisibilityOfElementLocated(toastMessage));
  	    
  	    driver.close();
  	    
  		}
}
