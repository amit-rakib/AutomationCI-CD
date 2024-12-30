import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class ExcelWork {
	WebDriver driver;
    String filePath = "/Users/apple/Downloads/download.xlsx";

    @BeforeMethod
    public void setUp() throws IOException {
    	//Global Properties class
    			Properties prop = new Properties();
    			FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"//src//amit//Global.properties");
    			prop.load(fis);
    			
    			
    			String browser = System.getProperty("browser")!=null? System.getProperty("browser"): prop.getProperty("browser");
    			//String browserName = prop.getProperty("browser");
    			
        if (browser.equalsIgnoreCase("chrome")) {
            driver = new ChromeDriver(); 
        } else if (browser.equalsIgnoreCase("firefox")) {
            driver = new FirefoxDriver(); 
        } else {
            throw new IllegalArgumentException("Browser not supported: " + browser);
        }
    }

      
	  
      @Test
      public void Test1() throws IOException, InterruptedException {
  		driver.get("https://rahulshettyacademy.com/upload-download-test/");
  		
  		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
  		// download
  		driver.findElement(By.id("downloadButton")).click();
  		
  		Thread.sleep(3000);
  		
  		// Edit excel 
  	    int row =	getRowNumber(filePath, "price");
        int col =		getColumnNumber(filePath, "Apple");
        
        updateCell(row, col, "500");
  		
  		
  		
  		FileInputStream fis = new FileInputStream(filePath);
  		
  		XSSFWorkbook workbook = new XSSFWorkbook(fis);
  	    XSSFSheet sheet = workbook.getSheet("Sheet1");
  	    
  		
  		//upload
  	    WebElement upload =	driver.findElement(By.xpath("//input[@type='file']"));
  	    upload.sendKeys(filePath);
  	    
  	    // Check Success Message
  	    By toastMessage = By.cssSelector(".Toastify__toast-body div:nth-child(2)");
  	    
  	    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
  	    wait.until(ExpectedConditions.visibilityOfElementLocated(toastMessage));
  	    String successMessage =  driver.findElement(toastMessage).getText();
  	    System.out.print(successMessage);
  	   
  	    wait.until(ExpectedConditions.invisibilityOfElementLocated(toastMessage));
  	    
  	    
  	  String updatedPrice =  driver.findElement(By.xpath("//div[text()='Apple']/parent::div/parent::div/div[@id='cell-4-undefined']")).getText();
  	    Assert.assertEquals("500", updatedPrice);
  	    
  		}
      
      @AfterMethod
      public void tearDown() {
    	  driver.close();
      }

	private void updateCell(int row, int col, String string) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	    XSSFSheet sheet = workbook.getSheet("Sheet1");
	  XSSFRow rowField =  sheet.getRow(row-1);
	XSSFCell cellField =  rowField.getCell(col+1);
	cellField.setCellValue(string);
	
	FileOutputStream fos = new FileOutputStream(filePath);
	workbook.write(fos);
	workbook.close();
	
		
	}

	private int getColumnNumber(String filePath2, String string) throws IOException {
		FileInputStream fis = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	    XSSFSheet sheet = workbook.getSheet("Sheet1");
	    Iterator<Row> rows =  sheet.iterator();
	    
	    int k=0;
	    int rowIndex = 0;
	    while(rows.hasNext()) {
	          Row row =rows.next();
	     Iterator<Cell> cells =  row.cellIterator();
	     while(cells.hasNext()) {
	    	Cell cell = cells.next();
	    	if(cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(string)) {
	    		rowIndex = k;
	    	}
	    
	     }
	     k++;
	    }
	   
	    System.out.println("column ="+rowIndex);
		
		return rowIndex;
		
		
	}

	private int getRowNumber(String filePath2, String price) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	    XSSFSheet sheet = workbook.getSheet("Sheet1");
	    Iterator<Row> rows =  sheet.iterator();
	    
	    int k=0;
	    int rowIndex = 0;
	    while(rows.hasNext()) {
	          Row row =rows.next();
	     Iterator<Cell> cells =  row.cellIterator();
	     while(cells.hasNext()) {
	    	Cell cell = cells.next();
	    	if(cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(price)) {
	    		rowIndex = k;
	    	}
	    	k++;
	     }
	     
	    }
	   
	    System.out.println("row ="+rowIndex);
		
		return rowIndex;
		
	}


}
