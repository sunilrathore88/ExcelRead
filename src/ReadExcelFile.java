import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;

public class ReadExcelFile {

	public static void main(String[] args) {
		
		 try {
		        // Open the Excel file
		        FileInputStream fis = new FileInputStream("D:\\D Drive\\RMS_AutomationScript\\ReadExcel\\src\\Testdatafile.xlsx");
		        // Access the required test data sheet
		        XSSFWorkbook wb = new XSSFWorkbook(fis);
		        XSSFSheet sheet = wb.getSheet("testdata");
		        // Loop through all rows in the sheet
		        // Start at row 1 as row 0 is header row
		      // System.out.println(sheet.getLastRowNum());
		        
		        for(int count = 1;count<=sheet.getLastRowNum();count++){
		            XSSFRow row = sheet.getRow(count);
		            
		            System.out.println("Running test case " + row.getCell(0).toString());
		            // Run the test for the current test data row
		            String result = runTest(row.getCell(1).toString(),row.getCell(2).toString());
		            XSSFCell cell = row.getCell(3);
		            cell.setCellValue(result);
		        }
		        fis.close();
		        
		        FileOutputStream outFile =new FileOutputStream(new File(System.getProperty("user.dir")+"\\Testdatafile.xlsx"));
		        wb.write(outFile);
		        outFile.close();
		    } catch (IOException e) {
		        System.out.println("Test data file not found");
		    }   

	}
	
	

public static String runTest(String strSearchString, String strPageTitle) {
         
        // Start a browser driver and navigate to Google
        //WebDriver driver = new HtmlUnitDriver();
		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
        driver.get("http://www.google.com");
 
        // Enter the search string and send it
        WebElement element = driver.findElement(By.name("q"));
        element.sendKeys(strSearchString);
        element.submit();
         
        // Check the title of the page
        if (driver.getTitle().equals(strPageTitle)) {
            System.out.println("Page title is " + strPageTitle + ", as expected");
           
          //Close the browser
            driver.quit();
            return "Pass";
        } else {
            System.out.println("Expected page title was " + strPageTitle + ", but was " + driver.getTitle() + " instead");
          //Close the browser
            driver.quit();
            return "Fail";
        }
 		
}

}
