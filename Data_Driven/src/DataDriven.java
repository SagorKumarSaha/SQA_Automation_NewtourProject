import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class DataDriven {

	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\USER\\Desktop\\All drivers for selenium\\chromedriver-win64\\chromedriver.exe");
		ChromeDriver driver = new ChromeDriver();
		
		driver.get("https://demo.guru99.com/test/newtours/");
		
		FileInputStream file = new FileInputStream("C:\\Users\\USER\\Desktop\\Registration.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int noOfRows = sheet.getLastRowNum();
		
		System.out.println("Number of rows = " + noOfRows);
		
		for(int row = 1; row<=noOfRows; row++)
		{
			XSSFRow current = sheet.getRow(row);
			String fn = current.getCell(0).getStringCellValue();
			String ln = current.getCell(1).getStringCellValue();
			int pn =(int) current.getCell(2).getNumericCellValue();
			String em = current.getCell(3).getStringCellValue();
			String ad = current.getCell(4).getStringCellValue();
			String ct = current.getCell(5).getStringCellValue();
			String st = current.getCell(6).getStringCellValue();
			int pc = (int)current.getCell(7).getNumericCellValue();
			String cn = current.getCell(8).getStringCellValue();
			String un = current.getCell(9).getStringCellValue();
			String pa = current.getCell(10).getStringCellValue();
			String cp = current.getCell(11).getStringCellValue();
			
			driver.findElement(By.linkText("REGISTER")).click();
			
			driver.findElement(By.name("firstName")).sendKeys(fn);
			driver.findElement(By.name("lastName")).sendKeys(ln);
			driver.findElement(By.name("phone")).sendKeys(String.valueOf(pn));
			driver.findElement(By.name("userName")).sendKeys(em);
			driver.findElement(By.name("address1")).sendKeys(ad);
			driver.findElement(By.name("city")).sendKeys(ct);
			driver.findElement(By.name("state")).sendKeys(st);
			driver.findElement(By.name("postalCode")).sendKeys(String.valueOf(pc));
			Select country = new Select(driver.findElement(By.name("country")));
			country.selectByVisibleText(cn);
			driver.findElement(By.name("email")).sendKeys(un);
			driver.findElement(By.name("password")).sendKeys(pa);
			driver.findElement(By.name("confirmPassword")).sendKeys(cp);
			
			driver.findElement(By.xpath("/html/body/div[2]/table/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/form/table/tbody/tr[17]/td/input")).click();
			
			String ex_title = "Register: Mercury Tours";
			
			if(driver.getTitle().equals(ex_title))
			{
				System.out.println("Registration is succesfull for row number = " + row);
			}
			else {
				System.out.println("Registration is failed for row number = " + row);
				driver.close();
			}
			
			try {
				Thread.sleep(2000);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		System.out.println("Data Driven Test is successfull...");
		driver.quit();

	}

}
