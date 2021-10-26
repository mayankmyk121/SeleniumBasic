import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class FlightAuto {

	public static void main(String[] args) throws InterruptedException, IOException {
		
		// Driver Management
		System.setProperty("webdriver.chrome.driver", System.getProperty("User.dir")+"\\Drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		
		//Waits
		WebDriverWait wait = new WebDriverWait(driver,1000);
		
		//Get WebPage
		driver.get("https://www.easemytrip.com/?msclkid=de917f0e6cc11a6c442e1e643c331372&utm_source=bing&utm_medium=cpc&utm_campaign=Competitors%20RLSA&utm_term=paytm%20flight%20offers%201000%20cashback&utm_content=Competitors%20PM");
		
		//Getting Data
		String[] data = new String[3];
		
		FileInputStream fi = new FileInputStream(System.getProperty("User.dir")+"\\Data\\data.xlsx");
		XSSFWorkbook book = new XSSFWorkbook(fi);
		XSSFSheet sheet = book.getSheet("flight");
		Row row;
		Cell cell;
		
		int i=0;
		Iterator<Row> ritr = sheet.iterator();
		while(ritr.hasNext()) {
			row = ritr.next();
			Iterator<Cell> citr = row.cellIterator();
			while(citr.hasNext()) {
				cell = citr.next();
				DataFormatter format = new DataFormatter();
				data[i] = format.formatCellValue(cell);
				i++;
			}
		}
		
		
		//Source
		driver.findElement(By.xpath("//*[@id=\'FromSector_show\']")).click();
		driver.findElement(By.xpath("//*[@id=\'FromSector_show\']")).sendKeys(data[0]);
		Thread.sleep(500);
		driver.findElement(By.xpath("//*[@id=\'FromSector_show\']")).sendKeys(Keys.ENTER);
		Thread.sleep(500);
		
		//Destination
		driver.findElement(By.xpath("//*[@id=\'Editbox13_show\']")).sendKeys(data[1]);
		Thread.sleep(500);
		driver.findElement(By.xpath("//*[@id=\'Editbox13_show\']")).sendKeys(Keys.ENTER);
		Thread.sleep(500);
		
		//Date
		driver.findElement(By.xpath("//*[@id=\'fiv_1_25/10/2021\']")).click();
		
		//Search
		driver.findElement(By.xpath("/html/body/form/div[5]/div[2]/div[3]/div[1]/div[7]/input")).click();
		
		//Page Load Up
		wait.until(ExpectedConditions.elementToBeClickable(driver.findElement(By.xpath("//*[@id=\'ResultDiv\']/div/div"))));
		
		//Fetching Cheapest Data
		String company = driver.findElement(By.xpath("/html/body/form/div[9]/div[4]/div/div[2]/div[2]/div/div/div[4]/div[1]/div[1]/div[1]/div[1]/div/div[2]/span[1]")).getText();
		String fromTime = driver.findElement(By.xpath("/html/body/form/div[9]/div[4]/div/div[2]/div[2]/div/div/div[4]/div[1]/div[1]/div[1]/div[2]/span[1]")).getText();
		String toTime = driver.findElement(By.xpath("/html/body/form/div[9]/div[4]/div/div[2]/div[2]/div/div/div[4]/div[1]/div[1]/div[1]/div[4]/span[1]")).getText();
		String way = driver.findElement(By.xpath("/html/body/form/div[9]/div[4]/div/div[2]/div[2]/div/div/div[4]/div[1]/div[1]/div[1]/div[3]/span[2]")).getText();
		String rate = driver.findElement(By.xpath("/html/body/form/div[9]/div[4]/div/div[2]/div[2]/div/div/div[4]/div[1]/div[1]/div[1]/div[5]/div[1]/div[2]")).getText();
		System.out.println("Cheapest flight is of "+ company + " timming will be " + fromTime + " to " + toTime + " rate is " + rate + " way is "+ way);

		//Quit
		driver.quit();
	}
}
