package POIEditTests;

import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import Constants.Constant;

public class Tests {
	static WebDriver driver;

	@BeforeClass
	public static void setup() {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\Administrator\\Downloads\\chromedriver_win32/chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	@AfterClass
	public static void tearDown() {
		driver.quit();
	}

	@Test

	public void test1() {
		driver.navigate().to(Constant.URL1);
		FileInputStream file = null;
		try {
			file = new FileInputStream(Constant.PATH_TESTDATA + Constant.FILE_TESTDATA);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(file);
		} catch (IOException e) {
			e.printStackTrace();
		}
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getRow(0).getLastCellNum();
		for (int i = 1; i <= rowCount + 1; i++) {
			XSSFCell cellUser = sheet.getRow(i).getCell(0);
			XSSFCell cellPassword = sheet.getRow(i).getCell(1);
			XSSFCell cellResult = sheet.getRow(i).createCell(2);
			WebElement addUser = driver.findElement(By
					.xpath("/html/body/div/center/table/tbody/tr[2]/td/div/center/table/tbody/tr/td[2]/p/small/a[3]"));
			addUser.click();
			WebElement userName = driver.findElement(By.xpath(
					"/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/div/center/table/tbody/tr[1]/td[2]/p/input"));
			userName.sendKeys(cellUser.getStringCellValue());
			WebElement passwordFeild = driver.findElement(By.xpath(
					"/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/div/center/table/tbody/tr[2]/td[2]/p/input"));
			passwordFeild.sendKeys(cellPassword.getStringCellValue());
			passwordFeild.submit();

			WebElement login = driver.findElement(By
					.xpath("/html/body/div/center/table/tbody/tr[2]/td/div/center/table/tbody/tr/td[2]/p/small/a[4]"));
			login.click();

			userName = driver.findElement(By.xpath(
					"/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/p/input"));
			userName.sendKeys(cellUser.getStringCellValue());
			passwordFeild = driver.findElement(By.xpath(
					"/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/p/input"));
			passwordFeild.sendKeys(cellPassword.getStringCellValue());
			passwordFeild.submit();
			WebElement successCheck = driver
					.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/big/blockquote/blockquote/font/center/b"));
			System.out.println(successCheck.getText());
			cellResult.setCellValue(successCheck.getText());

		}
		Boolean success = false;
		try {
			FileOutputStream out = new FileOutputStream(new File(Constant.PATH_TESTDATA + Constant.FILE_TESTOUT));
			workbook.write(out);
			out.close();
			success = true;
			System.out.println("File Saved without issue");
		} catch (IOException e) {
			e.printStackTrace();
		}
		assertTrue(success);
	}
}
