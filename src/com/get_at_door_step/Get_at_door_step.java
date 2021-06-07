package com.get_at_door_step;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Get_at_door_step {

	public static WebDriver d;
	public static Robot robot;

	public void invokebrowser() throws IOException, InterruptedException, AWTException {

		// Provide the Chrome driver location

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\samdany\\Desktop\\chromedriver.exe");
		d = new ChromeDriver();
		d.manage().window().maximize();
		d.manage().deleteAllCookies();

       // Provide the URL to be loaded

		d.get("http://localhost:8080/wordpress/wp-login.php?loggedout=true&wp_lang=en_US");
		d.manage().timeouts().implicitlyWait(10, TimeUnit.MICROSECONDS);

		// Provide the spread sheet location

		String filepath = "C:\\Users\\samdany\\Desktop\\get_at_door_step_test_data.xlsx";
		FileInputStream in = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Starts fetching User name and Password from the spread sheet

		for (int row = 1; row <= sheet.getLastRowNum(); row++) {
			XSSFCell cell = sheet.getRow(row).getCell(0);
			d.findElement(By.id("user_login")).sendKeys(cell.getStringCellValue());
			XSSFCell cell1 = sheet.getRow(row).getCell(1);
			d.findElement(By.xpath("//input[@id='user_pass']")).sendKeys(cell1.getStringCellValue());
		}

		// Clicking on Login button

		d.findElement(By.xpath("//input[@id='wp-submit']")).click();

		// Clicking on Products menu to add Products

		d.findElement(By.xpath("//div[text()='Products'][1]")).click();
		Thread.sleep(2000);

		// Clicking on Products menu to add Products

		d.findElement(By.xpath("//div[text()='Products'][1]")).click();
		Thread.sleep(2000);

		XSSFSheet sheet1 = workbook.getSheetAt(1);
		int rows = sheet1.getLastRowNum();
		int cols = sheet1.getRow(1).getLastCellNum();
		Thread.sleep(2000);

		// Starts fetching the product details from the spread sheet

		// The below 'for' loop is used to iterate the rows given in the spread sheet

		for (int r = 1; r <= rows; r++) {

			// Clicking on Add New button in Products menu

			d.findElement(By.xpath("//div[@class='wrap']//a[text()='Add New']")).click();

			// Fetching each row's data from the spread sheet

			XSSFRow row = sheet1.getRow(r);

			// The below 'for' loop is used to iterate the columns in the above fetched row

			for (int c = 0; c < cols; c++) {

				try {

					// Fetching each column's data from the spread sheet

					XSSFCell cell = row.getCell(c);

					// Below Switch case is used to fetch the data from spread sheet based on the
					// type format of the cells

					switch (cell.getCellType()) {

					case STRING:

						// Below If condition is to check the data available in C0 and enter the same in
						// the Product Name field

						if (c == 0) {

							WebElement name = d.findElement(By.xpath("//input[@id='title']"));
							name.clear();
							name.sendKeys(cell.getStringCellValue());

						}

						// Below If condition is to check the data available in C1 and enter the same in
						// the Product Description field

						if (c == 1) {

							d.switchTo().frame("content_ifr");
							WebElement desc = d.findElement(By.xpath("//body[@id='tinymce'][1]"));
							desc.clear();
							desc.sendKeys(cell.getStringCellValue());
							d.switchTo().parentFrame();

						}

						// Below If condition is to check the data available in C5 and enter the same in
						// the Short Description field

						if (c == 5) {

							d.switchTo().frame("excerpt_ifr");
							WebElement sdesc = d.findElement(By.xpath("//body[@id='tinymce'][1]"));
							sdesc.clear();
							sdesc.sendKeys(String.valueOf(cell.getStringCellValue()));
							d.switchTo().parentFrame();

						}

						if (c == 6) {

							d.findElement(By.xpath("//span[text()='Inventory']")).click();
							WebElement sku = d.findElement(By.xpath("//input[@id='_sku']"));
							sku.clear();
							sku.sendKeys(cell.getStringCellValue());

						}

					case FORMULA:

						// Below If condition is to check the data available in C2 and enter the same in
						// the Regular Price field

						if (c == 2) {

							WebElement rprice = d.findElement(By.xpath("//input[@id='_regular_price']"));
							rprice.clear();
							rprice.sendKeys(String.valueOf(cell.getNumericCellValue()));

						}

						// Below If condition is to check the data available in C3 and enter the same in
						// the Sales Price field

						if (c == 3) {

							WebElement sprice = d.findElement(By.xpath("//input[@id='_sale_price']"));
							sprice.clear();
							sprice.sendKeys(String.valueOf(cell.getNumericCellValue()));

						}

					case NUMERIC:

						// Below If condition is to check the data available in C3 and enter the same in
						// the Sales Price field

						if (c == 3) {

							WebElement sprice = d.findElement(By.xpath("//input[@id='_sale_price']"));
							sprice.clear();
							sprice.sendKeys(String.valueOf(cell.getNumericCellValue()));

						}

					default:
						break;

					}

				} catch (Exception e) {

					e.printStackTrace();

				}

				// Goes back to Line:88 to fetch each column's data of a row

			}

			// Scrolling up because to find the Publish button and clicking the same

			robot = new Robot();
			robot.keyPress(KeyEvent.VK_PAGE_UP);
			robot.keyRelease(KeyEvent.VK_PAGE_UP);
			robot.keyPress(KeyEvent.VK_PAGE_UP);
			robot.keyRelease(KeyEvent.VK_PAGE_UP);
			Thread.sleep(5000);
			d.findElement(By.xpath("//input[@id='publish']")).click();

			/*
			 * After clicking on Publish button, the UI loads for sometime. So added an
			 * explicit wait to validate if the loading is completed and the edit page is
			 * displayed
			 */

			WebDriverWait wait = new WebDriverWait(d, 10);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1[text()='Edit Product']")));

			// Goes back to line:76 to fetch the next row's data and add the next product
		}

		// Quitting the browser

		d.quit();
	}

	// Main method

	public static void main(String[] args) throws IOException, InterruptedException, AWTException {
		Get_at_door_step g = new Get_at_door_step();
		g.invokebrowser();
	}

}
