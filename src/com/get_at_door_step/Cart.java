package com.get_at_door_step;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
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

public class Cart {

	public static WebDriver d;
	public static Robot robot;

	public void invokecart() throws IOException, InterruptedException, AWTException {

		// Provide the Chrome driver location

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\samdany\\Desktop\\chromedriver.exe");
		d = new ChromeDriver();
		d.manage().window().maximize();
		d.manage().deleteAllCookies();

		// Provide the spread sheet location

		String filepath = "C:\\Users\\samdany\\Desktop\\get_at_door_step_test_data.xlsx";
		FileInputStream in = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);

		// Provide the URL to be loaded

		d.get("http://localhost:8080/wordpress/index.php/shop/");

		// Providing the sheet no and no of data available in spread sheet

		XSSFSheet sheet1 = workbook.getSheetAt(1);
		int rowsheet1 = sheet1.getLastRowNum();
		String product_name;
		int r;
		int product_count = 0;

		// Scrolling down because to find the Billing details element and clicking the same

		robot = new Robot();
		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);

		// Starts fetching the product details from the spread sheet

		// The below 'for' loop is used to iterate the rows given in the spread sheet

		for (r = 1; r <= rowsheet1; r++) {

			// Fetching each row's data from the spread sheet

			XSSFRow row = sheet1.getRow(r);
			XSSFCell cellsheet1 = row.getCell(0);

			product_name = cellsheet1.getStringCellValue();
			WebElement add_cart = d.findElement(By.xpath("//a[contains(@aria-label,'" + product_name + "')]"));
			System.out.println(product_name);
			if (add_cart.isDisplayed()) {
				product_count++;
				add_cart.click();
				Thread.sleep(3000);
			} else {
				System.out.println(product_name + " not found");
			}

		}

		// Checking the given product count and selected count is same

		if (rowsheet1 == product_count) {

			WebDriverWait wait = new WebDriverWait(d, 35);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[text()='" + rowsheet1 + "']")));
			WebElement cart_count = d.findElement(By.xpath("//span[text()='" + rowsheet1 + "']"));
			cart_count.click();

		} else {
			System.out.println("nothing added in the cart");
		}

		// Check the subtotal
        
		// Scrolling down because to find the Subtotal element and clicking the same
		
		robot = new Robot();
		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);

		WebElement subtotal = d.findElement(By.xpath("//*[@id=\"post-8\"]/div/div/div/div/div[2]/div/table/tbody/tr[1]/td/span"));

		// Extracting the Subtotal text

		String subtotal_txt = subtotal.getText();

		// Removing the special characters

		subtotal_txt = subtotal_txt.replace("$", "");

		// Parsing the text to Double

		double subtotal_no = Double.parseDouble(subtotal_txt);

		// Checking subtotal is greater than 100

		if (subtotal_no > 100) {
			WebElement checkout = d.findElement(By.xpath("//a[@class='checkout-button button alt wc-forward']"));
			Thread.sleep(3000);
			checkout.click();
		} else {
			System.out.println("To check out the Subtotal order should be greater than 100");
		}

		// Billing Details

		// Providing the sheet no and no of data available in spread sheet

		XSSFSheet sheet2 = workbook.getSheetAt(2);
		int rowsheet2 = sheet2.getLastRowNum();
		int colsheet2 = sheet2.getRow(1).getLastCellNum();
		

		String Fnametxt = null;
		String Lnametxt = null;

		// Starts fetching the Billing details from the spread sheet

		// The below 'for' loop is used to iterate the rows given in the spread sheet

		for (r = 1; r <= rowsheet2; r++) {

			XSSFRow row = sheet2.getRow(r);

			// Fetching each column's data from the spread sheet

			for (int c = 0; c < colsheet2; c++) {

				XSSFCell cellsheet2 = row.getCell(c);
				
				// Below Switch case is used to fetch the data from spread sheet based on the
				// type format of the cells
				
				switch (cellsheet2.getCellType()) {

				case STRING:
					
					// Below If condition is to check the data available in C0 and enter the same in First Name
					
					if (c == 0) 
					{
						WebElement Fname = d.findElement(By.xpath("//input[@id='billing_first_name']"));
						Fnametxt = null;
						Fnametxt = cellsheet2.getStringCellValue();
						Fname.clear();
						Fname.sendKeys(cellsheet2.getStringCellValue());
					}
					
					// Below If condition is to check the data available in C1 and enter the same in Last Name
					
					if (c == 1) 
					{
						WebElement Lname = d.findElement(By.xpath("//input[@id='billing_last_name']"));
						Lnametxt = null;
						Lnametxt = cellsheet2.getStringCellValue();
						Lname.clear();
						Lname.sendKeys(cellsheet2.getStringCellValue());
					}
					
					// Below If condition is to check the data available in C2 and enter the same in Country
					
					if (c == 2) 
					{
						d.findElement(By.xpath("//span[@class='select2-selection__arrow'][1]")).click();
						Thread.sleep(3000);
						WebElement country = d.findElement(By.xpath("//span[@class='select2-results']/ul//li[text()='"+ cellsheet2.getStringCellValue() + "']"));
						country.click();
						System.out.println(cellsheet2.getStringCellValue());
					}
					
					// Below If condition is to check the data available in C3 and enter the same in Country name
					
					if (c == 3) 
					{
						WebElement address = d.findElement(By.xpath("//input[@id='billing_address_1']"));
						address.click();
						address.clear();
						address.sendKeys(cellsheet2.getStringCellValue());
					}

					// Below If condition is to check the data available in C4 and enter the same in Town name
					
					if (c == 4) 
					{
						WebElement town = d.findElement(By.xpath("//input[@id='billing_city']"));
						town.click();
						town.clear();
						town.sendKeys(cellsheet2.getStringCellValue());
					}

					// Below If condition is to check the data available in C5 and enter the same in State name
					
					if (c == 5) 
					{
						d.findElement(By.xpath("//span[@id='select2-billing_state-container']")).click();
						Thread.sleep(3000);
						WebElement country = d.findElement(By.xpath("//span[@class='select2-results']/ul//li[text()='"+ cellsheet2.getStringCellValue().trim() + "']"));
						country.click();
					}
					
					// Below If condition is to check the data available in C5 and enter the same in email
					
					if (c == 7)
					{
						WebElement address = d.findElement(By.xpath("//input[@id='billing_email']"));
						address.click();
						address.clear();
						address.sendKeys(cellsheet2.getStringCellValue());
					}
					
				case NUMERIC:

					// Below If condition is to check the data available in C6 and enter the same in Postcode
					
					if (c == 6) 
					{
						WebElement pincode = d.findElement(By.xpath("//input[@id='billing_postcode']"));
						pincode.click();
						pincode.clear();
						
						//Type casting the decimal number to int
						int no = (int) cellsheet2.getNumericCellValue();
						String value = String.valueOf(no);
						pincode.sendKeys(String.valueOf(value));
					}

				default:
					break;

				}
				// Goes back to Line:154 to fetch each column's data of a row

			}
			
			// Scrolling down because to find the Place order element and clicking the same
			
			robot.keyPress(KeyEvent.VK_PAGE_DOWN);
			robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
			robot.keyPress(KeyEvent.VK_PAGE_DOWN);
			robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
			Thread.sleep(6000);
			
			// Clicking on Place order button
			
			d.findElement(By.xpath("//button[@id='place_order']")).click();
			
		}

		// login to admin to compelete the order

		// Opening the new Tab
		
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		// Switching the control to the new tab
		
		String ParentWindowhandle = d.getWindowHandle();
		
		System.out.println(ParentWindowhandle+ "parent window");
		
		// The below 'for each' loop is used to iterate the no of child tabs
		
		for (String childtab : d.getWindowHandles()) 
		{
		  d.switchTo().window(childtab);
		}
		
		d.get("http://localhost:8080/wordpress/wp-admin/edit.php?post_type=shop_order");
		
		// Provide the spread sheet location

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

		// Scrolling down because to find the button and clicking the same

		robot = new Robot();
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);

		// Getting the order namespace based on order name

		WebDriverWait wait = new WebDriverWait(d, 10);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//strong[contains(text(),'"+Fnametxt+" "+Lnametxt+"')][1]")));
		WebElement ordername =d.findElement(By.xpath("//strong[contains(text(),'"+Fnametxt+" "+Lnametxt+"')][1]"));
		String orderid = ordername.getText();

		// Waits until the wait condition satisfy

		WebDriverWait wait_order_id = new WebDriverWait(d, 10);
		wait_order_id.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//strong[text()='"+orderid+"']")));
		
		// Extracting the numbers alone from string
		
		orderid = orderid.replaceAll("[^\\d]", "");
		
		// Parsing to integer
		int orderidno = Integer.parseInt(orderid);
		
		System.out.println(orderidno+"orderid no");
		
		// Checks the order check box based on id
		
		WebElement order_chk_box = d.findElement(By.xpath("//*[@id='cb-select-"+orderidno+"']"));
		order_chk_box.click(); 
		
		// Selects the Drop down and change the status to completed
		
		WebElement status = d.findElement(By.xpath("//select[@id='bulk-action-selector-top']//option[text()='Change status to completed']"));
		status.click();

		// Clicking on apply button

		WebElement apply = d.findElement(By.xpath("//input[@id='doaction']"));
		apply.click();

		// Quitting the browser

		d.quit();

	}

	public static void main(String[] args) throws IOException, InterruptedException, AWTException {
		Cart c = new Cart();
		c.invokecart();
	}

}
