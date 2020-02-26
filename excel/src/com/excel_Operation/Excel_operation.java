package com.excel_Operation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Excel_operation
{

	public static void main(String[] args) throws IOException 
	{
		WebDriver driver = null;
		
		String url = "https://opensource-demo.orangehrmlive.com/";
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\vani\\Desktop\\Automation\\excel\\driverfiles\\chromedriver.exe");
		
		driver = new ChromeDriver();
		
		driver.navigate().to(url);
		
		driver.manage().window().maximize();
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	
	
		FileInputStream file = new FileInputStream("C:\\Users\\vani\\Desktop\\Automation\\excel\\src\\com\\excel\\login.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(file);
		XSSFSheet sheet = workBook.getSheet("Sheet1");

		int rowCount=sheet.getLastRowNum();

		for(int i=1;i<=rowCount;i++)
		{
		// goes to an active Row
		Row row=sheet.getRow(i);

		WebElement username = driver.findElement(By.name("txtUsername"));
		username.clear();
		username.sendKeys(row.getCell(0).getStringCellValue());

		// <input name="txtPassword" id="txtPassword" autocomplete="off" type="password">

		//WebElement password = driver.findElement(By.name(properties.getProperty("password")));
		WebElement password = driver.findElement(By.name("txtPassword"));
		password.clear();
		password.sendKeys(row.getCell(1).getStringCellValue());

		// <input type="submit" name="Submit" class="button" id="btnLogin" value="LOGIN">

		WebElement SignIn = driver.findElement(By.name("Submit"));

		SignIn.click();


		//<a href="#" id="welcome" class="panelTrigger">Welcome Admin</a>

		WebElement Welcome_Admin = driver.findElement(By.linkText("Welcome Admin"));

		String expected_Text="Welcome Admin";
		System.out.println("The expected result is :"+expected_Text);

		String actual_Text = Welcome_Admin.getText();
		System.out.println(" The actual text is :"+actual_Text );

		String expected_HomePageTitle="Find";
		System.out.println("The expected Title of the New Tours Home Page is:"+expected_HomePageTitle);

		String actual_WebPageTitle=driver.getTitle();
		System.out.println(" The actual title of the NewTours WebPage is :"+actual_WebPageTitle );

		if(actual_WebPageTitle.contains(expected_HomePageTitle))
		{
			System.out.println(" LogIN Successfull - PASS");
			row.createCell(2).setCellValue("LogIN Successfull - PASS");
		}
		else
		{
			System.out.println(" LogIn Failed - FAIL");
			row.createCell(2).setCellValue("LogIn Failed - FAIL");
		}

		driver.navigate().back();

		FileOutputStream file1 = new FileOutputStream("C:\\Users\\vani\\Desktop\\Automation\\excel\\src\\com\\excel\\login.xlsx");
		workBook.write(file1);

		}


		}

}
