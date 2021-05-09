package Heromotocorp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class HeromotoOptionsData extends Commons {


	// Runs the header setup for the Excel File for the first time
	@Test( enabled=false)
	public void HeromotocorpHeaders_Setup(){
		HeaderSetup h = new HeaderSetup();
		System.out.println("Headers Configured");

	}


	@Test(enabled=true)
	public void data_Scrapping_Heromotocorp() throws Exception {


		System.setProperty("webdriver.chrome.driver","D:\\Java Workspace\\Heromotocorp\\Driver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS) ;
		driver.get("https://groww.in/options/hero-motocorp-ltd");
		Thread.sleep(5000);
		

		try {
			inputStream = new FileInputStream(new File("D:\\OptionsData\\Heromotocorp.xlsx"));
		} catch (FileNotFoundException e11) {
			// TODO Auto-generated catch block
			e11.printStackTrace();
		}  
		try {
			workbook = new XSSFWorkbook(inputStream);
		} catch (IOException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}
		sheet = workbook.getSheetAt(0);
		
		// ----- Row counter invocation
		try {
			inputStream1 = new FileInputStream(new File("D:\\OptionsData\\RowCounter.xlsx"));
		} catch (FileNotFoundException e11) {
			// TODO Auto-generated catch block
			e11.printStackTrace();
		}  
		try {
			workbook1 = new XSSFWorkbook(inputStream1);
		} catch (IOException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}
		sheet1 = workbook1.getSheetAt(0);
		double CounterFunction= sheet1.getRow(0).getCell(0).getNumericCellValue();
		int rowcounter = (int) Math.round(CounterFunction);
		System.out.println(rowcounter);
		rowC = sheet.createRow(rowcounter+1);

		for(int i=rowcounter;i<=50000;i++) {
			

			String customized_Date = DateAndTime.customized_time();
			System.out.println(customized_Date);
			String spotPrice = driver.findElement(By.xpath("//*[@class='optc56SpotPriceText']")).getText();
			  String sptSubString = spotPrice.substring(13);


			Cell TimeStamp = rowC.createCell(0);
			TimeStamp.setCellValue(customized_Date);
            Cell SpotPrice = rowC.createCell(1);
            SpotPrice.setCellValue(sptSubString);
            
            // Put data collection
            String Pe_2200 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,200.00']/following::td[1]")).getText();
            Pe_2200=Pe_2200.substring(1, 6);
            System.out.println(Pe_2200);
            Cell Put_2200 = rowC.createCell(2);
            Put_2200.setCellValue(Pe_2200);
            String Pe_2250 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,250.00']/following::td[1]")).getText();
            Pe_2250=Pe_2250.substring(1, 6);
            System.out.println(Pe_2250);
            Cell Put_2250 = rowC.createCell(3);
            Put_2250.setCellValue(Pe_2250);
            String Pe_2300 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,300.00']/following::td[1]")).getText();
            Pe_2300=Pe_2300.substring(1, 6);
            System.out.println(Pe_2300);
            Cell Put_2300 = rowC.createCell(4);
            Put_2300.setCellValue(Pe_2300);
            String Pe_2350 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,350.00']/following::td[1]")).getText();
            Pe_2350=Pe_2350.substring(1, 6);
            System.out.println(Pe_2350);
            Cell Put_2350 = rowC.createCell(5);
            Put_2350.setCellValue(Pe_2350);
            String Pe_2400 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,400.00']/following::td[1]")).getText();
            Pe_2400=Pe_2400.substring(1, 6);
            System.out.println(Pe_2400);
            Cell Put_2400 = rowC.createCell(6);
            Put_2400.setCellValue(Pe_2400);
            String Pe_2450 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,450.00']/following::td[1]")).getText();
            Pe_2450=Pe_2450.substring(1, 6);
            System.out.println(Pe_2450);
            Cell Put_2450 = rowC.createCell(7);
            Put_2450.setCellValue(Pe_2450);
            String Pe_2500 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,500.00']/following::td[1]")).getText();
            Pe_2500=Pe_2500.substring(1, 6);
            System.out.println(Pe_2500);
            Cell Put_2500 = rowC.createCell(8);
            Put_2500.setCellValue(Pe_2500);
            String Pe_2550 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,550.00']/following::td[1]")).getText();
            Pe_2550=Pe_2550.substring(1, 6);
            System.out.println(Pe_2550);
            Cell Put_2550 = rowC.createCell(9);
            Put_2550.setCellValue(Pe_2550);
            String Pe_2600 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,600.00']/following::td[1]")).getText();
            Pe_2600=Pe_2600.substring(1, 6);
            System.out.println(Pe_2600);
            Cell Put_2600 = rowC.createCell(10);
            Put_2600.setCellValue(Pe_2600);
            String Pe_2650 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,650.00']/following::td[1]")).getText();
            Pe_2650=Pe_2650.substring(1, 6);
            System.out.println(Pe_2650);
            Cell Put_2650 = rowC.createCell(11);
            Put_2650.setCellValue(Pe_2650);
            String Pe_2700 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,700.00']/following::td[1]")).getText();
            Pe_2700=Pe_2700.substring(1, 6);
            System.out.println(Pe_2700);
            Cell Put_2700 = rowC.createCell(12);
            Put_2700.setCellValue(Pe_2700);
            String Pe_2750 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,750.00']/following::td[1]")).getText();
            Pe_2750=Pe_2750.substring(1, 6);
            System.out.println(Pe_2750);
            Cell Put_2750 = rowC.createCell(13);
            Put_2750.setCellValue(Pe_2750);
            String Pe_2800 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,800.00']/following::td[1]")).getText();
            Pe_2800=Pe_2800.substring(1, 6);
            System.out.println(Pe_2800);
            Cell Put_2800 = rowC.createCell(14);
            Put_2800.setCellValue(Pe_2800);
            String Pe_2850 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,850.00']/following::td[1]")).getText();
            Pe_2850=Pe_2850.substring(1, 6);
            System.out.println(Pe_2850);
            Cell Put_2850 = rowC.createCell(15);
            Put_2850.setCellValue(Pe_2850);
            String Pe_2900 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,900.00']/following::td[1]")).getText();
            Pe_2900=Pe_2900.substring(1, 6);
            System.out.println(Pe_2900);
            Cell Put_2900 = rowC.createCell(15);
            Put_2900.setCellValue(Pe_2900);
            String Pe_2950 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,950.00']/following::td[1]")).getText();
            Pe_2950=Pe_2950.substring(1, 6);
            System.out.println(Pe_2950);
            Cell Put_2950 = rowC.createCell(16);
            Put_2950.setCellValue(Pe_2950);

           
            // Call data collection-------------------------------------------------------------------------------

            
            String Ce_2700 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,700.00']/preceding::td[1]")).getText();
            Ce_2700=Ce_2700.substring(1, 6);
            System.out.println(Ce_2700);
            Cell Call_2700 = rowC.createCell(17);
            Call_2700.setCellValue(Ce_2700);
            
            String Ce_2750 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,750.00']/preceding::td[1]")).getText();
            Ce_2750=Ce_2750.substring(1, 6);
            System.out.println(Ce_2750);
            Cell Call_2750 = rowC.createCell(18);
            Call_2750.setCellValue(Ce_2750);
            
            String Ce_2800 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,800.00']/preceding::td[1]")).getText();
            Ce_2800=Ce_2800.substring(1, 6);
            System.out.println(Ce_2800);
            Cell Call_2800 = rowC.createCell(19);
            Call_2800.setCellValue(Ce_2800);
            
            


			try {
				outputStream = new FileOutputStream("D:\\OptionsData\\Heromotocorp.xlsx");
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			try {
				workbook.write(outputStream);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}







			Thread.sleep(10000);
			int lastRowNum = sheet.getLastRowNum();
			System.out.println(lastRowNum);
			
		
			Thread.sleep(110000);
			rowIncrementer= sheet1.createRow(0);
			  Cell rowval = rowIncrementer.createCell(0);
				rowval.setCellValue(lastRowNum+1);
				outputStream1 = new FileOutputStream("D:\\OptionsData\\RowCounter.xlsx");
				workbook1.write(outputStream1);
			System.out.println("Row incremented Current count is "+lastRowNum);
			int increment_val =lastRowNum+1;
			rowC=sheet.createRow(increment_val);
			driver.navigate().refresh();

		}



					




	}







}