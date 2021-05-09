package Heromotocorp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Test extends Commons {

	public static void main(String[] args) {

System.setProperty("webdriver.chrome.driver","D:\\Java Workspace\\Heromotocorp\\Driver\\chromedriver.exe");
WebDriver driver = new ChromeDriver();
driver.manage().window().maximize();
driver.manage().deleteAllCookies();
driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS) ;
driver.get("https://groww.in/options/hero-motocorp-ltd");
		String spotPrice = driver.findElement(By.xpath("//*[@class='optc56SpotPriceText']")).getText();
		  String sptSubString = spotPrice.substring(13);
		  System.out.println(sptSubString);
	}
}

