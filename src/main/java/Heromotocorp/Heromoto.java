package Heromotocorp;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Heromoto {
public static FileOutputStream outputStream;
public static FileOutputStream outputStream1;
public static FileInputStream inputStream;
public static FileInputStream inputStream1;
public static XSSFWorkbook workbook;
static XSSFWorkbook workbook1;
public static Sheet sheet;
static Sheet sheet1;
static Row row;
static Row rowCounter;







Heromoto(){
try {
inputStream = new FileInputStream(new File("D:\\OptionsData\\Heromotocorp.xlsx"));
} catch (FileNotFoundException e1) {
// TODO Auto-generated catch block
e1.printStackTrace();
}  
try {
workbook = new XSSFWorkbook(inputStream);
} catch (IOException e1) {
// TODO Auto-generated catch block
e1.printStackTrace();
}
sheet = workbook.getSheetAt(0);
XSSFCellStyle style = workbook.createCellStyle();
XSSFFont font=workbook.createFont();
font.setBold(true);
style.setFont(font);
Row rowtitle = sheet.createRow(0);
Cell Total_bal = rowtitle.createCell(0);
Total_bal.setCellValue("TimeStamp");
Cell SpotPrice = rowtitle.createCell(1);
SpotPrice.setCellValue("SpotPrice");
Cell Pay_In = rowtitle.createCell(2);
Pay_In.setCellValue("Pe_2200");
Cell Pay_Inq = rowtitle.createCell(3);
Pay_Inq.setCellValue("Pe_2250");
Cell Pay_Inn = rowtitle.createCell(4);
Pay_Inn.setCellValue("Pe_2300");
Cell Pay_Ins = rowtitle.createCell(5);
Pay_Ins.setCellValue("Pe_2350");
Cell Wallet_balance = rowtitle.createCell(6);
Wallet_balance.setCellValue("Pe_2400");
Cell Contest_amt = rowtitle.createCell(7);
Contest_amt.setCellValue("Pe_2450");
Cell Winning_amt = rowtitle.createCell(8);
Winning_amt.setCellValue("Pe_2500");
Cell Credit_Days = rowtitle.createCell(9);
Credit_Days.setCellValue("Pe_2550");
Cell Credit_Dayss = rowtitle.createCell(10);
Credit_Dayss.setCellValue("Pe_2600");
Cell Credit_Daysss = rowtitle.createCell(11);
Credit_Daysss.setCellValue("Pe_2650");
Cell Cre_2700 = rowtitle.createCell(12);
Cre_2700.setCellValue("Pe_2700");
Cell Cre_2750 = rowtitle.createCell(13);
Cre_2750.setCellValue("Pe_2750");
Cell Cre_2800 = rowtitle.createCell(14);
Cre_2800.setCellValue("Pe_2800");
Cell Cre_2850 = rowtitle.createCell(15);
Cre_2850.setCellValue("Pe_2850");
Cell Cre_2900 = rowtitle.createCell(16);
Cre_2900.setCellValue("Pe_2900");
Cell Cre_2950 = rowtitle.createCell(17);
Cre_2950.setCellValue("Pe_2950");
Cell Cre_3000 = rowtitle.createCell(18);
Cre_3000.setCellValue("Pe_3000");
Cell Cre_3050 = rowtitle.createCell(19);
Cre_3050.setCellValue("Pe_3050");
Cell Cre_3100 = rowtitle.createCell(20);
Cre_3100.setCellValue("Pe_3100");

//    Calls

Cell Cal_2700 = rowtitle.createCell(21);
Cal_2700.setCellValue("Ce_2700");
Cell Cal_2750 = rowtitle.createCell(22);
Cal_2750.setCellValue("Ce_2750");
Cell Cal_2800 = rowtitle.createCell(23);
Cal_2800.setCellValue("Ce_2800");
Cell Cal_2850 = rowtitle.createCell(24);
Cal_2850.setCellValue("Ce_2850");
Cell Cal_2900 = rowtitle.createCell(25);
Cal_2900.setCellValue("Ce_2900");
Cell Cal_2950 = rowtitle.createCell(26);
Cal_2950.setCellValue("Ce_2950");
Cell Cal_3000 = rowtitle.createCell(27);
Cal_3000.setCellValue("Ce_3000");
Cell Cal_3050 = rowtitle.createCell(28);
Cal_3050.setCellValue("Ce_3050");
Cell Cal_3100 = rowtitle.createCell(29);
Cal_3100.setCellValue("Ce_3100");
Cell Cal_3150 = rowtitle.createCell(30);
Cal_3150.setCellValue("Ce_3150");
Cell Cal_3200 = rowtitle.createCell(31);
Cal_3200.setCellValue("Ce_3200");


for(int k=0;k<=31;k++){
sheet.setColumnWidth(k,4000);
}
for(int j = 0; j<=31; j++)
rowtitle.getCell(j).setCellStyle(style);

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






}


public static void main(String args[]) throws InterruptedException{
Heromoto heromoto = new Heromoto();

System.setProperty("webdriver.chrome.driver","D:\\Java Workspace\\Heromotocorp\\Driver\\chromedriver.exe");
WebDriver driver = new ChromeDriver();
driver.manage().window().maximize();
driver.manage().deleteAllCookies();
driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS) ;
driver.get("https://groww.in/options/hero-motocorp-ltd");
Thread.sleep(5000);
for(int i=0;i<50;i++){
String spotPrice = driver.findElement(By.xpath("//*[@class='optc56SpotPriceText']")).getText();
System.out.println(spotPrice);
driver.findElement(By.xpath("//*[@class='pos-rel valign-wrapper se55SelectBox clrText optc56ActiveOptBox']")).click();
/* driver.findElement(By.xpath("(//*[@class='se55DropdownPara '])[1]")).click();*/
Thread.sleep(10000);
String Pe_2200 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,200.00']/following::td[1]")).getText();
Pe_2200=Pe_2200.substring(1, 6);
System.out.println(Pe_2200);
String Pe_2250 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,250.00']/following::td[1]")).getText();
Pe_2250=Pe_2250.substring(1, 6);
System.out.println(Pe_2250);
String Pe_2300 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,300.00']/following::td[1]")).getText();
Pe_2300=Pe_2300.substring(1, 6);
System.out.println(Pe_2300);
String Pe_2350 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,350.00']/following::td[1]")).getText();
Pe_2350=Pe_2350.substring(1, 6);
System.out.println(Pe_2350);
String Pe_2400 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,400.00']/following::td[1]")).getText();
Pe_2400=Pe_2400.substring(1, 6);
System.out.println(Pe_2400);
String Pe_2450 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,450.00']/following::td[1]")).getText();
Pe_2450=Pe_2450.substring(1, 6);
System.out.println(Pe_2450);
String Pe_2500 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,500.00']/following::td[1]")).getText();
Pe_2500=Pe_2500.substring(1, 6);
System.out.println(Pe_2500);
String Pe_2550 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,550.00']/following::td[1]")).getText();
Pe_2550=Pe_2550.substring(1, 6);
System.out.println(Pe_2550);
String Pe_2600 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,600.00']/following::td[1]")).getText();
Pe_2600=Pe_2600.substring(1, 6);
System.out.println(Pe_2600);
String Pe_2650 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,650.00']/following::td[1]")).getText();
Pe_2650=Pe_2650.substring(1, 6);
System.out.println(Pe_2650);
String Pe_2700 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,700.00']/following::td[1]")).getText();
Pe_2700=Pe_2700.substring(1, 6);
System.out.println(Pe_2700);
String Pe_2750 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,750.00']/following::td[1]")).getText();
Pe_2750=Pe_2750.substring(1, 6);
System.out.println(Pe_2750);
String Pe_2800 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,800.00']/following::td[1]")).getText();
Pe_2800=Pe_2800.substring(1, 6);
System.out.println(Pe_2800);



//------------------------------Calls Configuration----------------------------------------------//  



String Ce_2700 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,700.00']/preceding::td[1]")).getText();
Ce_2700=Ce_2700.substring(1, 6);
System.out.println(Ce_2700);
String Ce_2750 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,750.00']/preceding::td[1]")).getText();
Ce_2750=Ce_2750.substring(1, 6);
System.out.println(Ce_2750);
String Ce_2800 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,800.00']/preceding::td[1]")).getText();
Ce_2800=Ce_2800.substring(1, 6);
System.out.println(Ce_2800);
String Ce_2850 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,850.00']/preceding::td[1]")).getText();
Ce_2850=Ce_2850.substring(1, 6);
System.out.println(Ce_2850);
String Ce_2900 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,900.00']/preceding::td[1]")).getText();
Ce_2900=Ce_2900.substring(1, 6);
System.out.println(Ce_2900);
String Ce_2950 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='2,950.00']/preceding::td[1]")).getText();
Ce_2950=Ce_2950.substring(1, 6);
System.out.println(Ce_2950);
String Ce_3000 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='3,000.00']/preceding::td[1]")).getText();
Ce_3000=Ce_3000.substring(1, 6);
System.out.println(Ce_3000);
String Ce_3050 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='3,050.00']/preceding::td[1]")).getText();
Ce_3050=Ce_3050.substring(1, 6);
System.out.println(Ce_3050);
String Ce_3100 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='3,100.00']/preceding::td[1]")).getText();
Ce_3100=Ce_3100.substring(1, 6);
System.out.println(Ce_3100);
String Ce_3150 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='3,150.00']/preceding::td[1]")).getText();
Ce_3150=Ce_3150.substring(1, 6);
System.out.println(Ce_3150);
String Ce_3200 = driver.findElement(By.xpath("//*[@class='opr84StrikeCell' and text()='3,200.00']/preceding::td[1]")).getText();
Ce_3200=Ce_3200.substring(1, 6);
System.out.println(Ce_3200);

Thread.sleep(50000);
driver.navigate().refresh();

}

}









}