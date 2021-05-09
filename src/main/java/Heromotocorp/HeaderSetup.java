package Heromotocorp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HeaderSetup extends Commons{{


	String HeaderTags ="TimeStamp,SpotPrice,Put_2200,Put_2250,Put_2300,Put_2350,Put_2400,Put_2450,Put_2500,Put_2550,Put_2600,Put_2650,Put_2700,Put_2750,Put_2800,Call_2700,Call_2750,Call_2800,Call_2850,Call_2900,Call_2950,Call_3000,Call_3050,Call_3100,Call_3150,Call_3200";
	 Split_header = HeaderTags.split(",");

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
	XSSFCellStyle style = workbook.createCellStyle();
	XSSFFont font=workbook.createFont();
	font.setBold(true);
	style.setFont(font);
	for(int k=0;k<=Split_header.length;k++){
		sheet.setColumnWidth(k,4000);
	}
	rowtitle = sheet.createRow(0);

	for(int val=0;val<Split_header.length;val++)
	{
		Cell validation = rowtitle.createCell(val);
		validation.setCellValue(Split_header[val]);
		
	



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
}

}