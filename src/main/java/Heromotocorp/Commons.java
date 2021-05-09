package Heromotocorp;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Commons {

	
	public static FileOutputStream outputStream = null;
	public static FileInputStream inputStream =null;
	public static XSSFWorkbook workbook;
	public static Sheet sheet;
	static Row row;
	public static FileOutputStream outputStream1 = null;
	public static FileInputStream inputStream1 =null;
	public static XSSFWorkbook workbook1;
	public static Sheet sheet1;
    static Row rowtitle;
    static Row rowCounter;
    static Row rowC;
    public static String[] Split_header;
    static Row rowIncrementer;

}
