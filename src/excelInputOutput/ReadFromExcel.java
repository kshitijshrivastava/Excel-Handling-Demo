package excelInputOutput;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String value1 = getCellData(2,1);
		System.out.println(value1);
		String value = setCellData(1,1);
		System.out.println(value);
		
		
		
		

	}
	
	public static String getCellData(int rownum, int colnum) throws IOException
	{
		FileInputStream fis = new FileInputStream("C:\\Users\\kshitij.shrivastava\\Desktop\\Selenium Udemy\\excel_data.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh= wb.getSheet("dataSheet");
		XSSFRow row = sh.getRow(rownum);
		XSSFCell cell = row.getCell(colnum);
		String value = cell.getStringCellValue();
		//fis.close();
	    return value;
	}
	
	public static String setCellData(int rownum, int colnum) throws IOException
	{
FileInputStream fis = new FileInputStream("C:\\Users\\kshitij.shrivastava\\Desktop\\Selenium Udemy\\excel_data.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh= wb.getSheet("dataSheet");
		XSSFRow row = sh.getRow(rownum);
		XSSFCell cell = row.getCell(colnum);
		cell.setCellValue("Anisha");
		//fis.close();
		FileOutputStream fos = new FileOutputStream("C:\\Users\\kshitij.shrivastava\\Desktop\\Selenium Udemy\\excel_data.xlsx"); 
		wb.write(fos); 
		fos.close();
		String cellData = cell.getStringCellValue();
		return cellData;
	}

}
