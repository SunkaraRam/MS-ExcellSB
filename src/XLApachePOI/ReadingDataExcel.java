package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadingDataExcel {

	public static void main(String[] args) throws IOException
	{
		
		FileInputStream fi = new FileInputStream("D:\\ReadDt.xlsx");
		
		Workbook wb =new XSSFWorkbook(fi);
		
		Sheet s1 = wb.getSheet("Love");

		Row r1=s1.getRow(1);
		
		Cell c1 = r1.getCell(0);
		
		Cell c2 = r1.getCell(1);
		
		String uid,psw;
		
		uid = c1.getStringCellValue();
		
		psw = c2.getStringCellValue();
		
		System.out.println(uid+" " + psw);
		
		/*
		//Cell c1 =  r1.getCell(0);
	
		Cell c2 = r1.getCell(1);
	
		
		int pwd;
	
		//uid = c1.getStringCellValue();
	
		pwd = (int) c2.getNumericCellValue();
	
		System.out.println(pwd);
		
		
	wb.close();
	fi.close();
	
*/
		
		
	}

}
