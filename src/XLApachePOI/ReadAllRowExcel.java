package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAllRowExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi = new FileInputStream("D:\\ReadDt.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet s1 = wb.getSheet("Love");	
		
		int rc = s1.getLastRowNum();
		
		Row r;
		
		Cell c1,c2;
		
		String uid,pwd;
		
		for (int i = 1; i <= rc; i++) {
			
			r = s1.getRow(i);
			
			c1 = r.getCell(0);
			
			c2 = r.getCell(1);
			
			uid = c1.getStringCellValue();
			
			pwd = c2.getStringCellValue();
			
			System.out.println(uid+"  "+pwd);
			
			
		}
		/*
		//For Sheet Name Printing
		String s1 = wb.getSheetName(0);
		
		System.out.println(s1);
		*/
	}

}
