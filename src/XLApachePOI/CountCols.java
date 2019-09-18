package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class CountCols {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi = new FileInputStream("d://Sample.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet s1 = wb.getSheet("LoginData");
		
		Row r1= s1.getRow(0);
		
		Row r2 = s1.getRow(1);
		
		int countcol = r1.getLastCellNum();
		
		System.out.println("Columns in First Row : " + countcol);
		
		countcol = r2.getLastCellNum();
		
		System.out.println("Columns in Second Row : " + countcol);
		
		
	}

}                              
	