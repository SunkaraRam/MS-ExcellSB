package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NullExceReadingData {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi =new FileInputStream("D:\\ReadDt.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet s1 = wb.getSheet("Love");
		
		Row r1 = s1.getRow(1);
				
		Cell c1;
		
		String cvalue;
		try {
			

			c1 = r1.getCell(0);
			
			cvalue = c1.getStringCellValue();
			
			
		} catch (Exception e) {
			
			cvalue = " ";
		}
		
		
	
	}

}
