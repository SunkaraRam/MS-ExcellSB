package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataIntoXL {

	public static void main(String[] args) throws IOException {
			
			FileInputStream fi= new FileInputStream("D:\\ReadDt.xlsx");
			
			Workbook wb = new XSSFWorkbook(fi);
			
			Sheet s1 = wb.getSheet("Love");
			
			Row r1,r2;
			
			Cell c1,c2;
			
			r1 = s1.getRow(1);
			
			c1 = r1.createCell(2);
			
			c1.setCellValue("Pass");
			
			
			r2 = s1.getRow(2);
			
			c2 = r2.createCell(3);
			
			c2.setCellValue("Fail");
			
			FileOutputStream fo = new FileOutputStream("D:\\ReadDt.xlsx");
			
			wb.write(fo);
			
			wb.close();
			
			fi.close();
			
			fo.close();
 		
	}

}
