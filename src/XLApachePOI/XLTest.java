package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLTest {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi = new FileInputStream("D:\\Sample.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		wb.createSheet("DemoSheet143");
		
		FileOutputStream fo =new FileOutputStream("d://RS11.xlsx");
		
		wb.write(fo);
		
		wb.close();
		
		fi.close();
		 
		fo.close();
	}

}
