package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sun.rowset.internal.Row;

public class CountRows {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi =new FileInputStream("d://Sample.xlsx");
		
		Workbook wb =new XSSFWorkbook(fi);
		
		Sheet s1= wb.getSheet("LoginData");
		
		Sheet s2 = wb.getSheet("EmpData");
		
		int rc = s1.getLastRowNum();
		
		System.out.println("Rows in Login Data : " +rc);
		
		rc =s2.getLastRowNum();
		
		System.out.println("Rows in EmpData : " +rc);
		    
		wb.close();
		
		fi.close();		

	}

}
