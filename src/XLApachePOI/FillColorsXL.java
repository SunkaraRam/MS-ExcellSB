package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class FillColorsXL {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi= new FileInputStream("D:\\ReadDt.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet s1 = wb.getSheet("Love");
		
		Row r1 = s1.getRow(1);
		
		Row r2 = s1.getRow(2);
		
		Cell c1 = r1.createCell(3);
		
		Cell c2 = r2.createCell(3);
		
		c2.setCellValue("Fails");
		
		CellStyle style1 =  wb.createCellStyle();
		
		style1.setFillForegroundColor(IndexedColors.RED.getIndex());
		
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		c2.setCellStyle(style1);
		
		c1.setCellValue("Pass");
		
		CellStyle style =  wb.createCellStyle();
		
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		c1.setCellStyle(style);
		
		FileOutputStream fo = new FileOutputStream("D:\\newfile.xlsx");
		
		wb.write(fo);
		
		wb.close();
		
		fi.close();
		
		fo.close();
	}

}
