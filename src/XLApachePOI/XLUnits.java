package XLApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

public class XLUnits {
	
	public static FileInputStream fi;
	
	public static Workbook wb;
	
	public static Sheet ws;

	public static Row row;
	
	public static Cell cell;
	
	public static String celldata;
	
	public static CellStyle style;
	
	public static FileOutputStream fo;
	
	
	
	public static int getRowCount(String xlfile, String xlsheet) throws IOException {
		
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(fi);
		
		ws = wb.getSheet(xlsheet);
		
		int rowcount = ws.getLastRowNum();
		
		return rowcount;		
		
	}
	
	public static int getColCount(String xlfile,String xlsheet, int rowcount) throws IOException {
		
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(fi);
		
		ws = wb.getSheet(xlsheet);
		
		row = ws.getRow(rowcount);
		
		int cellcount = row.getLastCellNum();
		
		wb.close();
		
		fi.close();
		
		return cellcount;
	}
	
	public static String getCellData(String xlfile, String xlsheet, int rowcount,int colcount) throws IOException {
	
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(fi);
		
		ws = wb.getSheet(xlsheet);
		
		row = ws.getRow(rowcount);
		
		cell = row.getCell(1);
		
		try {
			
			celldata = cell.getStringCellValue();
			
		} catch (Exception e) {
			
			celldata = "";
		}
		
		wb.close();
		
		fi.close();
		
		return celldata;
		
		}
	
	public static void setCellData(String xlfile,String xlsheet, int rowcount,int colcount,String data) throws IOException {
		
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(fi);
		
		ws = wb.getSheet(xlsheet);
		
		row = ws.getRow(rowcount);
		
		cell = row.createCell(colcount);
		
		cell.setCellValue(data);
		
		fo = new FileOutputStream("d://Rajitha.xlsx");

		wb.write(fo);
		
		wb.close();
		
		fi.close();
		
		fo.close();
		
		
	}
	
	public static void fillGreenColor(String xlfile,String xlsheet, int rowcount,int colcount, String data) throws IOException {
		
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(fi);
		
		ws = wb.getSheet(xlsheet);
		
		row = ws.getRow(rowcount);
		
		cell = row.createCell(colcount);
		
		style = wb.createCellStyle();
		
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		
		fo = new FileOutputStream(xlfile);

	}
	
	public static void fillRedColor(String xlfile,String xlsheet,int rowcount,int colcount,String data) throws IOException {
		
		
		fi = new FileInputStream(xlfile);
		
		wb = new XSSFWorkbook(xlsheet);
		
		ws = wb.getSheet("Love");
		
		row  = ws.getRow(rowcount);
				
		cell  = row.createCell(colcount);
		
		
		
	}
	
	
	
	}
	


