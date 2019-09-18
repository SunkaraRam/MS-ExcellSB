package XLApachePOI;

import java.io.IOException;

public class DemoTest {

	public static void main(String[] args) throws IOException {
		
		int rc = XLUnits.getRowCount("D:\\Sample.xlsx", "LoginData" );
		
		System.out.println("Rows In LoginData Sheet: " + rc);

		int colcount = XLUnits.getColCount("D:\\Sample.xlsx","LoginData", 2);
		
		System.out.println("Colls in Row : " + colcount);
		
		String data = XLUnits.getCellData("D:\\Sample.xlsx","LoginData", 2 , 1);
		
		System.out.println(data);
		
		XLUnits.setCellData("D:\\Sample.xlsx","LoginData", 2 , 8 ,"Jahn");
		
		XLUnits.fillGreenColor("D:\\Sample.xlsx","LoginData", 2, 8, "Pass");
		
		
	}

}
