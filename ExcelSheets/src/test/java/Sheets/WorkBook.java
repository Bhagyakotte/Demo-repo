package Sheets;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkBook {
	// htyvbijn
	//98hijoml
	//9nimpl

	private static String path;

	public static void main(String[] args) throws IOException  {
	FileInputStream file=new FileInputStream(System.getProperty("user.dir")+"\\properties\\example.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheet("sheet1");
		int totalRows=sheet.getLastRowNum();
		int totalCells=sheet.getRow(1).getLastCellNum();
		System.out.println("number of rows:"+totalRows);
		System.out.println("number of cells:"+totalCells);
		//changes
		for(int r=0;r<=totalRows;r++)
		{
			XSSFRow currentRow=sheet.getRow(r);
			for(int c=0;c<totalCells;c++) {
		XSSFCell cell=currentRow.getCell(c);	
		if (cell != null) {
		    System.out.print(cell.toString());
		} 
		System.out.println();
		}
		
	workbook.close();
	file.close();
	}

}
}
