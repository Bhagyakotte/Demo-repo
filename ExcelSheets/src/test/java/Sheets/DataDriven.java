package Sheets;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {

		String path = "C:\\lab\\ExcelSheets\\testdata\\Book1.xlsx";
		File f = new File(path);
		FileInputStream fs = new FileInputStream(f);
		XSSFWorkbook w = new XSSFWorkbook(fs);
		XSSFSheet s = w.getSheetAt(0);
		
		int rows = s.getLastRowNum();
		for (int i = 0; i < rows; i++) {
			XSSFRow r = s.getRow(i);
			System.out.println("Iteration-reading data from row"+i);
			int cells = r.getLastCellNum();
			for (int c = 0; c < cells; c++) {
				XSSFCell cell = r.getCell(c);
				System.out.println(cell.toString());
			}
		}
	}

}
