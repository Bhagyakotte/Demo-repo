package Sheets;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataIntoExcel {

	public static void main(String[] args) throws IOException {
    FileOutputStream file=new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\myfile.xlsx");
    XSSFWorkbook w=new XSSFWorkbook();
    XSSFSheet s=w.createSheet("Data");
    
    XSSFRow row1=s.createRow(0);
    row1.createCell(0).setCellValue("java");
    row1.createCell(1).setCellValue(18);
    row1.createCell(2).setCellValue("Automation");
    
    XSSFRow row2=s.createRow(1);
    row2.createCell(0).setCellValue("Python");
    row2.createCell(1).setCellValue(3);
    row2.createCell(2).setCellValue("Automation");
    
    XSSFRow row3=s.createRow(2);
    row3.createCell(0).setCellValue("C#");
    row3.createCell(1).setCellValue(5);
    row3.createCell(2).setCellValue("Automation");

    w.write(file);
    w.close();
    file.close();
    System.out.println("file is created...");

	
	}

}
