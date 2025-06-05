package Sheets;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDynamicdata {

	public static void main(String[] args) throws IOException {
	    FileOutputStream file=new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\myfile_dynamic.xlsx");
	    XSSFWorkbook w=new XSSFWorkbook();
	    XSSFSheet s=w.createSheet("DynamicData");
	    Scanner sc=new Scanner(System.in);
	    System.out.println("Enter how many rows ");
	    int noOfrows=sc.nextInt();
	    System.out.println("Enter how many cells");
	    int noOfcells=sc.nextInt();
	       for(int r=0;r<=noOfrows;r++) {
	    	   
	       XSSFRow currentRow=s.createRow(r);
	    	    for(int c=0;c<noOfcells;c++) {
	    	    	XSSFCell cell=currentRow.createCell(c);
	    	    	cell.setCellValue(sc.next());
	    	    }
	       }
	    	    	w.write(file);
	    	        w.close();
	    	        file.close();
	    	        System.out.println("file is created...");

	    	   
	       }
	

}
