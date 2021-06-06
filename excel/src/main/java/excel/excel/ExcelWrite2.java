package excel.excel;

import java.io.FileNotFoundException;


import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWrite2 {

	public static void main(String[] args) throws IOException {
		
		int[] serial=new int [5];
		for(int i=0;i<serial.length;i=i++) {
		serial[i]=i+1;	
		}
		
		String f="data/Sample5.xls";
		FileOutputStream fw=new FileOutputStream(f);

		Workbook wb=new HSSFWorkbook();
		Sheet sh=wb.createSheet("Sheet1");
		Row r=sh.createRow(5);
		
		
		for(int i=0;i<serial.length;i=i++) {
			r=sh.createRow(i+1);
		
		}
			Cell c=r.createCell(5);
			c.setCellValue("java");
		
		
		wb.write(fw);
		wb.close();
		fw.close();
		


	}

}
