package excel.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWrite1 {

	public static void main(String[] args) throws IOException {
		
		String f="data/Sample3.xls";
		FileOutputStream fw=new FileOutputStream(f);
		
		Workbook wb=new HSSFWorkbook();
		Sheet sh=wb.createSheet("sheet1");
		Row r=sh.createRow(0);
		Row r1=sh.createRow(1);
		Row r2=sh.createRow(2);
		Row r3=sh.createRow(3);
		
		Cell c=r.createCell(0);
		Cell c1=r.createCell(1);
		Cell c2=r.createCell(2);
		Cell c3=r.createCell(3);
		
		Cell c01=r1.createCell(0);
		Cell c11=r1.createCell(1);
		Cell c12=r1.createCell(2);
		Cell c13=r1.createCell(3);
		
		Cell c02=r2.createCell(0);
		Cell c21=r2.createCell(1);
		Cell c22=r2.createCell(2);
		Cell c23=r2.createCell(3);
		
		Cell c03=r3.createCell(0);
		Cell c31=r3.createCell(1);
		Cell c32=r3.createCell(2);
		Cell c33=r3.createCell(3);
		
		c01.setCellValue("java01");
		c11.setCellValue("java11");
		c12.setCellValue("java12");
		c13.setCellValue("java13");
		
		c02.setCellValue("java01");
		c21.setCellValue("java11");
		c22.setCellValue("java12");
		c23.setCellValue("java13");
		
		c03.setCellValue("java01");
		c31.setCellValue("java11");
		c32.setCellValue("java12");
		c33.setCellValue("java13");
		
		
		
		c.setCellValue("java");
		c1.setCellValue("java1");
		c2.setCellValue("java2");
		c3.setCellValue("java3");
		
		wb.write(fw);
		wb.close();
		fw.close();

	}

}
