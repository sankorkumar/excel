package excel.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelnew {

	public static void main(String[] args) throws IOException {
		
		String s="data/Sample.xlsx";
		FileOutputStream fw=new FileOutputStream(s);
		
		Workbook wb=new XSSFWorkbook();
		  Sheet  sh =wb.createSheet("Sheet1");
		 Row r =sh.createRow(0);
		 Cell c=r.createCell(0);
		 Cell c1=r.createCell(1);
		 
		 c.setCellValue("java");
		 c1.setCellValue("fun");
		 wb.write(fw);
		 fw.close();
		 
		 
		  

	}

}
