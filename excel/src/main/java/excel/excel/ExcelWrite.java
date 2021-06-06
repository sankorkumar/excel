package excel.excel;

import java.io.FileNotFoundException;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		
		String f="data/Sample.xls";
		FileOutputStream fw=new FileOutputStream(f);

		Workbook wb=new HSSFWorkbook();
		Sheet sh=wb.createSheet("Sheet1");
		Row r=sh.createRow(0);
		Cell c=r.createCell(0);
		c.setCellValue("java");
		
		wb.write(fw);
		wb.close();
		fw.close();
		
	}

}
