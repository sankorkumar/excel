package excel.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelReadFile {

	public static <Numeric> void main(String[] args) throws IOException {
		
		
		String file="data/Sample3.xls";
		FileInputStream fis=new FileInputStream(file);
		
		Workbook wb=new HSSFWorkbook(fis);
		Sheet sh=wb.getSheet("Sheet1");
		for(Row r:sh) {
			for(Cell c:r) {
				String s=c.getStringCellValue();
				System.out.println(s);
				 
				
			}
		}
		
		wb.close();
		fis.close();
		
		}

}
