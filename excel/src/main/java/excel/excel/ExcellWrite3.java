package excel.excel;

	
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcellWrite3 {
	
public static void main(String[] args) throws IOException {
		
		
		int[]serial = new int[10];
		for(int i=0; i<serial.length; i=i+1) {
			serial[i]=i+1;
		}
		
		String []name= new String[10];
		name[0]= "student A";
		name[1]= "student B";
		name[2]= "student C";
		name[3]= "student D";
		name[4]= "student E";
		name[5]= "student F";
		name[6]= "student G";
		name[7]= "student H";
		name[8]= "student I";
		name[9]= "student J";
		
		
		String[]result=new String[10];
		
		result[0]="Pass";
		result[1]="Pass";
		result[2]="Pass";
		result[3]="Fail";
		result[4]="Pass";
		result[5]="Pass";
		result[6]="Fail";
		result[7]="Pass";
		result[8]="NImv";
		result[9]="Pass";
		
		
		
		String f = "data/Sample1.xls";
		FileOutputStream fos = new FileOutputStream(f);
		
		// creat a workbook.
		Workbook wb = new HSSFWorkbook();
		// Creating cell style
		
		CellStyle style = wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		
		//creat sheet
		Sheet sh = wb.createSheet("Sheet1");
		// creat a Row 
		Row r = sh.createRow(0);
		
		// creat Cell and set value
		
		Cell c0= r.createCell(0);
		Cell c1 = r.createCell(1);
		Cell c2 = r.createCell(2);
		
		c0.setCellStyle(style);
		c1.setCellStyle(style);
		c2.setCellStyle(style);
		
		c0.setCellValue("Serial no. ");
		c1.setCellValue("Name of the Students ");
		c2.setCellValue("Result ");
		
		// creat Cell and Row for data
		
		for(int i=0; i<serial.length; i=i+1) {
			r=sh.createRow(i+1);
			
			for(int j=0; j<3
; j=j+1) {
				Cell c =r.createCell(j);
				c.setCellStyle(style);
				
				if(c.getColumnIndex()==0) {
					c.setCellValue(serial[i]);
				}
				else if(c.getColumnIndex()==1) {
					c.setCellValue(name[i]);
			}
				else if(c.getColumnIndex()==2) {
					c.setCellValue(result[i]);
			}
		}
		
	}
		
		// Auto resize columns
		for(int i=0; i<5; i=i+1 ) {
			sh.autoSizeColumn(i);
		}
		
		
		
		wb.write(fos);
		wb.close();
		fos.close();
		
	}

}
			
	


