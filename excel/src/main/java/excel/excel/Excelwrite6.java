package excel.excel;

import java.io.FileNotFoundException;

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


public class Excelwrite6 {

	public static void main(String[] args) throws IOException {
		
		int[]serial=new int[6];
		for(int i=0;i<serial.length;i=i+1) {
			serial[i]=i+1;
		}
			String []name= new String[6];
			name[0]= "santosh";
			name[1]= "sankor";
			name[2]= "maitree";
			name[3]= "aongkita";
			name[4]= "anado";
			name[5]= "putu";
			
			String[]deases=new String[6];
			deases[0]="ashma";
			deases[1]="acid";
			deases[2]="backpain";
			deases[3]="skin whiteness";
			deases[4]="ashma";
			deases[5]="no deases";
			
			String[]dr=new String[6];
			dr[0]="gorge 2nd";
			dr[1]="gorge. k";
			dr[2]="gorge. b";
			dr[3]="gorge .m";
			dr[4]="gorge .ll";
			dr[5]="gorge ben";
			
			
			String s="data/Sample8.xls";//file location.
			FileOutputStream fos=new FileOutputStream(s);
			//create work book.
			Workbook wb=new HSSFWorkbook();
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
			Cell c3= r.createCell(3);
			
			c0.setCellStyle(style);
			c1.setCellStyle(style);
			c2.setCellStyle(style);
			c3.setCellStyle(style);
			
			c0.setCellValue("Serial no. ");
			c1.setCellValue("Name of patient ");
			c2.setCellValue("deases");
			c3.setCellValue("dr.name");
			
			// creat Cell and Row for data
			
			for(int t=0; t<serial.length; t=+1) {
				r=sh.createRow(t+1);
			}
				for(int j=0; j<5; j=j+1) {
					Cell c =r.createCell(j);
					c.setCellStyle(style);
				
					if(c.getColumnIndex()==0) {
						c.setCellValue(serial[j]);
					}
					else if(c.getColumnIndex()==1) {
						c.setCellValue(name[j]);
				}
					else if(c.getColumnIndex()==2) {
						c.setCellValue(deases[j]);
					}
					else if(c.getColumnIndex()==3) {
						c.setCellValue(dr[j]);
					}
			

				}
		
	
		// Auto resize columns
				for(int i=0; i<6; i=i+1 ) {
					sh.autoSizeColumn(i);
				}
				
				
				
			
				wb.write(fos);
				wb.close();
				fos.close();
				
			}

		}

		 

			
			
			
			
			
			
			
			
			
			
			

		

	


