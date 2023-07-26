package XL_Demo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class File_Input_Output {

	public static void main(String[] args) throws IOException 
	{
		
		FileInputStream fi = new FileInputStream("D:\\Testing From .XLSX\\Test File_1.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet ws = wb.getSheet("EmpData");// Sheet Name
		Row row = ws.getRow(1);
		Cell cell = row.createCell(4);
		cell.setCellValue("Pass");// Write Data
		
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN1.getIndex());
		style.setFillPattern(FillPatternType.DIAMONDS);
		cell.setCellStyle(style);
		
		FileOutputStream fo = new FileOutputStream("D:\\Testing From .XLSX\\Test File_1.xlsx");
		wb.write(fo);
		
		wb.close();
		
		
		
		
		
		
             		
	}

}
