package Maven_Selenium.Selenium_Java_Project;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcel {
	@Test
	public void readDataFromExcel() throws IOException   {
		FileInputStream file=new FileInputStream("D:\\New XLSX Worksheet.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sheet.getRow(2).getCell(2).getNumericCellValue());
		
		//To write data in excel
		Row row=sheet.createRow(6);
		Cell cell=row.createCell(4);
		cell.setCellValue("Selenium Testing");
		FileOutputStream fos=new FileOutputStream("D:\\New XLSX Worksheet.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("end of writing data in excel");
		
		
		
		
		
	}
	

}
