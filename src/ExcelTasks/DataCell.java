package ExcelTasks;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataCell {

	public static void main(String[] args) throws IOException{
		
	XSSFWorkbook workbook = new XSSFWorkbook();
	
	XSSFSheet sheet =workbook.createSheet("Date Formats");
	//Date Number Format
	XSSFCell cell=sheet.createRow(0).createCell(0);
	cell.setCellValue(new Date());
	
	XSSFCreationHelper creationHelper = workbook.getCreationHelper();
	
	//formats-1 dd-mm-yyyy
	XSSFCellStyle style1 = workbook.createCellStyle();
	style1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy"));
	
	XSSFCell cell1 = sheet.createRow(1).createCell(0);
	cell1.setCellValue(new Date());
	cell1.setCellStyle(style1);
	
	//formats-2 mm-dd-yyyy
		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy"));
		
		XSSFCell cell2 = sheet.createRow(2).createCell(0);
		cell2.setCellValue(new Date());
		cell2.setCellStyle(style2);
	
		//formats-3 mm-dd-yyyy hh:mm:ss
				XSSFCellStyle style3 = workbook.createCellStyle();
				style3.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy hh:mm:ss"));
				
				XSSFCell cell3 = sheet.createRow(3).createCell(0);
				cell3.setCellValue(new Date());
				cell3.setCellStyle(style3);
	
		//formats-4  hh:mm:ss
				XSSFCellStyle style4 = workbook.createCellStyle();
				style4.setDataFormat(creationHelper.createDataFormat().getFormat("hh:mm:ss"));
				
				XSSFCell cell4 = sheet.createRow(4).createCell(0);
				cell4.setCellValue(new Date());
				cell4.setCellStyle(style4);
				
				
	FileOutputStream fos =new FileOutputStream(".\\Data\\Dataformats.xlsx");
	
		workbook.write(fos);
		workbook.close();
		fos.close();
		
		System.out.println("Done...!!!");
	}

}
