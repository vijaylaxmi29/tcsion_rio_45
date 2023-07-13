package ExcelTasks;

import java.io.*;

import org.apache.poi.xssf.usermodel.*;


public class ReaderExcel {

	public static void main(String[] args) throws IOException {
		
		String path = ".\\Data\\ProjectData.xlsx";
		FileInputStream inputstream = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		//Using for loop
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		
		for(int i=0;i<=rows;i++) {
			XSSFRow row = sheet.getRow(i);//0
			
			for(int j=0;j<cols;j++) {
				XSSFCell cell = row.getCell(j);
				
				switch(cell.getCellType()){
				case STRING: System.out.print(cell.getStringCellValue());
												break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());
												break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
												break;
				}
				System.out.print("	|	");
			}
			System.out.println();
		}

	}

}
