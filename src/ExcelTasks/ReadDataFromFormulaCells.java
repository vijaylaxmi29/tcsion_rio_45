package ExcelTasks;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromFormulaCells {

	public static void main(String[] args) throws IOException {
		
		FileInputStream Excelfile = new FileInputStream(".\\Data\\ExcelFormula.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(Excelfile);
		XSSFSheet sheet =workbook.getSheet("Sheet1");
		
		//Using for loop
		
				int rows = sheet.getLastRowNum();
				int cols = sheet.getRow(0).getLastCellNum();
				
				for(int i=0;i<=rows;i++) {
					XSSFRow row = sheet.getRow(i);//0
					
					for(int j=0;j<cols;j++) {
						XSSFCell cell = row.getCell(j);
						
						switch(cell.getCellType()){
						case STRING: System.out.print(cell.getStringCellValue());
														break;
						case NUMERIC: System.out.print(cell.getNumericCellValue());
														break;
						case FORMULA: System.out.print(cell.getNumericCellValue());
														break;
						case BOOLEAN: System.out.print(cell.getBooleanCellValue());
														break;
						}
						System.out.print("	|	");
					}
					System.out.println();
				}
				Excelfile.close();  
				

	}

}
