package ExcelTasks;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProtectedExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(".\\Data\\PassExcel.xlsx");
		String password="root";
		
		XSSFWorkbook workbook =(XSSFWorkbook)WorkbookFactory.create(fis,password);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows=sheet.getLastRowNum();
		System.out.println(rows);
		int cols=sheet.getRow(0).getLastCellNum();
		System.out.println(cols);
		
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
		workbook.close();
		fis.close();
		}
	}

}
