package ExcelTasks;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCellbyJava {

	public static void main(String[] args) throws IOException {
		
		String path=".\\Data\\WriteExcelintoFile.xlsx";
		FileInputStream fis = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		sheet.getRow(16).getCell(2).setCellFormula("SUM(C2:C15)");
		
		fis.close();
		
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		
		workbook.close();
		fos.close();
		
		System.out.println("Successfully Done..");

	}

}
