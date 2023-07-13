package ExcelTasks;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		XSSFRow row = sheet.createRow(0);
		
		row.createCell(0).setCellValue(11);
		row.createCell(1).setCellValue(22);
		row.createCell(2).setCellValue(33);
		
		row.createCell(3).setCellFormula("A1*B1*C1");
		
		FileOutputStream fos = new FileOutputStream(".\\Data\\WriteExcel.xlsx");
		
		workbook.write(fos);
		fos.close();
		
		System.out.println("Sucessfully Done...");

	}

}
