package ExcelTasks;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel1 {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student Data");
		
		Object Stud[][] = { {"ID","Name","Batch"},
							{1,"Aman","BSC"},
							{2,"Bhushan","BA"},
							{3,"Chaman","BCA"},
							{4,"Dany","BBA"},
							{5,"Ekta","BCOM"}
						};
		
		// Using For Loop
		
		int rows = Stud.length;
		int cols = Stud[0].length;
		
		System.out.println(rows);
		System.out.println(cols);
		
		for(int i=0;i<rows;i++) {
			
			XSSFRow row = sheet.createRow(i);
			
			for(int j=0;j<cols;j++) {
				
				XSSFCell cell = row.createCell(j);
				Object value = Stud[i][j];
				
				if(value instanceof String) {
					cell.setCellValue((String)value);
				}
				if(value instanceof Integer) {
					cell.setCellValue((Integer)value);
				}
				if(value instanceof Boolean) {
					cell.setCellValue((Boolean)value);
				}
			}
		}
		
		String filepath =".\\Data\\Student.xlsx";
		FileOutputStream outstream = new FileOutputStream(filepath);
		workbook.write(outstream);
		
		outstream.close();
		
		System.out.println("File Created Sucessfully");

	}

}
