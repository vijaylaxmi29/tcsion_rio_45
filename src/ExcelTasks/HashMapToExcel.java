package ExcelTasks;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook=new XSSFWorkbook(); 
		XSSFSheet sheet = workbook.createSheet ("HashdataTable");
		
		Map<String, String> data=new HashMap<String, String>();
		data.put("1", "Arjun");
		data.put("2", "Bharat");
		data.put("3", "Chetan");
		data.put("4", "Dany");
		data.put("5", "Ekta");
		
		int rowno=0;
		
		for (Map.Entry entry:data.entrySet())
		{
			XSSFRow row=sheet.createRow(rowno++);
			
			row.createCell(0).setCellValue((String)entry.getKey()); 
			row.createCell(1).setCellValue((String)entry.getValue()); 
			

		}

		FileOutputStream fos=new FileOutputStream(".\\Data\\StudentHashMap.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("Written Succesfully..");




	}

}
