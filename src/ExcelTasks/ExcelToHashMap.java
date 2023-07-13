package ExcelTasks;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHashMap {
	public static void main(String args[]) throws IOException {
		
		FileInputStream fis=new FileInputStream(".\\Data\\StudentHashMap.xlsx"); 
		XSSFWorkbook workbook=new XSSFWorkbook (fis); 
		XSSFSheet sheet=workbook.getSheet("HashdataTable");

		int rows=sheet.getLastRowNum();
		
		HashMap<String, String> data=new HashMap<String, String>();

		//Reading data from excel to HashMap 
		for(int i=0;i<=rows; i++)
		{

		String key=sheet.getRow(i).getCell(0).getStringCellValue(); 
		String value=sheet.getRow(i).getCell(1).getStringCellValue();
		data.put(key, value);
		}
		
		//Read Data From Excel
		for(Map.Entry entry: data.entrySet())
		{
			System.out.println(entry.getKey()+"   "+entry.getValue());

			}

	}

}
