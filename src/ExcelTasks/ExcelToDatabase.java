package ExcelTasks;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDatabase {

	public static void main(String[] args) throws IOException, SQLException {
		
		//connect to database
		Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","root");
				
		//statement query
		Statement stmt = con.createStatement();
		
		//Create a new table in the Database"places"
		String sql ="create table places(Name varchar(15),CountryCode varchar(15),District varchar(15))";
		stmt.execute(sql);
		
		//Excel
		FileInputStream fis =new FileInputStream(".\\Data\\WriteCityExcel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rows =sheet.getLastRowNum();
		
		for(int i=0;i<=rows;i++) {
			XSSFRow row = sheet.getRow(i);
			//double Cid = row.getCell(0).getNumericCellValue();
			String Cname = row.getCell(0).getStringCellValue();
			String Ccode = row.getCell(1).getStringCellValue();
			String Cdis = row.getCell(2).getStringCellValue();
			//double Pno = row.getCell(4).getNumericCellValue();
			
			sql="insert into places values('"+Cname+"','"+Ccode+"','"+Cdis+"')";
			stmt.execute(sql);
			stmt.execute("commit");
		}
		workbook.close();
		fis.close();
		con.close();
		
		System.out.println("Sucessfully Done....!!");
		
		 
				
		
	}

}
