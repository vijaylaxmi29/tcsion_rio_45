package ExcelTasks;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SQLExcelOperations {

	public static void main(String[] args) throws IOException, SQLException {
		
		//connect to database
		Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","root");
		
		//statement query
		Statement stmt = con.createStatement();
		ResultSet rs = stmt.executeQuery("select * from city");
		
		//Excel
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("ID");
		row.createCell(1).setCellValue("Name");
		row.createCell(2).setCellValue("CountryCode");
		row.createCell(3).setCellValue("District");
		row.createCell(4).setCellValue("Population");
		
		int r=1;
		while(rs.next()) {
			double Cid = rs.getDouble("ID");
			String Cname =rs.getString("Name");
			String Ccode =rs.getString("CountryCode");
			String Cdis =rs.getString("District");
			double Pno = rs.getDouble("Population");
			
			row=sheet.createRow(r++);
			
			row.createCell(0).setCellValue(Cid);
			row.createCell(1).setCellValue(Cname);
			row.createCell(2).setCellValue(Ccode);
			row.createCell(3).setCellValue(Cdis);
			row.createCell(4).setCellValue(Pno);
		}
		FileOutputStream fos =new FileOutputStream(".\\Data\\CityExcel.xlsx");
		workbook.write(fos);
		
		workbook.close();
		fos.close();
		con.close();
		System.out.println("Successfully Done.....!!!");
	}

}
