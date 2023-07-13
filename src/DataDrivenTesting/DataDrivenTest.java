package DataDrivenTesting;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataDrivenTest {
	
	 WebDriver driver;
	 
     @BeforeClass
     public void setup()

     {

            System.setProperty("webdriver.chrome.driver","C://POIjar_files/chromedriver_win32/chromedriver.exe");

            driver=new ChromeDriver();
            driver.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
            driver.manage().window().maximize();

     }

     @Test(dataProvider="Data")
    public void loginTest(String username,String password,String result){
    	 
    	 driver.get("https://admin-demo.nopcommerce.com/login");
             WebElement txtEmail=driver.findElement(By.id("Email"));
             txtEmail.clear();
             txtEmail.sendKeys(username);
             WebElement txtPassword=driver.findElement(By.id("Password"));
             txtPassword.clear();
             txtPassword.sendKeys(password);

             driver.findElement(By.xpath("//input[@value='Log in']")).click(); //Login  button
             String result_title="Dashboard / nopCommerce administration";
             String act_title=driver.getTitle();

             if(result.equals("Valid"))
             {
                    if(result_title.equals(act_title))
                    {
                          driver.findElement(By.linkText("Logout")).click();
                          Assert.assertTrue(true);
                    }
                    else
                    {
                          Assert.assertTrue(false);
                    }

             }

             else if(result.equals("Invalid"))
             {
                    if(result_title.equals(act_title))

                    {

                          driver.findElement(By.linkText("Logout")).click();

                          Assert.assertTrue(false);

                    }

                    else

                    {

                          Assert.assertTrue(true);

                    }

             }
     }

     @DataProvider(name="Data")
     public String [][] getData() throws IOException
     {
/*
            String Data[][]= {

                                                     {"sushant@yahoo.com","pass@123","Valid"},

                                                     {"yahoo@sushant.com","123","Invalid"},

                                                     {"sus@hantyahoo.com","pass@123","Invalid"},

                                                     {"com.sushant@yahoo","1@23","Invalid"}

                                              };
*/
    	 //get the data from excel

         String path=".\\Data\\loginexcel.xlsx";
         XLUtility xlutil=new XLUtility(path);

         int totalrows=xlutil.getRowCount("Sheet1");
         int totalcols=xlutil.getCellCount("Sheet1",1);    

                     

         String Data[][]=new String[totalrows][totalcols];

         for(int i=1;i<=totalrows;i++) //1
         {
                for(int j=0;j<totalcols;j++) //0
                {
                    Data[i-1][j]=xlutil.getCellData("Sheet1", i, j);

                }                   
         }
            return Data;

     } 

     @AfterClass
     void tearDown(){
            driver.close();
     }

}
