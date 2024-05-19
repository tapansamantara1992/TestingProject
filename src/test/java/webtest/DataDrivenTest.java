package webtest;

import java.io.File;
import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.io.FileHandler;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Test;

public class DataDrivenTest 
{
  
	  public static WebDriver driver;
	  @Test(dataProvider ="Excel",dataProviderClass = CustomData.class)
	  public void testdatadriven(String username,String password) throws InterruptedException {
		  driver = new ChromeDriver();
		  //driver.manage().window().maximize();
		  driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		  driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
		  
		  driver.findElement(By.name("username")).sendKeys(username);
		  driver.findElement(By.name("password")).sendKeys(password);
		  driver.findElement(By.xpath("//button[@type='submit']")).click();
		  
		    Thread.sleep(2000);
		  	Assert.assertTrue(driver.getCurrentUrl().contains("dashboard"),"Login Fail");
			  System.out.println("Login Successfull");
			  
			//  driver.quit();
	
		 
	  }
	  
	  public void captureScreenshot(Object fileName)
		{
			TakesScreenshot ts=(TakesScreenshot)driver;
			File src=ts.getScreenshotAs(OutputType.FILE);
			File dest=new File(System.getProperty("user.dir")+"//Screenshots//"+fileName);
			try {
				FileHandler.copy(src, dest);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.println("ScreenShot captured!");
		}

	 
  }

