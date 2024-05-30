package commonFunctions;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.testng.Reporter;

import config.AppUtil;

public class FunctionLibrary extends AppUtil {
	public static boolean verify_login(String user , String pass) {
		driver.get(con.getProperty("Url"));
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.findElement(By.cssSelector(con.getProperty("Objuser"))).sendKeys(user);
		driver.findElement(By.cssSelector(con.getProperty("Objpass"))).sendKeys(pass);
		driver.findElement(By.cssSelector(con.getProperty("Objlogin"))).click();
		String expected = "dashboard";
		String actual = driver.getCurrentUrl();
		if(actual.contains(expected)) {
			Reporter.log("Login Success");
			return true;
		}
		else {
			String errormessage =driver.findElement(By.cssSelector(con.getProperty("Objerrror"))).getText();
			Reporter.log(errormessage);
		}
		return false;
	}

}
