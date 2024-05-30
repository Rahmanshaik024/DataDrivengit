package config;

import java.io.FileInputStream;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

public class AppUtil {
public static Properties con;
public static WebDriver driver;
@BeforeTest
public static void setup() throws Throwable {
	con = new Properties();
	con.load(new FileInputStream("E:\\ShaikRahman\\DDT_Framework\\PropertyFile\\Environment.properties"));
	if(con.getProperty("Browser").equalsIgnoreCase("chrome")) {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}
	else if(con.getProperty("Browser").equalsIgnoreCase("Firefox")) {
		driver = new FirefoxDriver();
	}
	else {
		Reporter.log("Browser Incorrect");
	}
}
@AfterTest
public static void teardown() {
	driver.close();
}

}
