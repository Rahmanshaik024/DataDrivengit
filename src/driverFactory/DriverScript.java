package driverFactory;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.Reporter;
import org.testng.annotations.Test;

import commonFunctions.FunctionLibrary;
import config.AppUtil;
import utilities.ExcelFileUtil;

public class DriverScript extends AppUtil {
String inputpath = "E:\\ShaikRahman\\DDT_Framework\\FileInput\\LoginData.xlsx";
String outputpath = "E:\\ShaikRahman\\DDT_Framework\\FileOutput\\DatadrivenResults.xlsx";
@Test
public void starttest() throws Throwable {
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	int rc = xl.rowCount("Login");
	Reporter.log("No of rows"+rc);
	for(int i=1;i<=rc;i++) {
		String username = xl.getCellData("Login", i, 0);
		String password = xl.getCellData("Login", i, 1);
		boolean res = FunctionLibrary.verify_login(username, password);
		if(res) {
			xl.setCellData("Login", i, 2, "Login Success", outputpath);
			xl.setCellData("login", i, 3, "pass",outputpath);
		}
		else {
			java.io.File screen = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(screen, new java.io.File("./Screenshot/Iteration/"+i+"Loginpage.png"));
			xl.setCellData("login", i, 2, "Login Fail", outputpath);
			xl.setCellData("Login", i, 3, "Fail", outputpath);
		}
	}
}
}
