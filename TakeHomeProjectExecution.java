package TakeHomeProject;

import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;



public class TakeHomeProjectExecution {
	
	TakeHomeProjectObjects object;
	WebDriver driver;
	
	@BeforeTest
	public void chromepage() {
		driver = new ChromeDriver();
		object = new TakeHomeProjectObjects(driver);
		PropertyConfigurator.configure("log4j.properties");
		//driver.manage().window().maximize();
		driver.get("https://www.galaxy.pk/");
	}
    @Test(description="Navigate To the Page",priority =1)
	public void NavigateToPage() throws Exception  {
		object.clicklaptopButton();
    }
    @Test(priority=2)
    public void TheLaptopNames() throws Exception {
		object.LaptopNames();
    }
    @Test(priority=3)
    public void TheExcelData() throws Exception {
		object.ExcelData();
    }
    
}