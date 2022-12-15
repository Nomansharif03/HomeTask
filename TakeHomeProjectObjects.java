package TakeHomeProject;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TakeHomeProjectObjects {


	@FindBy(xpath="//a[@class='departments-menu-v2-title']")
	WebElement clickOnProducts;
	@FindBy(xpath="//h2[@class='woocommerce-loop-product__title']")
	WebElement LaptopNames;
	@FindBy(xpath="//ul[@id='menu-all-departments-1']//li[@id='menu-item-4761']")
	WebElement clickonlaptopButton;
	@FindBy(xpath="//div[@class=\"product-short-description\"]")
	WebElement LaptopDescription;
	@FindBy(xpath="//select[@name='ppp']")
	WebElement ClickOnNumberofProducts;
	
	@FindBy(xpath="//span[@class='woocommerce-Price-currencySymbol']")
	WebElement LaptopPrice;
	

	
	
	@FindBy(xpath="//img[@class='attachment-woocommerce_thumbnail size-woocommerce_thumbnail jetpack-lazy-image jetpack-lazy-image--handled']\r\n")
	WebElement LaptopImages;
	ArrayList<String> LaptopNames1 = new ArrayList<String>();
	ArrayList<String> LaptopPrice1 = new ArrayList<String>();
	ArrayList<String> LaptopDescription1 = new ArrayList<String>();
	WebDriver driver;
	Logger logger = Logger.getLogger("WeekendTaskObject");
	public TakeHomeProjectObjects(WebDriver driver1){
		driver=driver1;
		PageFactory.initElements(driver1,this);}
	
	  public void clicklaptopButton() throws Exception {
		  driver.manage().window().maximize();
			Actions hover=new Actions(driver);
			hover.moveToElement(clickOnProducts).perform();
			Thread.sleep(1000);
			clickonlaptopButton.click();
			Select dd=new Select(ClickOnNumberofProducts);
			dd.selectByVisibleText("Show All");
			Thread.sleep(2000);
		  
	  }
	  public void LaptopNames() throws Exception {
		  ArrayList<String> LaptopNames1 = new ArrayList<String>(); 
		   for(int names=2;names<160;names=names+2) {
			   Thread.sleep(15000);
		   List<WebElement> LaptopNames2 = driver.findElements(By.xpath("//h2[@class='woocommerce-loop-product__title'])[" + names + ")"));
		   Thread.sleep(15000);
			   LaptopNames1.add(LaptopNames.getText());
			   Thread.sleep(15000);
			   System.out.println(LaptopNames1);
		   }
		   ArrayList<String> LaptopPrice1 = new ArrayList<String>(); 
		   for(int price=2;price<167;price++) {
			   Thread.sleep(15000);
		   List<WebElement> LaptopPrice2=driver.findElements(By.xpath("//span[@class='electro-price'])["+ price +")"));
		   Thread.sleep(15000);
		   LaptopPrice1.add(((WebElement) LaptopPrice2).getText());
		   Thread.sleep(15000);
		   System.out.println(LaptopPrice1);
			   
		   }
	   
		   ArrayList<String> LaptopDescription1 = new ArrayList<String>(); 
		   for(int description=1;description<159;description++) {
			   Thread.sleep(15000);
		   List<WebElement> LaptopDescription2=driver.findElements(By.xpath("//div[@class='product-short-description'])["+ description +")"));
		   Thread.sleep(15000);
		   LaptopDescription1.add(((WebElement) LaptopDescription2).getText());
		   Thread.sleep(15000);
		   System.out.println(LaptopDescription1);
			  
			  
		   }
		   
	  }  	   
	   public void ExcelData() throws Exception {
		   
		   XSSFWorkbook WorkBook1=new XSSFWorkbook();
		   XSSFSheet WorkBooksheet = WorkBook1.createSheet("ExcelData");
		   WorkBooksheet.createRow(0);
		   WorkBooksheet.getRow(0).createCell(0).setCellValue("Laptop Names");
		   WorkBooksheet.getRow(0).createCell(1).setCellValue("Laptop Price");
		   WorkBooksheet.getRow(0).createCell(2).setCellValue("Laptop Description");

			for(int i=1;i<LaptopNames1.size();i=i++) {
				WorkBooksheet.createRow(i+1);
	            WorkBooksheet.createRow(i+1).createCell(0).setCellValue(LaptopNames1.get(i));
	            }
			for(int i=1;i<LaptopPrice1.size();i=i++) {
				WorkBooksheet.createRow(i+1);
	            WorkBooksheet.getRow(i+1).createCell(1).setCellValue(LaptopPrice1.get(i));
			}
			for(int i=1;i<LaptopDescription1.size();i=i++) {
				WorkBooksheet.createRow(i+1);
	            WorkBooksheet.getRow(i+1).createCell(2).setCellValue(LaptopDescription1.get(i));
			}
//			 List<WebElement> image=driver.findElements(By.xpath("//div[@class='product-thumbnail product-item__thumbnail']")); 
//			int count=1;
//			for(WebElement e:image) {
//				String src1=(e.getAttribute("srcset"));
//				
//				URL imageURL=new URL(src1);
//				BufferedImage saveimage=ImageIO.read(imageURL);
//				ImageIO.write(saveimage, "jpg", new File("count"+".jpg"));
//				count++;
//			}
			
			File file=new File("C:\\Users\\acer\\Desktop\\ExcelData.xlsx");
			FileOutputStream fileoutputstream=new FileOutputStream(file);
			WorkBook1.write(fileoutputstream);
	
  }	
	  
}



