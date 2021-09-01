package TimeSheetTaskCreation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.List;
import java.util.Scanner;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class ExecuteTaskCreation 
{
	static String currentDir = System.getProperty("user.dir");
	String excelFilePath;
	String encodedString;
	 
	static WebDriver driver;
	
	public Sheet extractDatafromExcel(String sheetName) throws IOException
	{
		File file =    new File(currentDir.replace("CockPitTaskCreation", "")+"\\TaskCreationData.xlsx");

	    //Create an object of FileInputStream class to read excel file

	    FileInputStream inputStream = new FileInputStream(file);

	    @SuppressWarnings("resource")
		Workbook workBook = new XSSFWorkbook(inputStream);
	    Sheet dataSheet = workBook.getSheet(sheetName);
	    return dataSheet;
	    
	}
	
	public void LaunchBrowser() throws InterruptedException
	{
		System.setProperty("webdriver.chrome.driver",currentDir+"//Drivers//chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://newprojectmonitoring-aqx92i181n.dispatcher.us2.hana.ondemand.com/index.html?hc_reset");		
		
	}
	
	public void ElementExist(String findBy, String Value) throws InterruptedException
	{
		
		if(findBy.toLowerCase().equals("id"))
		{
			while(driver.findElements( By.id(Value) ).size() == 0)
		    {
		    	Thread.sleep(7000);
		    }
		}
		if(findBy.toLowerCase().equals("xpath"))
		{
			while(driver.findElements( By.xpath(Value) ).size() == 0)
		    {
		    	Thread.sleep(7000);
		    }
		}
	}
	public void WindowsAuthentication() throws Exception
	{
		Sheet dataSheet = new ExecuteTaskCreation().extractDatafromExcel("Sheet1");
		String userName = dataSheet.getRow(1).getCell(0).getStringCellValue();
	    String password = dataSheet.getRow(1).getCell(1).getStringCellValue();
	    
		byte[] decodedBytes = Base64.getDecoder().decode(password);
		String decodedString = new String(decodedBytes);
		
		new ExecuteTaskCreation().ElementExist("id", "i0116");
		
		driver.findElement(By.id("i0116")).sendKeys(userName);
		driver.findElement(By.id("idSIButton9")).click();;	
		
		new ExecuteTaskCreation().ElementExist("id", "passwordInput");
		
		driver.findElement(By.id("passwordInput")).sendKeys(decodedString);	
		driver.findElement(By.id("submitButton")).click();	
		new ExecuteTaskCreation().ElementExist("id", "idTxtBx_SAOTCC_OTC");
		
		System.out.println("Enter the authentication code you have received on your mobile phone and press enter:");
		Scanner scanner = new Scanner(System.in);
	    String inputString = scanner.nextLine();
	    	    
	    driver.findElement(By.id("idTxtBx_SAOTCC_OTC")).sendKeys(inputString);
	    driver.findElement(By.id("idSubmit_SAOTCC_Continue")).click();
	    //Thread.sleep(4000);
	    new ExecuteTaskCreation().ElementExist("id", "idSIButton9");
	    driver.findElement(By.id("idSIButton9")).click();
	    
	    Thread.sleep(15000);
	    new ExecuteTaskCreation().ElementExist("id", "__item0-__clone0-imgNav");
	}	
	
	public void NavigateToTaskCreation() throws Exception  
	{
		Sheet dataSheet = new ExecuteTaskCreation().extractDatafromExcel("Sheet1");
		String teamName = dataSheet.getRow(1).getCell(2).getStringCellValue();
		driver.findElement(By.id("__item0-__clone0-imgNav")).click();		
		
		new ExecuteTaskCreation().ElementExist("xpath", ".//div[span[span[text()='"+teamName+"']]]/span[@aria-expanded='false']");
		List<WebElement> elements = driver.findElements(By.xpath(".//div[span[span[text()='"+teamName+"']]]/span[@aria-expanded='false']"));
		
		if(elements.get(0).getAttribute("aria-expanded").equals("false"))
		{
			elements.get(0).click();
			Thread.sleep(5000);
		}
		
		Actions dragger = new Actions(driver);
		WebElement draggablePartOfScrollbar = driver.findElement(By.id("__xmlview0--tblProjectTask-vsb"));
		int numberOfPixelsToDragTheScrollbarDown = 100;
		dragger.moveToElement(draggablePartOfScrollbar).clickAndHold().moveByOffset(0,numberOfPixelsToDragTheScrollbarDown).release().perform();
		dragger.moveToElement(draggablePartOfScrollbar).clickAndHold().moveByOffset(0,numberOfPixelsToDragTheScrollbarDown).release().perform();
	}
	
	public void AddTask() throws Exception
	{
		Sheet dataSheet = new ExecuteTaskCreation().extractDatafromExcel("Sheet2");
		DataFormatter formatter = new DataFormatter();
		int rowCount = dataSheet.getLastRowNum();
		for(int i=0 ; i<rowCount; i++)
		{
			Cell Cell1 = dataSheet.getRow(i+1).getCell(0);
			String sprintNum = formatter.formatCellValue(Cell1);
			String title = driver.findElement(By.xpath("//tr[td[div[span[span[text()='Sprint "+sprintNum+"']]]]]")).getAttribute("title");
			if(title.toLowerCase().equals("click to select"))
			{
				driver.findElement(By.xpath("//tr[td[div[span[span[text()='Sprint "+sprintNum+"']]]]]")).click();
			}
			Thread.sleep(2000);
			
			driver.findElement(By.id("__xmlview0--nt-img")).click();	
			Thread.sleep(5000);
			
			String parent=driver.getWindowHandle();
			driver.switchTo().window(parent);
			
			driver.findElement(By.id("idTaskCategory-arrow")).click();
			Cell Cell2 = dataSheet.getRow(i+1).getCell(1);
			String test1 = formatter.formatCellValue(Cell2);
			if(test1.contains(" "))
			{
				String[] test2 = test1.split(" ");
				if(test2[0].contains("-"))
				{
					test2 = test2[0].split("-");
				}
				List<WebElement> multiElementsinList= driver.findElements(By.xpath("//ul[li[span[text()='"+test2[0]+"']]]/li/span[text()='"+test2[0]+"']"));
				for(WebElement element : multiElementsinList)
				{		
					String text = element.getText();
					if(element.getText().equals(test1))
					{
						element.click();
						break;
					}
				}
			}
			else
			{
				driver.findElement(By.xpath("//ul[li[span[text()='"+test1+"']]]/li/span[text()='"+test1+"']")).click();
			}
			Cell Cell3 = dataSheet.getRow(i+1).getCell(2);
			String date = formatter.formatCellValue(Cell3);
			
			Cell Cell4 = dataSheet.getRow(i+1).getCell(3);
			String description = formatter.formatCellValue(Cell4);
			
			Cell Cell5 = dataSheet.getRow(i+1).getCell(4);
			String plannedEfforts = formatter.formatCellValue(Cell5);
			
			driver.findElement(By.xpath("//span[@class='sapUiCalItemText'][text()='"+date+"']")).click();
			driver.findElement(By.xpath("//span[@class='sapUiCalItemText'][text()='"+date+"']")).click();		
			driver.findElement(By.id("idTaskDesc-inner")).sendKeys(description);
			driver.findElement(By.id("idEfforts-inner")).sendKeys(plannedEfforts);
			
			driver.findElement(By.id("idSkill-arrow")).click();	
			Cell Cell6 = dataSheet.getRow(i+1).getCell(5);
			String test3 = formatter.formatCellValue(Cell6);
			if(test3.contains(" "))
			{
				String[] test4 = test3.split(" ");
				if(test4[0].contains("-"))
				{
					test4 = test4[0].split("-");
				}
				List<WebElement> multiElementsinList= driver.findElements(By.xpath("//ul[li[text()='"+test4[0]+"']]/li[text()='"+test4[0]+"']"));
				for(WebElement element : multiElementsinList)
				{				
					if(element.getText().equals(test3))
					{
						element.click();
						break;
					}
				}
			}
			else
			{
				driver.findElement(By.xpath("//ul[li[text()='"+test3+"']]/li[text()='"+test3+"']")).click();
			}
			
			driver.findElement(By.id("idSubSkill-arrow")).click();		
			Thread.sleep(3000);
			
			Cell Cell7 = dataSheet.getRow(i+1).getCell(6);
			String test5 = formatter.formatCellValue(Cell7);
			if(test5.contains(" "))
			{
				String[] test6 = test5.split(" ");
				if(test6[0].contains("-"))
				{
					test6 = test6[0].split("-");
					new ExecuteTaskCreation().ElementExist("xpath", ".//ul[li[text()='"+test6[0]+"']]/li[text()='"+test6[0]+"'][contains(@id,'idSubSkill')]");
				}
				List<WebElement> multiElementsinList= driver.findElements(By.xpath("//ul[li[text()='"+test6[0]+"']]/li[text()='"+test6[0]+"'][contains(@id,'idSubSkill')]"));
				for(WebElement element : multiElementsinList)
				{				
					String suSkill = element.getText();
					if(suSkill.equals(test5))
					{
						element.click();
						break;
					}
				}
			}
			else
			{
				driver.findElement(By.xpath("//ul[li[text()='"+test5+"']]/li[text()='"+test5+"'][contains(@id,'idSubSkill')]")).click();
			}
			
			//driver.findElement(By.xpath("//bdi[text()='OK']")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//bdi[text()='Cancel']")).click();
	}
		
	}
	
	public void ValidateTaskAdded()
	{
		
	}
	
	public void CloseBrowser() throws InterruptedException
	{
		driver.close();
		driver.quit();
		driver = null;
	}
	
	public static void main(String[] args) throws InterruptedException 
	{
		try
		{
			int i = 20;
			String name = "kumar";
			System.out.println(i);
			System.out.println(name);
			ExecuteTaskCreation executeTaskCreation = new ExecuteTaskCreation();
			executeTaskCreation.LaunchBrowser();
			executeTaskCreation.WindowsAuthentication();
			executeTaskCreation.NavigateToTaskCreation();
			executeTaskCreation.AddTask();
			executeTaskCreation.CloseBrowser();
		}
		catch(Exception e)
		{
			ExecuteTaskCreation executeTaskCreation = new ExecuteTaskCreation();
			executeTaskCreation.CloseBrowser();	
			System.out.println(e);
		}
		// TODO Auto-generated method stub

	}

}
