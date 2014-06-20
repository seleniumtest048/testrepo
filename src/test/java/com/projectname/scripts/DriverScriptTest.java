package com.projectname.scripts;
import java.io.File;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.projectname.base.EmailReportUtil;
import com.projectname.base.Keywords;
import com.projectname.base.ReportUtil;
import com.projectname.utils.TestConstants;
import com.projectname.utils.TestUtil;

public class DriverScriptTest extends Keywords {
	
	public static String keyword;
	public static String stepDescription;
	public static String result;
	public static DriverScriptTest dstest;
	static float totalTestCaseCount,runTestCaseCount=0,failedTestCases;
	@BeforeClass
	public void beforeClass() throws Exception{
		initialize();
		log.info("Initialized All Resources Files");		
		setTestClassName(DriverScriptTest.this.getClass().getName());	
		log.info("Creatng Driver Script Object");
		dstest=new DriverScriptTest();
		log.info("Creating Test Suite");
		startTesting();
		log.info("Creating Test Suite For Email Report");
		emailStartTesting() ;		
	}
	@Test
	public void driverScript() throws Exception {
		log.info("Creating Suite In Detailed Test Report");
		ReportUtil.startSuite("Suite");
		log.info("Creating Suite In Mailing Report");
		EmailReportUtil.startSuite("Suite");
		// test data colom and test results starting row
		int colom=3,excelrow=1;
		int sno=1;
		log.info("Creating Excel Sheet For Test Results");
	//	generateExcel(sno);
		controlshet=controllerwb.getSheet("Suite");
		totalTestCaseCount=controlshet.getRows()-1;
		
		tdsheet=testdatawb.getSheet("Browser");
		for (int i = 1; i < tdsheet.getRows(); i++)
		{
			
			String browser=tdsheet.getCell(2,i).getContents();
			String run=tdsheet.getCell(3,i).getContents();
			String siteurl=tdsheet.getCell(4,i).getContents();
			switch (browser) {
			case "GC":
				if(run.equalsIgnoreCase("Y")){
					log.info("Launching Google Chrome.......");
					File gc = new File( path+TestConstants.CHROME_BROWSER_DIR_PATH);
					System.setProperty("webdriver.chrome.driver", gc.getAbsolutePath());
					try{
						Process P = Runtime.getRuntime().exec(path+TestConstants.AUTOIT_LOGIN_PATH);
						driver=new ChromeDriver();
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						P.destroy();
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}catch(Exception e){
						driver=new ChromeDriver();
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}
				}
			break;
			case "MF":
				if(browser.equalsIgnoreCase("MF") && run.equalsIgnoreCase("Y")){
					log.info("Launching Mozilla Firefox.......");
					try{
						Process P = Runtime.getRuntime().exec(path+TestConstants.AUTOIT_LOGIN_PATH);
						driver=new FirefoxDriver();
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						P.destroy();
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}catch(Exception e){
						driver=new FirefoxDriver();
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}
				}
				break;
			case "IE":
				if(browser.equalsIgnoreCase("IE") && run.equalsIgnoreCase("Y")){
					log.info("Launching Internet Explorer.......");
					File ie = new File( path+TestConstants.IE_BROWSER_DIR_PATH);
					System.setProperty("webdriver.ie.driver", ie.getAbsolutePath());
					try{
						Process P = Runtime.getRuntime().exec(path+TestConstants.AUTOIT_LOGIN_PATH);
						driver= new InternetExplorerDriver(); 
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						P.destroy();
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}catch(Exception e){
						System.setProperty("webdriver.ie.driver", ie.getAbsolutePath());
						driver= new InternetExplorerDriver(); 
						driverWait=new WebDriverWait(driver, 15);
						driver.get(siteurl);
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						driver.manage().window().maximize();
						controlscript(controlshet,testdatawb);
					}
					
				}
				break;
			case "BST":
				if(browser.equalsIgnoreCase("BST") && run.equalsIgnoreCase("Y")){
					String Username=tdsheet.getCell(5,i).getContents();
					String Key=tdsheet.getCell(6,i).getContents();
					String URLKey=tdsheet.getCell(7,i).getContents();
					String Browser=tdsheet.getCell(8,i).getContents();
					String Bversion=tdsheet.getCell(9,i).getContents();
					String OS=tdsheet.getCell(10,i).getContents();
					String OsVersion=tdsheet.getCell(11,i).getContents();
					browserStack(siteurl,Username,Key,URLKey,Browser,Bversion,OS,OsVersion);
					
				}
				break;
			default:
				break;
			}
		}
		
	}
	
	public  void controlscript(Sheet controlshet, Workbook testdatawb) throws Exception{
		int colom=3,excelrow=1;
		System.out.println("Control Sheet rows--"+controlshet.getRows());
		for (int i1 = 1; i1 < controlshet.getRows(); i1++) {
			log.info("Selecting Test Scenario From Controller File");
			String tsrunmode=controlshet.getCell(2,i1).getContents();
			if (tsrunmode.equalsIgnoreCase("Y")) {
				runTestCaseCount++;
				log.info("Navigate To Test Scenario Sheet");
				String tcaseid=controlshet.getCell(0,i1).getContents();
				Sheet tdsheet1=testdatawb.getSheet(tcaseid);
				//control sheet
				Sheet controlshet1=controllerwb.getSheet(tcaseid);
				String fileName=null;
				log.info("Loading Test Data Work Book");
				for (int j = 1; j < tdsheet1.getRows(); j++) {
					String tcaserunmode=tdsheet1.getCell(2,j).getContents();
					if (tcaserunmode.equalsIgnoreCase("y")) {
						String testcaseid=tdsheet1.getCell(0,j).getContents();
						String testdesc=tdsheet1.getCell(1,j).getContents();
						fileName = "Suite1_TC"+(testcaseid)+"_TS"+tcaseid+"_"+keyword+j+".png";
						stepDescription=testdesc;
						keyword=testcaseid;
						log.info("Passing Parameters Driver Script to ContolScript");
						result=controlScript(j, colom, tcaseid,controlshet1,testcaseid,stepDescription,keyword,fileName);
						if (failcount>=1 || rptFailCnt>=1) {
							result="Fail";
							log.info("Test Scenario Result --"+ result);
							if (failcount==0) {
								failedTestCases+=rptFailCnt;
							}else{
								failedTestCases+=failcount;
							}
							mainReport(keyword,result,fileName);
						//	createLabel(excelrow, testcaseid, testdesc, result);
							rptFailCnt=0;
							failcount=0;
						}else{
							result="Pass";
							log.info("Test Scenario Result --"+ result);
							mainReport(keyword,result,fileName);
						//	createLabel(excelrow, testcaseid, testdesc, result);
						}
						
					}  
					excelrow++;
				}
			}
			controlshet=controllerwb.getSheet("Suite");
		}
		float totalPassCount=runTestCaseCount-failedTestCases;
		System.out.println("Total Run Test Case Count ---"+runTestCaseCount);
		System.out.println("Total Failed Test Case Count--"+failedTestCases);
		System.out.println("Total Passed Test Cases ----"+totalPassCount);
		
//		wwb.write();
//		wwb.close();
		float passPercentage=(totalPassCount/runTestCaseCount)*100;
		float failPercentage=(failedTestCases/runTestCaseCount)*100;
		System.out.println(passPercentage);
		System.out.println(failPercentage);
		dstest.generateCsvFile(0,passPercentage,failPercentage);
//		generatePieChart();
		tdsheet=testdatawb.getSheet("Browser");
		closeBrowser();
	}
	
	public void browserStack(String siteurl, String username, String key, String uRLKey, String browser, String bversion, String oS, String osVersion) throws Exception{
		String USERNAME = username;
		String AUTOMATE_KEY = key;
		
		String URL = "http://" + USERNAME + ":" + AUTOMATE_KEY + uRLKey;
		DesiredCapabilities caps = new DesiredCapabilities();
		caps.setCapability("browser", browser);
		caps.setCapability("browser_version",bversion);
		caps.setCapability("os", oS);
		caps.setCapability("os_version", osVersion);
		caps.setCapability("browserstack.debug", "true");
		log.info("Launching  "+browser+".......");
		driver = new RemoteWebDriver(new URL(URL), caps);
		driver.get(siteurl);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		controlscript(controlshet,testdatawb);   
	}
	
	@AfterSuite
	public static void endScript() throws Exception{
		ReportUtil.updateEndTime(TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa"));
		EmailReportUtil.updateEndTime(TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa"));
		//mail();
	}
}