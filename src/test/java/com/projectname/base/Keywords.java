package com.projectname.base;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.ArrayList;
import jxl.Sheet;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.projectname.scripts.TestReport;

public class Keywords extends FunctionLibrary {

	
	/**
	 * Retrieving objects form control sheet
	 * @param i
	 * @param colom
	 * @param tdshetnum
	 * @param csheet
	 * @param fileName 
	 * @param keyword2 
	 * @param stepDescription 
	 * @throws Exception 
	 */
	 public String controlScript(int row, int colom, String tdshetnum, Sheet csheet, String testcaseid, String stepDescription, String keyword2, String fileName) throws Exception{
		 String result="res";
		 controlshet=csheet;
		 steps = new ArrayList<TestReport>();
		 Robot r=new Robot();
		 for (int k = 1; k < controlshet.getRows(); k++) {
			 tr=new TestReport();
			 String desc=controlshet.getCell(1, k).getContents();
			 String keyword=controlshet.getCell(2, k).getContents().toUpperCase();
			 String keywordtype=controlshet.getCell(3,k).getContents();
			 String objectProp=controlshet.getCell(4,k).getContents();
			 String object=OR.getProperty(objectProp);
			 String data=null;
			 Object testdata=null;
			
			  try{
				  switch(keyword){
				  case "RADIOBUTTON": case "BUTTON": case "CHECKBOX":
					  testdata= testData(colom,row,tdshetnum);
					  data=(String) testdata;
					  log.info("Clicking on "+data+" Radiobutton");
					  String buttonvalue=OR.getProperty(data);
					  actionElement(keyword,keywordtype,buttonvalue,data);
					  colom++;
					  break;
				  case "LINKS": 
					  testdata= testData(colom,row,tdshetnum);
					  data=(String) testdata;
					  log.info("Clicking on "+data);
					  actionElement(keyword,keywordtype,data,data);
					  colom++;
					  break;
				  case "CLICK":
					  log.info("Clicking on Button " + object);
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "INPUT":
					  log.info("Entering Data into "+object);
					  testdata= testData(colom,row,tdshetnum);
					  data=(String) testdata;
					  actionElement(keyword,keywordtype,object,data);
					  colom++;
					  break;
				  case "SELECT":
					  testdata= testData(colom,row,tdshetnum);
					  data=(String) testdata;
					  log.info("Selecting DropDown Value "+data);
					  actionElement(keyword,keywordtype,object,data);
					  colom++;
					  break;
				  case "WAIT":
					  log.info("Loading Page");
					Thread.sleep(10000);
					  break;
				  case "URL":
					  driver.get(object);
					  break;
				  case "VERIFYTEXT":
					  log.info("Verifying Text");
					result= verifyText(keyword,keywordtype,object,objectProp);
					  break;
				  case "COMPAIRTEXT":
					  log.info("Compairing Text");
					  result=  compairText(keyword,keywordtype,object,objectProp);
					  break;
				  case "NEWWINDOW":
					  log.info("Switch to New Window");
					  windowhandle();
				  break;
				  case "KEYDOWN": 
					  log.info("Clicking on Down Button");
					  Thread.sleep(3000);
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "ESC": 
					  log.info("Clicking on Escape Button");
					  Thread.sleep(3000);
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "ENTERKEY": 
					  log.info("Clicking on Enter Key");
					  
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "TAB":
					  log.info("Clicking on TAB Button");
					  actionElement(keyword,keywordtype,object,data);
					  break;	
				  case "CLOSEWINDOW":
					  log.info("Closing Child Window");
					  closeWindow();
				  break;
				  case "PARENTWINDOW":
					  log.info("Switching to Parent Window");
					  Mainwindow();
					  break;
				  case "FRAME":
					  log.info("Switch to frame------");
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case"ALERTVERIFY":
					  log.info("Verifying Alert-----");
					  testdata= testData(colom,row,tdshetnum);
					  data= (String) testdata;
					  alertverify(data);
					  colom++;
					  break;
				  case "CLOSE": case "CLOSEBROWSER":
					  log.info("Closing Browser");
					  closeBrowser();
					  break;
				  case "ALERTACCEPT":
					  log.info("Accepting Alert");
					  driver.switchTo().alert().accept();
					  break;
				  case "GETLINKS": 
					  log.info("Getting All Links from Webpage");
					  getLinks();
					  break;
				  case "GETTEXT": 
					  log.info("Getting Text From Webpage");
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "PASTETEXT": 
					  log.info("Paste Text into Textbox");
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "GETTITLE":
					  log.info("Getting Page Title");
					  String title=driver.getTitle();
					  break;
				  case "MOUSEOVER":
					  log.info("Mouse Over to "+object);
					  actionElement(keyword,keywordtype,object,data);
					  break;
				  case "AUTHENTICATION":
					  log.info("Executing VBScript");
					  authentication();
				  break;
				  case "SCROLLBAR":
					  log.info("Scroll down");
					  actionElement(keyword,keywordtype,object,data);
				  break;
				  case "GETATTRIBUTE":
					  log.info("Getting the value form WebPage-----");
					 result=getAttribute(keyword,keywordtype,object,objectProp);
				  break;
				  case "FORWARD":
					  driver.navigate().forward();
				  break;
				  case "BACK":
					  log.info("Clicking on Back Button");
					   driver.navigate().back();
				  break;
				  case "NEWTAB":
					  log.info("Creating Tab");
					  r.keyPress(KeyEvent.VK_CONTROL); 
					  r.keyPress(KeyEvent.VK_T);  
				  break;
				  default:
					  break;
				  }
				  reportSteps(result,desc,keyword,fileName,object,testcaseid);
				  result="res";
				  
			  }catch(UnhandledAlertException e){
				  driver.switchTo().alert().accept();
				  
			  }catch(Exception e)
			  {	
				  result="Fail";
				  log.info("",e);
				  failcount++;
				  report(result,desc,keyword,fileName,object,testcaseid);
				  e.printStackTrace();
				 	tr.result=result;
					tr.desc=desc;
					tr.keyword=keyword;
					tr.fileName=fileName;
					tr.object=object;
					tr.testcaseid=testcaseid;
					steps.add(tr);
				break;
			  }
		  }
		 reportEmailMain(result,fileName,testcaseid,failcount);
		 for (int i = 0; i < steps.size(); i++) {
			 String re=steps.get(i).result;
				String ds=steps.get(i).desc;
				String kw=steps.get(i).keyword;
				String fn=steps.get(i).fileName;
				String ob=steps.get(i).object;
				String tcid=steps.get(i).testcaseid;
				switch (re) {
				case "Fail":
					 		EmailReportUtil.addTestCaseSteps(ds,re);
					break;
				case "Pass": case "res":
					 		EmailReportUtil.addTestCaseSteps(ds,re);
					break;
				default:
					break;
				}
		 }
		 return result;
	 }
	 public static WebElement actionElement(String keyword, String keywordtype,String object,String data) throws Exception  {
		 WebElement welement=null;
		
			 switch (keyword.toUpperCase()) {
			 case "CLICK":
			 	welement=welement(keywordtype, object);
			 	welement.click();
				break;
			 case "INPUT":
			  	welement=welement(keywordtype, object);
			  	welement.clear();
			  	welement.sendKeys(data);
				break;
			 case "SELECT":
				 welement=welement(keywordtype, object);
				 new Select(welement).selectByVisibleText(data);
				 break;
			 case "VERIFYTEXT":
				 welement=welement(keywordtype, object);
				 break;
			 case "COMPAIRTEXT":
				 welement=welement(keywordtype, object);
				 break;
			 case "WAIT":
				 Thread.sleep(Long.parseLong(data));
				 break;
			 case "RADIOBUTTON": case "BUTTON":
				 System.out.println("object is--"+ object);
				 welement=welement(keywordtype, object);
				 welement.click();
				  break;
			 case "LINKS": 
				 welement=welement(keywordtype, data);
				 welement.click();
				 break;
			 case "MOUSEOVER":
				 welement=welement(keywordtype, object);
				 mouseover(welement);
				 break;
			 case "FRAME":
				 welement=welement(keywordtype, object);
				 frame(welement);
				 break;	
			 case "GETTEXT":
				 welement=welement(keywordtype, object);
				 pastevalue=  getText(welement);
				 break;
			 case "PASTETEXT": 
				 welement=welement(keywordtype, object);
				 welement.sendKeys(pastevalue);
				 break;
			  case "KEYDOWN": 
				  welement=welement(keywordtype, object);
				  welement.sendKeys(Keys.DOWN);
				  break;
			  case "ESC": 
				  log.info("Clicking on Escape Button");
				  welement=welement(keywordtype, object);
				  welement.sendKeys(Keys.ESCAPE);
				  break;
			  case "ENTERKEY": 
				  welement=welement(keywordtype, object);
				  welement.sendKeys(Keys.ENTER);
				  break;
			  case "TAB":
				  welement=welement(keywordtype, object);
				  welement.sendKeys(Keys.TAB);
				  break;
			  case "SCROLLBAR":
				  log.info("Scroll down");
				  welement=welement(keywordtype, object);
				  scrollBar(welement);
			  break;
			  case "GETATTRIBUTE":
				  log.info("Getting the value form WebPage-----");
				  welement=welement(keywordtype, object);
				 
			  break;
			  default:
				  break;
			  }
		
			return welement;
	 }
	 
	 public static WebElement welement(String keywordtype, String object) throws Exception{
			WebElement welement =null;
		try{	
			switch (keywordtype.toUpperCase()) {
			case "ID":
				driverWait.until(ExpectedConditions.visibilityOfElementLocated(By.id(object)));
				welement=driver.findElement(By.id(object));
				return welement ;
			case "XPATH":
				driverWait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(object)));
				welement=driver.findElement(By.xpath(object));
				return welement ;
			case "LINKTEXT":
				driverWait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(object)));
				welement=driver.findElement(By.linkText(object));
				return welement ;
			case "PLINK":
				driverWait.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(object)));
				welement=driver.findElement(By.partialLinkText(object));
				return welement ;
			case "CLASSNAME":
				driverWait.until(ExpectedConditions.visibilityOfElementLocated(By.className(object)));
				welement=driver.findElement(By.className((object)));
				return welement ;
			default:
				break;
			}
		}catch(Exception e){
			switch (keywordtype.toUpperCase()) {
			case "Id":
				Thread.sleep(900);
				welement=driver.findElement(By.id(object));
				return welement ;
			case "XPATH":
				Thread.sleep(900);
				welement=driver.findElement(By.xpath(object));
				return welement ;
			case "LINKTEXT":
				Thread.sleep(900);
				welement=driver.findElement(By.linkText(object));
				return welement ;
			case "PLINK":
				Thread.sleep(900);
				welement=driver.findElement(By.partialLinkText(object));
				return welement ;
			case "CLASSNAME":
				Thread.sleep(900);
				welement=driver.findElement(By.className((object)));
				return welement ;
			default:
				break;
			
			}
			e.getMessage();
		}
		return welement;
	 }

}
