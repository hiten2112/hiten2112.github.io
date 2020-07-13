using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

//log4net to be configured using the App.config xml file
[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace assignment_1075913
{
	[TestFixture]
	public class TestClass1
	{
		//setting up the log4net to access its methods
		private log4net.ILog log = log4net.LogManager.GetLogger(typeof(TestClass1));

		private IWebDriver driver;
		private string url = "http://10.82.180.36:81/Automation/MobileRecharge/#/home";


		[OneTimeSetUp]
		public void Setup()
		{
			log.Info("inside setup method");
			driver = new ChromeDriver(@"C:\IVS Files\Selenium\Drivers\recentDrivers\chromedriver_win32_2.43");
			driver.Navigate().GoToUrl(url);
			driver.Manage().Window.Maximize();

			//implicit wait
			driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(50);

			//page title
			string title = driver.Title;
			log.Info(title);







		}


		[Test]
		public void TestMethod()
		{
			log.Info("inside test method");
			driver.FindElement(By.LinkText("Login")).Click();

			driver.FindElement(By.Id("username")).SendKeys("Tony");
			driver.FindElement(By.Id("password")).SendKeys("TonyStark@123");
			driver.FindElement(By.XPath("/html/body/div[1]/div/div/form/div[3]/input")).Click();
			Thread.Sleep(7000);
			driver.FindElement(By.LinkText("DTH")).Click();
			Thread.Sleep(7000);
			driver.FindElement(By.XPath("/html/body/nav/ul[1]/li[3]/ul/li/a")).Click();
			Thread.Sleep(7000);

			//for drop down
			IWebElement select_city = driver.FindElement(By.XPath("//*[@id='back']/center/form/div[1]/select"));
			SelectElement city = new SelectElement(select_city);
			//IList<IWebElement> city_ddl = city.Options;
			city.SelectByValue("Bengluru");
			Thread.Sleep(7000);

			IWebElement location = driver.FindElement(By.XPath("//*[@id='back']/center/form/div[2]/select"));
			SelectElement loc = new SelectElement(location);
			loc.SelectByValue("Electronics City");
			Thread.Sleep(7000);

			//radio button
			driver.FindElement(By.XPath("//*[@id='back']/center/form/div[3]/input[1]")).Click();
			Thread.Sleep(7000);

			driver.FindElement(By.XPath("//*[@id='back']/center/form/div[4]/div/input")).Clear();
			Thread.Sleep(7000);

			driver.FindElement(By.XPath("//*[@id='back']/center/form/div[4]/div/input")).SendKeys("1");
			Thread.Sleep(7000);


			driver.FindElement(By.XPath("//*[@id='back']/center/form/div[5]/textarea")).SendKeys("Begur Hobli");
			///explicit wait
			WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(5000));

			driver.FindElement(By.XPath("//*[@id='back']/center/form/center/a")).Click();
			Thread.Sleep(7000);

			driver.FindElement(By.Id("wallet")).Click();
			Thread.Sleep(7000);
			driver.FindElement(By.XPath("//*[@id='quick']")).Click();
			Thread.Sleep(7000);

			//print meaasge apeear on screen
			string message = driver.FindElement(By.XPath("/html/body/div[5]/p")).Text;
			Console.WriteLine(message);


			driver.FindElement(By.XPath("/ html / body / div[5] / div[7] / div / button")).Click();
			Thread.Sleep(7000);

			driver.FindElement(By.XPath("//*[@id='success']/center/a[1]")).Click();
			Thread.Sleep(7000);

			driver.FindElement(By.Id("post")).SendKeys("8877998765");
			Thread.Sleep(7000);

			driver.FindElement(By.Id("roambsnl")).Click();
			Thread.Sleep(7000);
			string message1 = driver.FindElement(By.XPath("/ html / body / div[1] / div / div / div[2]")).Text;
			Console.WriteLine(message1);


			driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/button")).Click();
			Thread.Sleep(7000);

			driver.FindElement(By.Id("airteltopup")).Click();
			Thread.Sleep(7000);

			///table 
			IWebElement table = driver.FindElement(By.XPath("//*[@id='airtelto']"));
			IList<IWebElement> table_rows = table.FindElements(By.TagName("tr"));
			foreach (IWebElement row in table_rows)
			{
				IList<IWebElement> cols = row.FindElements(By.TagName("td"));
				foreach (IWebElement col in cols)
				{
					Console.WriteLine(col.Text);
				}


			}
			driver.FindElement(By.XPath("/html/body/div[1]/div/div/div/div[3]/button")).Click();
			Thread.Sleep(7000);
			//log.Error("cancel");

			driver.FindElement(By.XPath("/html/body/nav/ul[1]/li[2]/a")).Click();
			Thread.Sleep(7000);

			log.Error("postpaid in dropdown is not clickable");
			Thread.Sleep(9000);
			//excel file for input
			string excelpath = @"Z:\Data.xlsx";
			Application excelapp;
			Workbook wbook;
			Worksheet wsheet;
			excelapp = new Application();

			wbook = excelapp.Workbooks.Open(excelpath);
			wsheet = wbook.Worksheets["postpaid"];
			excelapp.Visible = true;

			int rowcount = wsheet.UsedRange.Rows.Count;
			//Console.WriteLine(rowcount);
			//int coloumncount = wsheet.UsedRange.Rows.Count;
			///Console.WriteLine(coloumncount);

			for (int i = 2; i <= rowcount; i++)
			{
				driver.FindElement(By.Id("post")).SendKeys(wsheet.Cells[i, 1].Text);
				Thread.Sleep(2000);
				driver.FindElement(By.Name("opernum")).SendKeys(wsheet.Cells[i, 2].Text);
				Thread.Sleep(2000);
				driver.FindElement(By.Name("posamo")).Clear();
				Thread.Sleep(2000);
				driver.FindElement(By.Name("posamo")).SendKeys(wsheet.Cells[i, 3].Text);
				Thread.Sleep(2000);
				driver.FindElement(By.XPath("//*[@id='postv']/form/center/a")).Click();
				Thread.Sleep(2000);

				driver.FindElement(By.Id("debit")).Click();
				Thread.Sleep(2000);
				driver.FindElement(By.Id("in1")).SendKeys("1234");
				Thread.Sleep(2000);
				driver.FindElement(By.Id("in2")).SendKeys("2345");
				Thread.Sleep(2000);
				driver.FindElement(By.Id("in3")).SendKeys("3456");
				Thread.Sleep(2000);
				driver.FindElement(By.Id("in4")).SendKeys("9456");
				Thread.Sleep(2000);


				//dropdown
				IWebElement month = driver.FindElement(By.Id("month"));
				SelectElement mon = new SelectElement(month);
				mon.SelectByValue("Feb");
				Thread.Sleep(2000);

				//dropdown
				IWebElement year = driver.FindElement(By.Id("year"));
				SelectElement yr = new SelectElement(year);
				yr.SelectByValue("2018");
				Thread.Sleep(2000);

				driver.FindElement(By.Name("cardCVV")).SendKeys("123");
				Thread.Sleep(2000);

				driver.FindElement(By.Id("confirm-purchase")).Click();
				Thread.Sleep(2000);


				driver.FindElement(By.XPath("/html/body/div[5]/div[7]/div/button")).Click();
				//screen shot
				Screenshot scrfile = ((ITakesScreenshot)driver).GetScreenshot();
				scrfile.SaveAsFile(@"C:\Users\apurva.tatekar\source\repos\image "+i+".png");

				Thread.Sleep(6000);

				///excel writeback
				string message2 = driver.FindElement(By.Id("successmessage")).Text;
				Thread.Sleep(2000);
				wsheet.Cells[i, 4] = message2;
				Thread.Sleep(6000);


				driver.FindElement(By.XPath("//*[@id='success']/center/a[2]")).Click();
				Thread.Sleep(2000);
			}

			wbook.Save();
			wbook.Close();
			excelapp.Quit();
			Thread.Sleep(2000);







		}




		[OneTimeTearDown]
		public void Teardown()
		{
			log.Info("inside tear down");
			driver.Close();
			driver.Quit();
		}
	}
}
