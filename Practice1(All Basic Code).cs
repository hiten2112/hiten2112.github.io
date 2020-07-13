using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;

namespace SelAssignments
{
    [TestFixture]
    public class Practice1
    {
        IWebDriver driver;
        string url = "http://10.82.180.36/Common/Login.aspx";
        [SetUp]
        public void SeUpMethod()
        {
            driver = new ChromeDriver(@"C:\IVS Files\Selenium\Drivers\recentDrivers\chromedriver_win32_2.43");
            driver.Navigate().GoToUrl(url);
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
        }
        [Test]
        public void TestMethod()
        {
            //Login
            driver.FindElement(By.Id("body_txtUserID")).SendKeys("donhere");
            driver.FindElement(By.Id("body_txtPassword")).SendKeys("don@123");
            driver.FindElement(By.Id("body_btnLogin")).Click();
            Thread.Sleep(2000);

            //Fetch the "Welcome Donhere" text
            String text = driver.FindElement(By.XPath("/html/body/form/div[3]/div[2]")).Text;
            Console.WriteLine(text);
            Thread.Sleep(2000);

            //Table Data
           
            IWebElement table = driver.FindElement(By.Id("body_cph_MyAccount_gvAccountDetails"));
            IList<IWebElement> table_rows = table.FindElements(By.TagName("tr"));
            int size_rows = table_rows.Count;
           
            for(int i=1;i<=size_rows-3;i++)
            {
               IList<IWebElement> table_cols = table_rows[i].FindElements(By.TagName("td"));
               int size_col = table_cols.Count;
               
                    if(table_cols[3].Text=="Current Account")
                    {
                        Console.WriteLine("Account Number : "+table_cols[0].Text+"\t\t"+"Balance : "+table_cols[1].Text);
                    }
                

            }
            Thread.Sleep(2000);
            //Mouse Hovering
            Actions action = new Actions(driver);
            IWebElement ChequeBookButton = driver.FindElement(By.XPath("/html/body/form/div[3]/div[4]/div[1]/div/div/ul/li[6]/a"));
            action.MoveToElement(ChequeBookButton).Perform();

            //Implicit Wait
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            //Click on the button appear  after hovering
            driver.FindElement(By.XPath("/html/body/form/div[3]/div[4]/div[1]/div/div/ul/li[6]/ul/li[2]/a")).Click();
            Thread.Sleep(2000);

            //DropDown and Printing whole DropDown element
            IWebElement formsrc = driver.FindElement(By.Id("body_cph_MyAccount_ddlStatusType"));
            SelectElement selectFrom = new SelectElement(formsrc);
            IList<IWebElement> alloptions = selectFrom.Options;

            int size = alloptions.Count;
            for (int i = 1; i <= size - 1; i++)
            {
                Console.WriteLine(alloptions[i].Text);
            }

            //Click on Loans
            driver.FindElement(By.Id("GeneralTabMenu_Loans_li_Cust")).Click();

            //Reading data from excelsheet
            excel.Application x1app = new excel.Application(); //creating instance of excel app
            excel.Workbook x1workbook = x1app.Workbooks.Open(@"C:\Users\hitesh.jangid01\Desktop\EMI.xlsx"); //getting the actual excel file
            excel._Worksheet x1worksheet = x1workbook.Sheets[1]; //Getting sheet from where data will being fetched "here data in sheet 1"
            excel.Range x1range = x1worksheet.UsedRange; //All the data available

            int rowcount = x1range.Rows.Count;
            Console.WriteLine(rowcount);

            for (int i = 2; i <= rowcount; i++)
            {
                IWebElement toSelect1 = driver.FindElement(By.Id("body_cph_Loans_ddlLoanType"));
                SelectElement oSelect1 = new SelectElement(toSelect1);
                oSelect1.SelectByText(x1range.Cells[i,1].Value2);
                Thread.Sleep(2000);

                IWebElement toSelect2 = driver.FindElement(By.Id("body_cph_Loans_ddlLoanName"));
                SelectElement oSelect2 = new SelectElement(toSelect2);
                oSelect2.SelectByText(x1range.Cells[i, 2].Value2);
                Thread.Sleep(2000);

                driver.FindElement(By.Id("body_cph_Loans_txtReqLoanAmount")).SendKeys("" + x1range.Cells[i, 3].Value2);
                Thread.Sleep(2000);

                driver.FindElement(By.Id("body_cph_Loans_txtNoOfEMI")).SendKeys("" + x1range.Cells[i, 4].Value2);
                Thread.Sleep(2000);

                driver.FindElement(By.Id("body_cph_Loans_btnViewEMI")).Click();
                Thread.Sleep(2000);

                String emi_amount = driver.FindElement(By.Id("body_cph_Loans_lblEMIAmountText")).Text;

                //write back to excel file
                x1range.Cells[i, 5] = emi_amount;
            }
            x1workbook.SaveAs("EMiEdited.xlsx");
            Thread.Sleep(2000);
            driver.FindElement(By.LinkText("About Us")).Click();
            Thread.Sleep(2000);
        }


        [TearDown]
        public void TearDownMethod()
        {
            driver.Quit();
        }
    }
}
