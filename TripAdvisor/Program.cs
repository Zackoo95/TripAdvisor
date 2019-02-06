using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;
using NUnit.Framework;
using System.IO;
using System.Linq.Expressions;
using OpenQA.Selenium.Remote;


namespace TripAdvisor
{

    public static class Program
    {
       

        private static void NewMethod(IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
        }
        
        public static string GetInnerHtml(this IWebElement element)
        {
            var remoteWebDriver = (RemoteWebElement)element;
            var javaScriptExecutor = (IJavaScriptExecutor)remoteWebDriver.WrappedDriver;
            var innerHtml = javaScriptExecutor.ExecuteScript("return arguments[0].innerHTML;", element).ToString();

            return innerHtml;
        }

       


        static void Main(string[] args)
        {
            int y = 0;
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            object misvalue = System.Reflection.Missing.Value;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            oSheet.Cells[1, 1] = "HotelName";
            oSheet.Cells[1, 2] = "HotelAddress";
            oSheet.Cells[1, 3] = "HotelNumber";
            oSheet.Cells[1, 4] = "HotelSite";
            oSheet.get_Range("A1", "D1").Font.Bold = true;
            oSheet.get_Range("A1", "D1").VerticalAlignment =
            Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            string name;
            string address;
            string town;
            string number;
            string site;
            IWebDriver driver = new ChromeDriver();
            var options = new ChromeOptions();
            options.AddArgument("no-sandbox");
            driver.Manage().Window.Maximize();
            //open chrome web driver
            driver.Navigate().GoToUrl("https://www.tripadvisor.com/Hotels-g153339-Canada-Hotels.html");
            NewMethod(driver);
            driver.SwitchTo().ActiveElement();
            //IWebElement element1 = driver.FindElement(By.Id("BODY_BLOCK_JQUERY_REFLOW"));
            //element1.SendKeys(Keys.Escape);
            //driver.FindElement(By.XPath("//*[@id=\"component_2\"]/div/div/span[1]/div/div/a")).Click();
            //driver.SwitchTo().ActiveElement();
            //IWebElement search = driver.FindElement(By.XPath("//*[@id=\"c_targeted_flyout_1\"]/div/div/div[1]/div[1]/div/input"));
            //search.SendKeys("Canada");
            //NewMethod(driver);
            //driver.FindElement(By.XPath("//*[@id=\"c_targeted_flyout_1\"]/div/div/div[1]/div[2]")).Click();
            //search.SendKeys(Keys.Enter);
            
            
            while (driver.FindElements(By.LinkText("Next")).Count > 0)
            {

                driver.FindElement(By.Id("global-nav-hotels")).Click();
                List<IWebElement> hotels = driver.FindElements(By.ClassName("property_title")).ToList();


                for (int i = 0; i < hotels.Count; i++)
                {
                    var x = 0;

                    if (y == 0)
                    {
                        x = i + 2;
                    }
                    else if(y>0)
                    {
                        x = y + i + 2;
                    }
                    hotels[i].Click();
                    driver.SwitchTo().Window(driver.WindowHandles.LastOrDefault());
                    name = driver.FindElement(By.Id("HEADING")).Text.ToString();
                    oSheet.Cells[x, 1] = name;
                    if (driver.FindElements(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[1]/span[2]/span[1]")).Count > 0)
                    {
                        IWebElement add = driver.FindElement(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[1]/span[2]/span[1]"));
                        address = GetInnerHtml(add);
                        oSheet.Cells[x, 2] = address;
                    }
                   
                    if (driver.FindElements(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[2]/a/span[2]")).Count > 0)
                    {
                        number = driver.FindElement(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[2]/a/span[2]")).Text.ToString();
                        oSheet.Cells[x, 3] = number;
                    }
                   

                    if (driver.FindElements(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[3]/span[2]")).Count > 0)
                    {
                        driver.FindElement(By.XPath("//*[@id=\"taplc_resp_hr_atf_hotel_info_0\"]/div/div[2]/div/div[2]/div/div[3]/span[2]")).Click();
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                        site = driver.Url;
                        oSheet.Cells[x, 4] = site;
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    
                    driver.Close();
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                }
                driver.FindElement(By.LinkText("Next")).Click();
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                hotels.Clear();
                y += 30;
                driver.FindElement(By.LinkText("Next")).Click();
            }
            oSheet.SaveAs("Canada Hotels.xlsx");
            driver.Close();
        }   
    }
}
