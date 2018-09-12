using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Excel;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;

namespace Amit_Project
{
    public class Class1
    {
        
        public static ChromeDriver  driver = new ChromeDriver();

        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

        //DesiredCapabilities capabilities = DesiredCapabilities.Equals();

        //capabilities.setCapability(ChromeOptions.CAPABILITY, chromeOptions);
        //capabilities.setJavascriptEnabled(true);
        //capabilities.setCapability(CapabilityType.PROXY, proxy);
        //capabilities.setCapability("acceptSslCerts", true); // Added this additionally
        //capabilities.setCapability("acceptInsecureCerts", true); // Added this additionally
        //capabilities.setCapability("ignore-certificate-errors", true); // Added this additionally

        [Test]
        public void Getdata()
        {
            
            Thread.Sleep(1000);
            string url = "https://www.realestate.com.au/sold/in-ardeer,+vic+3022/list-1?maxBeds=3&source=refinement";
            //driver.Navigate().GoToUrl("https://www.realestate.com.au/sold/with-1-bedroom-in-carlton/list-1?maxBeds=1&activeSort=solddate&source=refinement");
            //driver.Navigate().GoToUrl("https://www.realestate.com.au/sold/with-1-bedroom-in-carlton/list-2?maxBeds=1&activeSort=solddate");
            try
            { 
            driver.Navigate().GoToUrl(url);
            }
            catch(Exception)
            {
                driver.Navigate().GoToUrl(url);
            }
            Thread.Sleep(1000);

            driver.Manage().Window.Maximize();

            js.ExecuteScript("window.scrollBy(0,800)");

            Thread.Sleep(3000);

            string csv = string.Empty;
            string abspath1 = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]";

            csv = "Address,Price,BedRooms,BathRooms,CarParks,Sold On,PropertyType, Area, Sold By";
            csv += "\r\n";

            string XP = null;
            string NewValue = null;
            string ReturnText = null; 
            
            for (int k =1; k< 10; k++)
            { 

            for(int i =2; i<25; i++)
            {
                int j = 1;
                    if (k == 5)
                    {
                        j = 3;
                    }

                    XP = null;
                NewValue = null;

                //Address
                ///html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[3]/div[1]/div[1]/article[1]/div[4]/div[1]/div[2]/div[1]/h2[1]/a[1]/span[1]
                XP = abspath1 + "/div[" + j + "]"  + "/div[" + i + "]/article[1]/div[3]/div[1]/div[2]/div[1]/h2[1]/a[1]/span[1]";
                if(existsElement(XP,out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, ",");
                else
                {
                    XP = abspath1 + "/div[" + j + "]" +  "/div[" + i + "]/article[1]/div[4]/div[1]/div[2]/div[1]/h2[1]/a[1]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, ",");
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //Price
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[1]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, ",");
                else
                {
                    XP = XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[1]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ChangeString(ReturnText, ",");
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //Bedrooms
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[3]/ul[1]/li[1]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, ",");
                else
                {
                    XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[3]/ul[1]/li[1]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ChangeString(ReturnText, ",");
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //Bathrooms
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[3]/ul[1]/li[2]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ReturnText;
                else
                {
                    XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[3]/ul[1]/li[2]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ReturnText;
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //carparks
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[3]/ul[1]/li[3]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ReturnText;
                else
                {
                    XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[3]/ul[1]/li[3]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ReturnText;
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //SoldOnDate
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[2]/p[1]/span[2]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, "Sold on ");
                else
                {
                    XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[2]/p[1]/span[2]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ChangeString(ReturnText, "Sold on ");
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                //PropertyType
                XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[3]/div[1]/div[2]/p[1]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                    NewValue = ChangeString(ReturnText, "Sold on ");
                else
                {
                    XP = abspath1 + "/div[" + j + "]" + "/div[" + i + "]/article[1]/div[4]/div[1]/div[2]/p[1]/span[1]";
                    if (existsElement(XP, out ReturnText) == true)
                        NewValue = ChangeString(ReturnText, "Sold on ");
                }
                csv += NewValue;
                csv += ",";

                XP = null;
                NewValue = null;

                csv += "\r\n";

                if(i==8)
                    js.ExecuteScript("window.scrollBy(0,200)");

                js.ExecuteScript("window.scrollBy(0,500)");
            }
                XP = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[26]/div[1]/div[1]/div[1]/a[1]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                {
                    IWebElement nextbutton = driver.FindElement(By.XPath(XP));
                    nextbutton.Click();
                }

                XP = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[3]/div[11]/div[1]/div[1]/div[1]/a[1]";
                if (existsElement(XP, out ReturnText) == true)
                {
                    IWebElement nextbutton = driver.FindElement(By.XPath(XP));
                    nextbutton.Click();
                }
                XP = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[3]/div[26]/div[1]/div[1]/div[1]/a[1]/span[1]";
                if (existsElement(XP, out ReturnText) == true)
                {
                    IWebElement nextbutton = driver.FindElement(By.XPath(XP));
                    nextbutton.Click();
                }
            }

            string filePath = @"D:\Amit\Mydata.csv";
            File.WriteAllText(filePath, csv);

            driver.Quit();
        }

        private string ChangeString(string returnText, string v)
        {
            string newstring = returnText.Replace(v, string.Empty);
            return newstring;
        }

        private Boolean existsElement(String XP, out string ReturnText)
        {
            ReturnText = null;

            try
            {
                IWebElement element = driver.FindElement(By.XPath(XP));
                ReturnText = element.Text;               
            }
            
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        [Test]
        public void GetLandArea()
        {
            string csv2;
            csv2 = "Address,LandArea\r\n";

            driver.Navigate().GoToUrl("https://www.realestate.com.au/sold");
            Thread.Sleep(1000);

            driver.Manage().Window.Maximize();

            Thread.Sleep(3000);

            string ExcelPath = Resource1.ExcelPath;

                       ExcelLib.PopulateInCollection(ExcelPath, "Sheet2");
            for (int i = 7; i<8; i++)
            {
                string address;
                address = ExcelLib.ReadData(i, "Address");

                string landarea = null;
                Getla(address, out landarea);

                csv2 += address;
                csv2 += ",";
                csv2 += landarea;
                csv2 += "\r\n";
                driver.Navigate().GoToUrl("https://www.realestate.com.au/sold");
            }

            string filePath = @"D:\Amit\Mydata2.csv";
            File.WriteAllText(filePath, csv2);
            driver.Close();

        }

        public void Getla(string address, out string landarea)
        {

            //string ReturnText = null;
            //if (existsElement(XP, out ReturnText) == true)
            string ReturnText = null; 

            string XP = "/html[1]/body[1]/div[1]/div[1]/div[1]/form[1]/div[2]/div[1]/div[1]/div[1]/input[1]";
                IWebElement Search = driver.FindElement(By.XPath(XP));
            Search.Clear();
                Search.Click();
                Search.SendKeys(address);
                  Thread.Sleep(5000);

            XP = "/html[1]/body[1]/div[1]/div[1]/div[1]/form[1]/div[2]/div[1]/div[1]/div[1]/button[1]/span[1]";
            IWebElement button = driver.FindElement(By.XPath(XP));
            button.Click();

            Thread.Sleep(5000);
            Thread.Sleep(5000);

            js.ExecuteScript("window.scrollBy(0,500)");

            Thread.Sleep(5000);


            XP = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/article[1]/div[4]/div[1]/div[2]/div[1]/h2[1]/a[1]/span[1]";
            if (existsElement(XP, out ReturnText) == true)
            { 
                IWebElement Property = driver.FindElement(By.XPath(XP));
                Property.Click();
            }
            Thread.Sleep(5000);

            js.ExecuteScript("window.scrollBy(0,400)");

            Thread.Sleep(5000);

            XP = "/html[1]/body[1]/main[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[3]/a[1]";
            if (existsElement(XP, out ReturnText) == true)
            { 
                IWebElement History = driver.FindElement(By.XPath(XP));
                History.Click();
            }
            Thread.Sleep(5000);
            Thread.Sleep(5000);

            landarea = null;
            XP = "/html[1]/body[1]/div[1]/div[1]/main[1]/div[1]/div[2]/section[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]";
            if (existsElement(XP, out ReturnText) == true)
            {
                
                IWebElement element = driver.FindElement(By.XPath(XP));
                landarea = element.Text;
            }

            Thread.Sleep(5000);

           // driver.Navigate().GoToUrl("https://www.realestate.com.au/sold");
        }

        [Test]
        public void Proplink()
        {
            driver.Navigate().GoToUrl("https://www.realestate.com.au/property/26-upton-st-altona-vic-3018?source=property-search-p4ep");

            driver.Manage().Window.Maximize();

            IWebElement SearchAddress = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/form[1]/div[1]/div[1]/div[1]/div[1]/input[1]"));

            string ExcelPath = Resource1.ExcelPath;

            ExcelLib.PopulateInCollection(ExcelPath, "Sheet2");
                  string address;
                address = ExcelLib.ReadData(7, "Address");

                // Sending Address
                SearchAddress.SendKeys(address);

            Thread.Sleep(3000);

            //Selecting the address in drop down
            Actions action = new Actions(driver);   
            action.SendKeys(SearchAddress, Keys.ArrowDown).Perform();        
            action.SendKeys(SearchAddress, Keys.Return).Click();
            Thread.Sleep(1000);
            IWebElement Searchbutton = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/form[1]/div[1]/div[1]/div[1]/div[1]/button[1]"));
            Searchbutton.Click();

            Thread.Sleep(1000);

            IWebElement Abouttab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/main[1]/section[1]/div[1]/div[1]/div[2]/ul[1]/li[3]/a[1]"));
            Abouttab.Click();

            IWebElement LandArea = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/main[1]/div[1]/div[2]/section[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]"));
            string la = LandArea.Text;
        }

        public class ExcelLib
        {
            static List<Datacollection> dataCol = new List<Datacollection>();

            public class Datacollection
            {
                public int rowNumber { get; set; }
                public string colName { get; set; }
                public string colValue { get; set; }
            }


            public static void ClearData()
            {
                dataCol.Clear();  
            }


            private static DataTable ExcelToDataTable(string fileName, string SheetName)
            {
                // Open file and return as Stream
                using (System.IO.FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        excelReader.IsFirstRowAsColumnNames = true;

                        //Return as dataset
                        DataSet result = excelReader.AsDataSet();
                        //Get all the tables
                        DataTableCollection table = result.Tables;

                        // store it in data table
                        DataTable resultTable = table[SheetName];

                        //excelReader.Dispose();
                        //excelReader.Close();
                        // return
                        return resultTable;
                    }
                }
            }

            public static string ReadData(int rowNumber, string columnName)
            {
                try
                {
                    //Retriving Data using LINQ to reduce much of iterations

                    rowNumber = rowNumber - 1;
                    string data = (from colData in dataCol
                                   where colData.colName == columnName && colData.rowNumber == rowNumber
                                   select colData.colValue).SingleOrDefault();

                    //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;


                    return data.ToString();
                }

                catch (Exception e)
                {
                    //Added by Kumar
                    Console.WriteLine("Exception occurred in ExcelLib Class ReadData Method!" + Environment.NewLine + e.Message.ToString());
                    return null;
                }
            }

            public static void PopulateInCollection(string fileName, string SheetName)
            {
                ExcelLib.ClearData();
                DataTable table = ExcelToDataTable(fileName, SheetName);

                //Iterate through the rows and columns of the Table
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        Datacollection dtTable = new Datacollection()
                        {
                            rowNumber = row,
                            colName = table.Columns[col].ColumnName,
                            colValue = table.Rows[row - 1][col].ToString()
                        };


                        //Add all the details for each row
                        dataCol.Add(dtTable);

                    }
                }

            }
        }
    }
}
