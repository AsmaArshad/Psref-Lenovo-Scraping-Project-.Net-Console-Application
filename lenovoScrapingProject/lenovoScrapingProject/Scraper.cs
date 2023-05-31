using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using HtmlAgilityPack;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace lenovoScrapingProject
{
    class Scraper
    {
        readonly Utility objUtility = new Utility();
        public static string ProductType, Series, SubSeries, FullSeriesName;
        readonly List<string> lstSeries = new List<string>();
        public static HtmlDocument doc;
        public static string DataDirectory;
        public void ScrapData()
        {
            doc = new HtmlDocument();
            objUtility.NavigateUrl("https://psref.lenovo.com/");
            Console.WriteLine("Please Wait. Redirecting to lenovo site.");


            //string[] categories = { "Laptops", "Tablets", "Desktopsy_AIOs", "Workstations", "Monitors", "SmartCollaboration", "EdgeDevices" };
            string[] categories = { "Desktopsy_AIOs", "Workstations", "Monitors", "SmartCollaboration", "EdgeDevices" };

            for (int y = 0; y < categories.Length; y++)
            {
                var CategoryLink = Utility.driver.FindElement(By.XPath("//li[@data-classification='" + categories[y] + "']/a"));
                if (CategoryLink != null)
                {
                    CategoryLink.Click();
                }
                Console.WriteLine("Page is Redirecting to Category " + categories[y]);



                ProductType = CategoryLink.Text;
                var AllSeries = Utility.driver.FindElements(By.XPath("//div[@id='" + categories[y] + "']/div/div[@class='menu_productline_panel']/div/ul[@class='ul_productline']/li"));
                if (AllSeries != null)
                {
                    for (int i = 0; i < AllSeries.Count; i++)
                    {
                        doc.LoadHtml(Utility.driver.PageSource);
                        var All_Series = doc.DocumentNode.SelectNodes("//div[@id='" + categories[y] + "']/div/div[@class='menu_productline_panel']/div/ul[@class='ul_productline']/li");
                        if (All_Series != null)
                        {
                            Series = All_Series[i].InnerText;
                        }

                        var AllSeries1 = Utility.driver.FindElements(By.XPath("//div[@id='" + categories[y] + "']/div/div[@class='menu_productline_panel']/div/ul[@class='ul_productline']/li"));

                        Console.WriteLine("---------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("Series " + Series + " of Product Type " + ProductType + " Started Successfully");
                        Console.WriteLine("Product Type: " + ProductType);

                        AllSeries1[i].Click();
                        lstSeries.Clear();

                        var main_node = Utility.driver.FindElement(By.XPath("//div[@id='" + categories[y] + "']/div[@class='row']/div[@class='menu_series_panel']/div[@class='scrollbarset']/ul[" + (i + 1) + "]/li[last()]"));
                        if (main_node != null)
                        {
                            ((IJavaScriptExecutor)Utility.driver).ExecuteScript("arguments[0].scrollIntoView(true);", main_node);
                        }

                        doc.LoadHtml(Utility.driver.PageSource);

                        var SubSeriesList = doc.DocumentNode.SelectNodes("//div[@id='" + categories[y] + "']/div[@class='row']/div[@class='menu_series_panel']/div[@class='scrollbarset']/ul[" + (i + 1) + "]/li");
                        if (SubSeriesList != null)
                        {
                            for (int j = 1; j <= SubSeriesList.Count; j++)
                            {
                                var SubSeriesNode = doc.DocumentNode.SelectSingleNode("//div[@id='" + categories[y] + "']/div[@class='row']/div[@class='menu_series_panel']/div[@class='scrollbarset']/ul[" + (i + 1) + "]/li[" + j + "]/div/div[@class='li_s1 seriesName']");
                                if (SubSeriesNode != null)
                                {
                                    var SubSeries = SubSeriesNode.InnerText;
                                    var SeriesLinkNode = doc.DocumentNode.SelectNodes("//div[@id='" + categories[y] + "']/div[@class='row']/div[@class='menu_series_panel']/div[@class='scrollbarset']/ul[" + (i + 1) + "]/li[" + j + "]/div/div[@class='menu_product_panel']/div/ul/li");
                                    if (SeriesLinkNode != null)
                                    {
                                        for (int z = 1; z <= SeriesLinkNode.Count; z++)
                                        {
                                            var SeriesLink = doc.DocumentNode.SelectSingleNode("//div[@id='" + categories[y] + "']/div[@class='row']/div[@class='menu_series_panel']/div[@class='scrollbarset']/ul[" + (i + 1) + "]/li[" + j + "]/div/div[@class='menu_product_panel']/div/ul/li[" + z + "]/a");
                                            if (SeriesLink != null)
                                            {
                                                var SeriesFullName = SeriesLink.InnerText;
                                                var link = SeriesLink.GetAttributeValue("href", "");
                                                lstSeries.Add(SubSeries + "::" + SeriesFullName + "::" + link);
                                                Console.WriteLine("\nSubSeries Link " + link + " is Extracted.");
                                            }

                                        }
                                    }
                                }

                            }
                        }
                        Console.WriteLine("\nAll links of SubSeries Extracted Successfully!\n");

                        for (int l = 0; l < lstSeries.Count; l++)
                        {
                            var SeriesText = lstSeries[l];
                            var SplitSeries = SeriesText.Split("::");

                            SubSeries = SplitSeries[0];
                            Console.WriteLine("SubSeries: " + SubSeries);

                            FullSeriesName = SplitSeries[1];
                            Console.WriteLine("FullSeriesName: " + FullSeriesName);
                            var SeriesLink = "https://psref.lenovo.com" + SplitSeries[2];

                            objUtility.NavigateUrl(SeriesLink.ToString());


                            Thread.Sleep(1000);

                            var ModelsXpath = "//ul[@id='pro_ul_nav']/li[2]/span[@data-div='divModels']";
                            Utility.WaitForElement(ModelsXpath);

                            var ModelsLink = Utility.driver.FindElement(By.XPath(ModelsXpath));
                            if (ModelsLink != null)
                            {
                                ModelsLink.Click();
                                Console.WriteLine("Please Wait. Redirecting to Models");
                            }

                            Thread.Sleep(2000);

                            //download file

                            var downloadLink = Utility.driver.FindElement(By.XPath("//div[@class='export']/a"));
                            if (downloadLink != null)
                            {
                                downloadLink.Click();
                            }

                            Thread.Sleep(25000);
                            //Thread.Sleep(1000 * 60 * 2);

                            // Set the current directory path
                            DataDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Data\\" + categories[y];
                            string newFileName = Series + "_" + SubSeries + "_" + FullSeriesName + ".xls";
                            objUtility.Rename_MoveFile(newFileName, DataDirectory);
                            Console.WriteLine((l + 1) + " File out of " + lstSeries.Count + " Downloaded Successfully!\n");

                        }
                        Console.WriteLine("Series " + Series + " of Product Type " + ProductType + " Completed Successfully");
                        Console.WriteLine("---------------------------------------------------------------------------------------------------------\n");


                        CategoryLink = Utility.driver.FindElement(By.XPath("//li[@data-classification= '" + categories[y] + "']/a"));
                        if (CategoryLink != null)
                        {
                            CategoryLink.Click();
                        }
                    }
                }
            }
        }
    }
}
