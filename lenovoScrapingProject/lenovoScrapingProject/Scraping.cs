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
    class Scraping
    {
        readonly Utility objUtility = new Utility();
        public static string ProductType, Series, SubSeries, FullSeriesName;
        readonly List<string> lstSeries = new List<string>();
        public static HtmlDocument doc;
        public void ScrapData()
        {
            doc = new HtmlDocument();
            objUtility.NavigateUrl("https://psref.lenovo.com/");
            Console.WriteLine("Please Wait. Redirecting to lenovo site.");


            string[] categories = { "Laptops", "Tablets", "Desktopsy_AIOs", "Workstations", "Monitors", "SmartCollaboration", "EdgeDevices" };

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
                        AllSeries = Utility.driver.FindElements(By.XPath("//div[@id='" + categories[y] + "']/div/div[@class='menu_productline_panel']/div/ul[@class='ul_productline']/li"));
                        if (AllSeries != null)
                        {
                            Series = AllSeries[i].Text;
                        }
                        Console.WriteLine("---------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("Series " + Series + " of Product Type " + ProductType + " Started Successfully");
                        Console.WriteLine("Product Type: " + ProductType);

                        AllSeries[i].Click();
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
                                                Console.WriteLine("SubSeries Link " + link + " is Extracted.");
                                            }

                                        }
                                    }
                                }

                            }
                        }
                        Console.WriteLine("All links of SubSeries Extracted Successfully!");

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
                            downloadLink.Click();

                            // Set the download directory path
                            string downloadDirectory = @"C:\Users\Administrator\Downloads\";

                            // Set the current directory path
                            string currentDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Data\\" + categories[y];
                            // Get all files in the download directory with the desired file extension
                            string[] files = Directory.GetFiles(downloadDirectory, "*" + ".xls");

                            if (files.Length > 0)
                            {
                                string sourceFilePath = files[0];
                                // Generate the new file name based on the file title
                                string newFileName = Series + "_" + SubSeries + "_" + FullSeriesName + ".xls";
                                string destinationPath = Path.Combine(currentDirectory, newFileName);

                                try
                                {
                                    // Move the file to the current directory
                                    File.Move(sourceFilePath, destinationPath);

                                    Console.WriteLine("File moved successfully.");

                                    Console.WriteLine("Columns added successfully.");


                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Error occurred while moving the file: " + ex.Message);
                                }
                            }
                            else
                            {
                                Console.WriteLine("No XLSX files found in the download directory.");
                            }


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
