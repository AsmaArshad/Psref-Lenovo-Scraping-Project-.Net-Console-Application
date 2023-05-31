using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace lenovoScrapingProject
{
    class WithdrawnProduct
    {
        Utility objUtility = new Utility();
        public static string ProductType, Series, SubSeries, FullSeriesName;
        public static HtmlDocument doc;
        readonly List<string> lstSeries = new List<string>();
        public static string DataDirectory;

        public void WithdrawnData()
        {
            doc = new HtmlDocument();
            objUtility.NavigateUrl("https://psref.lenovo.com/WithdrawnProducts");
            string[] categories = { "Monitors" };
            //string[] categories = { "Laptops", "Tablets", "Desktops & AIOs", "Workstations", "Servers", "Monitors" };
            for (int y = 0; y < categories.Length; y++)
            {
                SelectElement selectElement = new SelectElement(Utility.driver.FindElement(By.Id("ddlclassification")));
                selectElement.SelectByText(categories[y]);
                ProductType = categories[y];


                var ProductDropdown = Utility.driver.FindElements(By.XPath("//select[@id='ddlproductline']/option"));
                for (int i = 1; i < ProductDropdown.Count; i++)
                {
                    if (i > 1)
                    {
                        SelectElement selectProductFamily = new SelectElement(Utility.driver.FindElement(By.Id("ddlclassification")));
                        selectProductFamily.SelectByText(categories[y]);
                        ProductType = categories[y];
                    }

                    var productNode = Utility.driver.FindElement(By.XPath("//select[@id='ddlproductline']/option[" + (i + 1) + "]"));
                    if (productNode != null)
                    {
                        SelectElement selectProduct = new SelectElement(Utility.driver.FindElement(By.Id("ddlproductline")));
                        selectProduct.SelectByText(productNode.Text);
                        Series = productNode.Text;

                        doc.LoadHtml(Utility.driver.PageSource);
                        lstSeries.Clear();
                        var AllSeries = doc.DocumentNode.SelectNodes("//div[@id='withdoawnproductlistconetent']/div[@class='seriesDiv']");
                        for (int j = 0; j <= AllSeries.Count; j++)
                        {

                            var SeriesNode = doc.DocumentNode.SelectSingleNode("//div[@id='withdoawnproductlistconetent']/div[@class='seriesDiv'][" + j + "]/h2");
                            if (SeriesNode != null)
                            {
                                SubSeries = SeriesNode.InnerText;
                                var SubSeriesNode = doc.DocumentNode.SelectNodes("//div[@id='withdoawnproductlistconetent']/div[@class='seriesDiv'][" + j + "]/div[@class='row']/div");
                                for (int l = 0; l <= SubSeriesNode.Count; l++)
                                {
                                    var SubSeriesName = doc.DocumentNode.SelectSingleNode("//div[@id='withdoawnproductlistconetent']/div[@class='seriesDiv'][" + j + "]/div[@class='row']/div[" + l + "]/a");
                                    if (SubSeriesName != null)
                                    {
                                        FullSeriesName = SubSeriesName.InnerText;
                                        var link = "https://psref.lenovo.com" + SubSeriesName.GetAttributeValue("href", "");
                                        lstSeries.Add(SubSeries + "::" + FullSeriesName + "::" + link);
                                        Console.WriteLine("\nSubSeries Link " + link + " is Extracted.");
                                    }
                                }
                            }
                        }
                        Console.WriteLine("\nAll links of SubSeries Extracted Successfully!\n");

                        for (int z = 0 ; z < lstSeries.Count; z++)
                        {
                            var SeriesText = lstSeries[z];
                            var SplitSeries = SeriesText.Split("::");

                            SubSeries = SplitSeries[0];
                            Console.WriteLine("SubSeries: " + SubSeries);

                            FullSeriesName = SplitSeries[1];
                            Console.WriteLine("FullSeriesName: " + FullSeriesName);
                            objUtility.NavigateUrl(SplitSeries[2]);
                            Thread.Sleep(1000);

                            string PageFound = "";
                            var PageNotFoundNode = Utility.driver.FindElement(By.XPath("//div/div[2]/span"));
                            if(PageNotFoundNode != null)
                            {
                                PageFound = PageNotFoundNode.Text;
                            }
                            if (PageFound != "404")
                            {

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

                                Thread.Sleep(20000);
                                //Thread.Sleep(1000 * 60 * 2);

                                // Set the current directory path
                                DataDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\WithdrawnProduct\\" + categories[y];
                                string newFileName = Series + "_" + SubSeries + "_" + FullSeriesName + ".xls";
                                objUtility.Rename_MoveFile(newFileName, DataDirectory);
                                Console.WriteLine((z + 1) + " File out of " + lstSeries.Count + " Downloaded Successfully!\n");
                                Console.WriteLine("---------------------------------------------------------------------------------------------------------");
                            }
                            else
                            {
                                Console.WriteLine(SplitSeries[2] + " Not Found\n");
                            }
                        }

                        Console.WriteLine("Series " + Series + " of Product Type " + ProductType + " Completed Successfully");
                        Console.WriteLine("---------------------------------------------------------------------------------------------------------\n");
                        objUtility.NavigateUrl("https://psref.lenovo.com/WithdrawnProducts");

                    }
                }
            }
        }
    }
}
