using System;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Data;
using System.Linq;
using System.Diagnostics;
using OpenQA.Selenium.Interactions;
using System.Threading;

namespace lenovoScrapingProject
{
    class Utility
    {
        public static ChromeDriver driver;
        public static DataTable dataTbl = new DataTable();
        public static string csvPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Data.csv";
        public static bool networkAvailable;

        public void ChromDriverServices()
        {
            var options = new ChromeOptions();
            options.AddArguments("disable-infobars");
            //options.AddArgument("user-data-dir=C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data");
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, options);
            driver.Manage().Window.Maximize();
        }

        public void NavigateUrl(string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);
                Thread.Sleep(2000);
            }
            catch
            {
                Thread.Sleep(30000);
                NavigateUrl(url);
            }
        }

        public void Rename_MoveFile(string newFileName, string directory)
        {

            // Set the download directory path
            string downloadDirectory = @"C:\Users\Administrator\Downloads\";
          
            // Get all files in the download directory with the desired file extension
            string[] files = Directory.GetFiles(downloadDirectory, "*" + ".xls");
            var status = "";
            for (int z = 0; z < files.Length; z++)
            {
                string sourceFilePath = files[z];
                var SeriesName = driver.Url.ToString().Split("/");
                var SeriesName1 = SeriesName[SeriesName.Length - 1].Replace("_", "");
                var fileName = files[z].Replace("Built_in", "Builtin").Replace("plus", "and").Replace("_", "").ToString();

                if (files[z].Replace("_", "").Contains(SeriesName1) || fileName.Contains(SeriesName1))
                {
                    string destinationPath = Path.Combine(directory, newFileName.Replace("/", "__").Replace(@"""", @""));

                    try
                    {
                        // Move the file to the current directory
                        File.Move(sourceFilePath, destinationPath);
                        status = "moved";
                        Console.WriteLine("File moved successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error occurred while moving the file: " + ex.Message);
                    }
                    break;
                }
            }

            if(status!="moved")
            {
                    Console.WriteLine("No XLSX files found in the download directory.");
                    Thread.Sleep(1000 * 60 * 3);
                    Rename_MoveFile(newFileName, directory);
            }
        }
        


        public static void WaitForElement(string xpath)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(3));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(xpath)));
            }
            catch
            {

            }
        }



        public void ScrollPageUptoElement(IWebElement element)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(element);
            actions.Perform();
        }



        public void Chrome_Instance_Removal()
        {
            try
            {
                Process[] chromeDriverProcesses = Process.GetProcessesByName("chrome");
                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }
            }
            catch { }
        }

        public static void CheckNetConnection()
        {
            networkAvailable = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
        }
    }
}
