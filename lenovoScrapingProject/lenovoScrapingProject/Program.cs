using System;

namespace lenovoScrapingProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Utility objUtility = new Utility();
            Scraper objScrap = new Scraper();
            WithdrawnProduct objWithdraw = new WithdrawnProduct();

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(" " + "\n---------------------------------------------------------------------------------------------------------");
            Console.WriteLine("\t\t\tWELCOME TO THE DASHBOARD OF LENOVO");
            Console.WriteLine("---------------------------------------------------------------------------------------------------------\n");
            Console.ResetColor();
            Console.WriteLine("\n---------------------------------------------------------------------------------------------------------");

            Utility.CheckNetConnection();
            if (Utility.networkAvailable == true)
            {
                Console.WriteLine("Internet Connection Available.");
                objUtility.Chrome_Instance_Removal();
                objUtility.ChromDriverServices();
                //objScrap.ScrapData();
                objWithdraw.WithdrawnData();


            }
        }
    }
}
