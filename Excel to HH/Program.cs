using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Drawing;
using WordPressPCL.Client;
using System.IO;
using System.Drawing.Imaging;
using System.Windows.Forms;
using WordPressPCL.Utility;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Support.UI;
using WordPressPCL.Models.Exceptions;
using System.Linq;
using WordPressPCL.Models;

namespace Excel_to_HH
{
    class Program
    {
        public static void Main(string[] args)
        {
            //initialize selenium
            IWebDriver driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            Wordpress.CleanMedia().Wait();

            //Grabs last week's posts straight from website
            Product.updates.Clear();
            Wordpress.ReadWebsite(driver, wait, Product.updates).Wait();
            Selenium.GetOldPostData(driver, wait);

            Selenium.goToPosts(driver);

            Wordpress.UpdatePosts(driver).Wait();

            Product.products.Clear();
            excel.ReadProducts();
            Formatting.correctCategories();

            Product.removeUpdates();

            Wordpress.AddPics(driver).Wait();
            foreach (Product product in Product.products)
            {
                Wordpress.CreatePost(driver, product).Wait();
            }

            Wordpress.removeDuplicates();
            Wordpress.CleanImagesFolder().Wait();
            excel.WriteHHPosts();

            driver.Close();
            driver.Quit();


        }


    }
}
