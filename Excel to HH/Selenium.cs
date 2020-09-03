using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using WordPressPCL;
using WordPressPCL.Models;

namespace Excel_to_HH
{
    class Selenium
    {
        public static void removeDuplicates(IWebDriver driver, WebDriverWait wait)
        {
            foreach (Product product in Product.products)
            {
                product.change = false;
            }

            foreach (Product product1 in Product.products)
            {
                String name1;

                driver.Navigate().GoToUrl(product1.URL);

                try
                {
                    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//span [@id = 'productTitle']")));
                    name1 = driver.FindElement(By.XPath("//span [@id = 'productTitle']")).Text;
                }
                catch (Exception e)
                {
                    if(e is NoSuchElementException)
                    {
                        name1 = "";
                        product1.change = true;
                    }
                    else if(e is WebDriverTimeoutException)
                    {
                        name1 = "";
                        product1.change = true;
                    }
                    else
                    {
                        throw;
                    }
                }

                foreach (Product product2 in Product.products)
                {
                    String name2;

                    driver.Navigate().GoToUrl(product2.URL);

                    try
                    {
                        name2 = driver.FindElement(By.XPath("//span [@id = 'productTitle']")).Text;
                    }
                    catch (Exception e)
                    {
                        if (e is NoSuchElementException)
                        {
                            name2 = "";
                            product2.change = true;
                        }
                        else if (e is WebDriverTimeoutException)
                        {
                            name2 = "";
                            product2.change = true;
                        }
                        else
                        {
                            throw;
                        }
                    }

                    if(name1 == name2)
                    {
                        product2.change = true;
                    }
                }
            }

            restart:
            foreach (Product product in Product.products)
            {
                if(product.change)
                {
                    Product.products.Remove(product);
                    Wordpress.removeProduct(product).Wait();
                    goto restart;
                }
            }
        }
        public static String getLink(IWebDriver driver, WebDriverWait wait, Post post)
        {
            int totalPages;
            String tag = Wordpress.SendGetTag(Wordpress.GetTag(post));
            driver.FindElement(By.XPath("//div [@class = 'wp-menu-image dashicons-before dashicons-admin-post']")).Click();
            if (driver.FindElements(By.XPath("//span [@class = 'total-pages']")).Count > 0)
            {
                totalPages = int.Parse(driver.FindElement(By.XPath("//span [@class = 'total-pages']")).Text);
            }
            else
            {
                totalPages = 1;
            }

            for (int x = 1; x <= totalPages; x++)
            {
                List<IWebElement> elements = driver.FindElements(By.XPath("//td [@class = 'tags column-tags']")).ToList();
                foreach (IWebElement element in elements)
                {
                    if(element.Text == tag)
                    {
                        return driver.FindElement(By.XPath("(//span [@class = 'url'])[" + (elements.IndexOf(element)+1).ToString() + "]")).GetAttribute("innerHTML");
                    }
                }
                if(x!=totalPages)
                {
                    driver.FindElement(By.XPath("//a [@class = 'next-page button']")).Click();
                    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//td [@id = 'cb']")));
                }
            }
            MessageBox.Show("Link not found");
            return "Not Found";

        }


        public static void GetOldPostData(IWebDriver driver, WebDriverWait wait)
        {
            foreach (Product update in Product.updates)
            {
                driver.Navigate().GoToUrl(update.URL);
                try
                {
                    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//span [@id = 'productTitle']")));
                }
                catch (OpenQA.Selenium.WebDriverTimeoutException e)
                {
                    if (driver.FindElements(By.XPath("//img[@alt = 'Dogs of Amazon']")).Count > 0)
                    {
                        update.Price = -1;
                        update.Xprice = -1;
                        update.change = true;
                    }
                    else
                    {
                        throw;
                    }
                }

                //Grab Price
                if (driver.FindElements(By.XPath("//span [@id = 'priceblock_ourprice']")).Count > 0)
                {
                    update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_ourprice']");
                }
                else if (driver.FindElements(By.XPath("//span [@id = 'priceblock_saleprice']")).Count > 0)
                {
                    update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_saleprice']");
                }
                else if (driver.FindElements(By.XPath("//span [@id = 'priceblock_dealprice']")).Count > 0)
                {
                    update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_dealprice']");
                }
                else
                {
                    update.Price = -1;
                    update.change = true;
                }

                //Grab Xprice
                if (driver.FindElements(By.XPath("//span [@class = 'priceBlockStrikePriceString a-text-strike']")).Count > 0)
                {
                    update.Xprice = Formatting.ElementToInt(driver, "//span [@class = 'priceBlockStrikePriceString a-text-strike']");
                }
                else
                {
                    update.Xprice = -1;
                    update.change = true;
                }

            }

       
        }

        public static void goToPosts(IWebDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            driver.Navigate().GoToUrl("https://zed.exioite.com/wp-login.php");
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//input[@id='user_login']")));

            IWebElement username = driver.FindElement(By.XPath("//input[@id='user_login']"));
            IWebElement password = driver.FindElement(By.XPath("//input[@id='user_pass']"));

            Thread.Sleep(500);
            username.Clear();
            username.SendKeys(@"zaid@exioite.com");
            Thread.Sleep(500);
            password.Clear();
            password.SendKeys(@"*xuFKWOX@t8Oc$8fgALK4HLh");
            Thread.Sleep(500);

            driver.FindElement(By.XPath("//input[@id='wp-submit']")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div[contains(text(),'Posts')]")));

            driver.FindElement(By.XPath("//div[contains(text(),'Posts')]")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//a[contains(text(),'Add New')]")));

           
        }

        //TODO test if new addLink() method works
        public static void addLink(IWebDriver driver, String URL, String id)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            driver.Navigate().Refresh();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//a [@class='row-title']")));

            bool clicked = false;
            int totalPages;
            if(driver.FindElement(By.XPath("//span [@class='total-pages']")).Text == "")
            {
                totalPages = 1;
            }
            else
            {
                totalPages = int.Parse(driver.FindElement(By.XPath("//span [@class='total-pages']")).Text);
            }
            for (int x = 1; x <= totalPages; x++)
            {
                foreach (IWebElement element in driver.FindElements(By.XPath("//td [@class = 'tags column-tags']//a")))
                {
                    if (element.Text == id)
                    {
                        element.Click();
                        clicked = true;
                        break;
                    }
                }
                if (!clicked && x!=totalPages)
                {
                    driver.FindElement(By.XPath("//a [@class = 'next-page button']")).Click();
                }
            }

            if(!clicked)
            {
                MessageBox.Show("Post not found (Method: Selenium.AddLink())");
            }

            driver.FindElement(By.XPath("//a [@class = 'row-title']")).Click();
            Thread.Sleep(500);

            if (driver.FindElements(By.XPath("//div [@class = 'components-guide__container']")).Count > 0)
            {
                driver.FindElement(By.XPath("//button [@aria-label = 'Close dialog']")).Click();
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@data-label = 'Document']")));
            } 

            driver.FindElement(By.XPath("//button [@data-label = 'Document']")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [contains(text(),'Page Links To')]")));

            if (driver.FindElements(By.XPath("//label [contains(text(),'A custom URL')]")).Count == 0)
            {
                driver.FindElement(By.XPath("//button [contains(text(),'Page Links To')]")).Click();
            }

            driver.FindElement(By.XPath("//label [contains(text(),'A custom URL')]")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("// input  [@class = 'components-text-control__input']")));

            driver.FindElement(By.XPath("// input  [@class = 'components-text-control__input']")).Clear();

            Thread.Sleep(500);
            driver.FindElement(By.XPath("// input  [@class = 'components-text-control__input']")).SendKeys(URL);
            Thread.Sleep(500);

            driver.FindElement(By.XPath("//input [@data-testid = 'plt-newtab']")).Click();

            driver.FindElement(By.XPath("//button [contains(text(),'Update')]")).Click();
            Thread.Sleep(500);

            driver.FindElement(By.XPath("//a [@aria-label = 'View Posts']")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//h1 [contains(text(),'Posts')]")));
            
        }
        public static void WaitForPageLoad(WebDriverWait wait)
        {
            wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
        }

        //public static bool checkProducts(IWebDriver driver, Product product, Product update)
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        //    driver.Navigate().GoToUrl(update.URL);
        //    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//span [@id = 'productTitle']")));

        //    //Grab Price
        //    if (driver.FindElements(By.XPath("//span [@id = 'priceblock_ourprice']")).Count > 0)
        //    {
        //        update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_ourprice']");
        //    }
        //    else if (driver.FindElements(By.XPath("//span [@id = 'priceblock_saleprice']")).Count > 0)
        //    {
        //        update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_saleprice']");
        //    }
        //    else if (driver.FindElements(By.XPath("//span [@id = 'priceblock_dealprice']")).Count > 0)
        //    {
        //        update.Price = Formatting.ElementToInt(driver, "//span [@id = 'priceblock_dealprice']");
        //    }

        //    //Grab Xprice
        //    if (driver.FindElements(By.XPath("//span [@class = 'priceBlockStrikePriceString a-text-strike']")).Count > 0)
        //    {
        //        update.Xprice = Formatting.ElementToInt(driver, "//span [@class = 'priceBlockStrikePriceString a-text-strike']");
        //    }
        //}
    }
}
