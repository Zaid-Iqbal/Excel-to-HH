using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Excel_to_HH
{
    class Product
    {

        public static List<Product> products = new List<Product>();
        public static List<Product> updates = new List<Product>();
        public static List<Product> removes = new List<Product>();

        public String Name 
        {
            get;
            set;
        }
        public int Price
        {
            get;
            set;
        }
        public int Xprice
        {
            get;
            set;
        }
        public String Category
        {
            get;
            set;
        }
        public int CatID
        {
            get;
            set;
        }
        public String ID
        {
            get;
            set;
        }
        public String URL
        {
            get;
            set;
        }

        public int WPID
        {
            get;
            set;
        }
        public int tagID
        {
            get;
            set;
        }
        public bool change
        {
            get;
            set;
        }


        //used with products.xlsx
        public Product(String name, int price, int xprice, String category, String url, String id)
        {
            Name = name;
            Price = price;
            Xprice = xprice;
            Category = category;
            ID = id;
            URL = url;
            change = false;
        }

        //used mostly with updates
        public Product(String name, int category, String url, int id)
        {
            Name = name;
            CatID = category;
            tagID = id;
            URL = url;
            change = false;
        }

        public static void removeUpdates()
        {
            restart:
            foreach (Product update in updates)
            {
                foreach (Product product in products)
                {
                    if(update.Name == product.Name || update.URL == product.URL || update.ID == product.ID || (update.Price == product.Price && update.Xprice == product.Xprice) ||Formatting.FuzzyMatch(product.Name,update.Name))
                    {
                        products.Remove(product);
                        goto restart;
                    }

                }
            }
            
        }
    }
}
