using System;
using System.Collections;
using System.Reflection;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Net.WebSockets;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
namespace Excel_to_HH
{
    class excel
    {
        static Excel.Application xlApp = new Excel.Application();
        static Excel.Workbook xlwb = xlApp.Workbooks.Open(@"C:\Users\email\Desktop\Hardware Hub\products.xlsx");
        static Excel.Worksheet xlws = (Excel.Worksheet)xlwb.Sheets[1];
        //static Excel.Range range = xlws.Range("A1:");

        public excel()
        {
            //Excel.Application app = xlApp;
            //Excel.Workbook wb = xlwb;
            //Excel.Worksheet ws = (Excel.Worksheet)xlsheet;
        }

        public static List<Product> writeProducts(List<Product> products)
        {
            List<Image> images = new List<Image>();
            Excel.Pictures pics = xlws.Pictures(Missing.Value) as Excel.Pictures;
            for (int x = 1; x <= pics.Count; x++)
            {
                pics.Item(x).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
            }
            foreach (Excel.Range row in xlws.Rows)
            {
                if (row.Cells[1] != null)
                {
                    products.Add(new Product(
                        row.Cells[1].ToString(),
                        int.Parse(row.Cells[2].ToString()),
                        int.Parse(row.Cells[3].ToString()),
                        row.Cells[5].ToString(),
                        null,
                        row.Cells[7].ToString()
                    ));
                }
            }
        }

    }
}
