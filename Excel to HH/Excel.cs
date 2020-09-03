using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPressPCL.Models;

namespace Excel_to_HH
{
    class excel
    {

        public static void WriteHHPosts()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws = wb.CreateSheet();

            IRow header = ws.CreateRow(0);
            header.CreateCell(0).SetCellValue("Name");
            header.CreateCell(1).SetCellValue("Price");
            header.CreateCell(2).SetCellValue("Old Price");
            header.CreateCell(3).SetCellValue("Difference");
            header.CreateCell(4).SetCellValue("Category");
            header.CreateCell(5).SetCellValue("ID");
            header.CreateCell(6).SetCellValue("URL");

            int rowcount = 1;
            foreach (Post post in Wordpress.SendGetPosts(Wordpress.GetPosts()))
            {
                IRow row = ws.CreateRow(rowcount);

                row.CreateCell(0).SetCellValue(post.Title.Rendered);
                row.CreateCell(1).SetCellValue(int.Parse(Formatting.getPrice(post.Content.Rendered)));
                row.CreateCell(2).SetCellValue(int.Parse(Formatting.getXprice(post.Content.Rendered)));
                row.CreateCell(3).SetCellValue(int.Parse(Formatting.getXprice(post.Content.Rendered)) - int.Parse(Formatting.getPrice(post.Content.Rendered)));
                row.CreateCell(4).SetCellValue(Formatting.CatIDtoCategory(post.Categories[0]));
                row.CreateCell(5).SetCellValue(Wordpress.SendGetTag(Wordpress.GetTag(post)));
                row.CreateCell(6).SetCellValue(post.Link);

                rowcount++;
            }

            wb.Write(new FileStream(@"C:\Users\email\Desktop\Hardware Hub\HHPosts.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            wb.Close();
        }

        public static void testExcel()
        {
            IWorkbook wb = new XSSFWorkbook();
            wb.CreateSheet();
            ISheet ws = wb.GetSheetAt(0);

            ws.CreateRow(0).CreateCell(0).SetCellValue("test");

            Stream stream = new FileStream("C:\\Users\\email\\Desktop\\Hardware Hub\\test.xlsx", FileMode.Open);
            wb.Write(stream);
            wb.Close();
        }


        public static void ReadProducts()
        {
            IWorkbook wb = new XSSFWorkbook("C:\\Users\\email\\Desktop\\Hardware Hub\\products.xlsx");
            ISheet ws = wb.GetSheetAt(0);

            for (int x = 1; x <= ws.LastRowNum; x++)
            {
                IRow row = ws.GetRow(x);

                if(row.Cells[0].ToString() != "Stop")
                {
                    String name = row.Cells[0].ToString();
                    String price = row.Cells[1].ToString();
                    String xprice = row.Cells[2].ToString();
                    String category = row.Cells[4].ToString();
                    String id = row.Cells[6].ToString();
                    String url;
                    if (id.Substring(0,4) == "http")
                    {
                        id = row.Cells[5].ToString();
                        url = row.Cells[6].ToString();
                    }
                    else
                    {
                        url = row.Cells[7].ToString();
                    }

                    Product.products.Add(new Product
                    (
                        name,
                        int.Parse(price),
                        int.Parse(xprice),
                        category,
                        url,
                        id
                    ));
                }
                else
                {
                    break;
                } 
            }
            wb.Close();

        }

    }
}
