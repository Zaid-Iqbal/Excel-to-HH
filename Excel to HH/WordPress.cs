using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using Org.BouncyCastle.Asn1.IsisMtt.X509;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPressPCL;
using WordPressPCL.Client;
using WordPressPCL.Models;

namespace Excel_to_HH
{
    class Wordpress
    {
        public static List<Post> SendGetPosts(Task<List<Post>> posts)
        {
            posts.Wait();
            return posts.Result.ToList();
        }

        public static async Task<List<Post>> GetPosts()
        {
            var client = await GetClient();
            return client.Posts.GetAll().Result.ToList();
        }

        public static async Task CleanImagesFolder()
        {
            var client = await GetClient();
            List<String> goodID = new List<String>();

            foreach (MediaItem pic in client.Media.GetAll().Result)
            {
                goodID.Add(pic.Title.Rendered);
            }

            restart:
            foreach (String filePath in Directory.GetFiles(@"C:\Users\email\Desktop\Hardware Hub\images"))
            {
                if (!goodID.Contains(Formatting.getIDfromFile(filePath)))
                {
                    File.Delete(filePath);
                    goto restart;
                }
            }

        }
        public static async Task removeProduct(Product product)
        {
            var client = await GetClient();
            client.Posts.Delete(int.Parse(SendGetTag(GetTag(product))));
        }
        public static List<String> sendLinks(Task<List<String>> list)
        {
            return list.Result;
        }
        public static async Task<List<String>> getLinks()
        {
            var client = await GetClient();
            List<String> Links = new List<string>();
            foreach (Post post in client.Posts.GetAll().Result)
            {
                Links.Add(post.Link);
            }
            return Links;
        }
        public static async Task removeDuplicates()
        {
            var client = await GetClient();
            restart:
            foreach (Post post1 in client.Posts.GetAll().Result)
            {
                foreach (Post post2 in client.Posts.GetAll().Result)
                {
                    String title1 = post1.Title.Rendered;
                    String title2 = post2.Title.Rendered;
                    if (post1.Id != post2.Id && (title1 == title2 || Formatting.FuzzyMatch(title1,title2) || Formatting.FuzzyMatch(title2, title1)))
                    {
                        await client.Media.Delete(post2.FeaturedMedia.Value);
                        await client.Tags.Delete(post2.Tags[0]);
                        await client.Posts.Delete(post2.Id);
                        goto restart;
                    }
                }
            }
        }

        public static String SendGetTag(Task<String> str)
        {
            return str.Result;
        }

        public static async Task<String> GetTag(Product product)
        {
            var client = await GetClient();
            List<Tag> tags = client.Tags.GetAll().Result.ToList();
            if (await client.IsValidJWToken())
            {
                foreach (Tag tag in tags)
                {
                    if (product.ID == tag.Name)
                    {
                        return tag.Id.ToString();
                    }
                }

                return "Tag Not Found";
            }
            return "client not approved";
        }

        public static async Task<String> GetTag(Post post)
        {
            var client = await GetClient();
            List<Tag> tags = client.Tags.GetAll().Result.ToList();
            if (await client.IsValidJWToken())
            {
                foreach (Tag tag in tags)
                {
                    if (post.Tags[0] == tag.Id)
                    {
                        return tag.Name;
                    }
                }
                //for (int x = 0; x < tags.Count; x++)
                //{
                //    if(post.Tags[0] == tags[x].Id)
                //    {
                //        Selenium.tag = tags[x].Name;
                //        goto end;
                //    }
                //    count++;
                //}
                return "Tag Not Found";
            }
            return "client not approved";
        }
        public static async Task CleanMedia()
        {
            var client = await GetClient();
            if(await client.IsValidJWToken())
            {
                List<int> used = new List<int>();
                foreach (Post post in client.Posts.GetAll().Result.ToList())
                {
                    foreach (MediaItem pic in client.Media.GetAll().Result.ToList())
                    {
                        if(post.FeaturedMedia == pic.Id)
                        {
                            used.Add(pic.Id);
                            break;
                        }
                    }
                }
                foreach (MediaItem pic in client.Media.GetAll().Result.ToList())
                {
                    if (!used.Contains(pic.Id))
                    {
                        await client.Media.Delete(pic.Id);
                    }
                }
            }
            
        }


        public static async Task addTestPosts(int amount)
        {
            WordPressClient client = await GetClient();
            for (int x = 1; x <= amount; x++)
            {
                Post post = new Post
                {
                    Title = new Title(Formatting.RandomString(1))
                };
                client.Posts.Create(post).Wait();
            }
        }


        private static async Task<WordPressClient> GetClient()
        {
            //JWT authentication
            var client = new WordPressClient("https://zed.exioite.com/wp-json/");
            client.AuthMethod = AuthMethod.JWT;
            await client.RequestJWToken("EMAIL", "PASSWORD");
            return client;
        }
        public static async Task ReadWebsite(IWebDriver driver, WebDriverWait wait, List<Product> list)
        {
            WordPressClient client = await GetClient();
            if (await client.IsValidJWToken())
            {
                foreach (Post post in client.Posts.GetAll().Result)
                {
                    list.Add(new Product(
                        post.Title.Rendered,
                        post.Categories[0],
                        post.Link,
                        post.Tags[0]
                        ));
                }
                tagIDtoID(Product.updates).Wait();
            }
        }

        //public static async Task ReadLastWeek2(IWebDriver driver, WebDriverWait wait)
        //{
        //    WordPressClient client = await GetClient();
        //    if (await client.IsValidJWToken())
        //    {
        //        foreach (Post post in client.Posts.GetAll().Result)
        //        {
        //            Product.updates.Add(new Product(
        //                post.Title.Rendered,
        //                post.Categories[0],
        //                Selenium.getLink(driver, wait, post),
        //                post.Tags[0]
        //                ));
        //        }
        //        tagIDtoID(Product.updates).Wait();
        //    }
        //}

        public static async Task UpdatePosts(IWebDriver driver)
        {
            WordPressClient client = await GetClient();

            foreach (Post post in client.Posts.GetAll().Result)
            {
                foreach (Product update in Product.updates)
                {

                    if (post.Tags[0] == update.tagID)
                    {
                        //if update is no longer on sale, delete from wordpress
                        if (update.Price == -1 || update.Xprice == -1)
                        {
                            await client.Posts.Delete(post.Id);
                            await client.Tags.Delete(post.Tags[0]);
                            await client.Media.Delete(post.FeaturedMedia.Value);
                        }
                        else
                        {
                            //Otherwise update with the new prices
                            Post updatePost = new Post
                            {
                                Id = post.Id,
                                Content = new Content("$" + update.Xprice + "-->" + "$" + update.Price),
                            };
                            await client.Posts.Update(updatePost);
                        }
                    }
                }
            }

            
        }

        public static async Task CreatePost(IWebDriver driver, Product product)
        {
            int[] catID = new int[1];
            catID[0] = product.CatID;

            WordPressClient client = await GetClient();
            if (await client.IsValidJWToken())
            {
                int mediaID = -1;
                bool found = false;
                foreach (MediaItem item in client.Media.GetAll().Result)
                {
                    if (item.Slug == product.ID)
                    {
                        mediaID = item.Id;
                        Tag newTag = new Tag
                        {
                            Name = product.ID,
                            Slug = product.ID,
                        };
                        await client.Tags.Create(newTag);

                        int[] tagID = new int[1];

                        foreach (Tag tag in client.Tags.GetAll().Result)
                        {
                            if (tag.Name == product.ID)
                            {
                                tagID[0] = tag.Id;
                            }
                        }                        
                        Post post = new Post
                        {
                            Title = new Title(product.Name),
                            Content = new Content("$" + product.Xprice + "-->" + "$" +product.Price),
                            Categories = catID,
                            FeaturedMedia = mediaID,
                            Tags = tagID
                        };
                        await client.Posts.Create(post);
                        Selenium.addLink(driver, product.URL, product.ID);
                        found = true;
                    }
                }
                if(found==false)
                {
                    MessageBox.Show("Unable to find image and post");
                }
            }
            

        }

        public static async Task tagIDtoID(List<Product> products)
        {
            WordPressClient client = await GetClient();
            if (await client.IsValidJWToken())
            {
                foreach (Product product in products)
                {
                    foreach (Tag tag in client.Tags.GetAll().Result)
                    {
                        if(tag.Id == product.tagID)
                        {
                            product.ID = tag.Name;
                        }
                    }
                }
                foreach (Product update in Product.updates)
                {
                    if(update.ID == null)
                    {
                        MessageBox.Show("Wordpress.tagIDtoID(); failed to find and update all tags");
                    }
                }
            }
                
        }

        public static async Task TestPost()
        {
            int[] catID = new int[1];
            catID[0] = 7;

            WordPressClient client = await GetClient();
            if (await client.IsValidJWToken())
            {
                int mediaID = 0;
                foreach (MediaItem item in client.Media.GetAll().Result)
                {
                    if (item.Slug == "logo")
                    {
                        mediaID = item.Id;

                        Tag newTag = new Tag
                        {
                            Name = "test",
                            Slug = "test",
                        };
                        await client.Tags.Create(newTag);

                        int[] tagID = new int[1];
                        foreach (Tag tag in client.Tags.GetAll().Result)
                        {
                            if (tag.Name == "test")
                            {
                                tagID[0] = tag.Id;
                            }
                        }
                        Post post = new Post
                        {
                            Title = new Title("Test Post"),
                            Content = new Content("Test Content"),
                            Categories = catID,
                            FeaturedMedia = mediaID,
                            Tags = tagID
                        };
                        await client.Posts.Create(post);
                    }
                }
            }

        }

        public static async Task AddPics(IWebDriver driver)
        {
            var client = await GetClient();
            try
            {
                foreach(Product product in Product.products)
                {
                    if (await client.IsValidJWToken())
                    {
                        Image pic = Image.FromFile(@"C:\Users\email\Desktop\Hardware Hub\images\" + product.ID + ".png");
                        await client.Media.Create(Formatting.ToStream(pic, ImageFormat.Png), product.ID, "image/png");
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Wordpress.AddPics() Error: " + e.Message);
            }
        }



        public static async Task TestTag()
        {
            try
            {
                WordPressClient client = await GetClient();
                if (await client.IsValidJWToken())
                {
                    foreach (Tag tag in client.Tags.GetAll().Result)
                    {
                        MessageBox.Show(tag.Name);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error:" + e.Message);
            }
        }

    }
}
