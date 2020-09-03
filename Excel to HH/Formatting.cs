using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPressPCL.Models;

namespace Excel_to_HH
{
    class Formatting
    {
        public static String sample = @"< p >$50 &#8211;&gt;$28</p>\n";

        public static String CatIDtoCategory(int CatID)
        {
            if (CatID == 7)
            {
                return "Audio";
            }
            else if (CatID == 9)
            {
                return "Computers and Tablets";
            }
            else if (CatID == 13)
            {
                return "Displays";
            }
            else if (CatID == 5)
            {
                return "Gaming";
            }
            else if (CatID == 10)
            {
                return "Misc";
            }
            else if (CatID == 6)
            {
                return "PC Parts";
            }
            else if (CatID == 8)
            {
                return "Phones";
            }
            else
            {
                MessageBox.Show("Formatting.CatIDtoCategory() Error: Categroy not found");
                return "";
            }

        }

        public static int GetNthIndex(string s, char t, int n)
        {
            int count = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == t)
                {
                    count++;
                    if (count == n)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        public static String getIDfromFile(String file)
        {
            int slashIndex = GetNthIndex(file, '\\', 6) + 1;
            return file.Substring(slashIndex,file.IndexOf('.')-slashIndex);
        }

        public static String getXprice(String str)
        {
            int dollar = str.IndexOf("$");
            str = str.Substring(dollar + 1);
            for (int x = 1; x < str.Length - 1; x++)
            {
                foreach (char c in str.Substring(0, x))
                {
                    if (c != '1' && c != '2' && c != '3' && c != '4' && c != '5' && c != '6' && c != '7' && c != '8' && c != '9' && c != '0')
                    {
                        return str.Substring(0, x - 1);
                    }
                }
            }
            return "not found";
        }

        public static String getPrice(String str)
        {
            int dollar = str.IndexOf("$");
            str = str.Substring(dollar + 1);
            dollar = str.IndexOf("$");
            str = str.Substring(dollar + 1);
            for (int x = 1; x < str.Length-1; x++)
            {
                foreach (char c in str.Substring(0,x))
                {
                    if (c != '1' && c != '2' && c != '3' && c != '4' && c != '5' && c != '6' && c != '7' && c != '8' && c != '9' && c != '0')
                    {
                        return str.Substring(0,x-1);
                    }
                }
            }
            return "not found";
        }
        public static string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static int IDtoNum(String str)
        {
            String newstr = Regex.Replace(str, "[^.0-9]", "");
            try
            {
                return int.Parse(newstr);
            }
            catch (Exception)
            {
                return int.Parse(newstr.Substring(0, 8));
            }
        }

        

        public static int ElementToInt(IWebDriver driver, String xpath)
        {
            return (int)Math.Round(Convert.ToDouble((driver.FindElement(By.XPath(xpath)).Text).Replace("$","")));
        }

        public static int ElementToInt(IWebElement element)
        {
            return (int)Math.Round(Convert.ToDouble((element.Text).Replace("$", "")));
        }

        public static void correctCategories()
        {
            foreach (Product product in Product.products)
            {
                if (product.Category == "Laptops")
                {
                    product.CatID = 9;
                }
                else if (product.Category == "Desktops")
                {
                    product.CatID = 6;
                }
                else if (product.Category == "Monitors")
                {
                    product.CatID = 13;
                }
                else if (product.Category == "Networking")
                {
                    product.CatID = 10;
                }
                else if (product.Category == "Computer Components")
                {
                    product.CatID = 6;
                }
                else if (product.Category == "Storage")
                {
                    product.CatID = 6;
                }
                else if (product.Category == "TV & Video")
                {
                    product.CatID = 13;
                }
                else if (product.Category == "Speakers")
                {
                    product.CatID = 7;
                }
                else if (product.Category == "Headphones")
                {
                    product.CatID = 7;
                }
                else if (product.Category == "Bluetooth Earbuds")
                {
                    product.CatID = 7;
                }
                else if (product.Category == "Phones")
                {
                    product.CatID = 8;
                }
                else if (product.Category == "Misc")
                {
                    product.CatID = 10;
                }
                else if (product.Category == "Audio")
                {
                    product.CatID = 7;
                }
                else if (product.Category == "Computers & Tablets")
                {
                    product.CatID = 9;
                }
                else if (product.Category == "Displays")
                {
                    product.CatID = 13;
                }
                else if (product.Category == "Gaming")
                {
                    product.CatID = 5;
                }
                else if (product.Category == "PC parts")
                {
                    product.CatID = 6;
                }
                else if (product.Category == "PC Parts")
                {
                    product.CatID = 6;
                }
                else
                {
                    product.CatID = -1;
                }
            }
        }

        public static Stream ToStream(Image image, ImageFormat format)
        {
            var stream = new System.IO.MemoryStream();
            image.Save(stream, format);
            stream.Position = 0;
            return stream;
        }

        /// <summary>
        /// Does a fuzzy search for a pattern within a string.
        /// </summary>
        /// <param name="stringToSearch">The string to search for the pattern in.</param>
        /// <param name="pattern">The pattern to search for in the string.</param>
        /// <returns>true if each character in pattern is found sequentially within stringToSearch; otherwise, false.</returns>
        public static bool FuzzyMatch(string stringToSearch, string pattern)
        {
            var patternIdx = 0;
            var strIdx = 0;
            var patternLength = pattern.Length;
            var strLength = stringToSearch.Length;

            while (patternIdx != patternLength && strIdx != strLength)
            {
                if (char.ToLower(pattern[patternIdx]) == char.ToLower(stringToSearch[strIdx]))
                    ++patternIdx;
                ++strIdx;
            }

            return patternLength != 0 && strLength != 0 && patternIdx == patternLength;
        }

        /// <summary>
        /// Does a fuzzy search for a pattern within a string, and gives the search a score on how well it matched.
        /// </summary>
        /// <param name="stringToSearch">The string to search for the pattern in.</param>
        /// <param name="pattern">The pattern to search for in the string.</param>
        /// <param name="outScore">The score which this search received, if a match was found.</param>
        /// <returns>true if each character in pattern is found sequentially within stringToSearch; otherwise, false.</returns>
        public static bool FuzzyMatch(string stringToSearch, string pattern, out int outScore)
        {
            // Score consts
            const int adjacencyBonus = 5;               // bonus for adjacent matches
            const int separatorBonus = 10;              // bonus if match occurs after a separator
            const int camelBonus = 10;                  // bonus if match is uppercase and prev is lower

            const int leadingLetterPenalty = -3;        // penalty applied for every letter in stringToSearch before the first match
            const int maxLeadingLetterPenalty = -9;     // maximum penalty for leading letters
            const int unmatchedLetterPenalty = -1;      // penalty for every letter that doesn't matter


            // Loop variables
            var score = 0;
            var patternIdx = 0;
            var patternLength = pattern.Length;
            var strIdx = 0;
            var strLength = stringToSearch.Length;
            var prevMatched = false;
            var prevLower = false;
            var prevSeparator = true;                   // true if first letter match gets separator bonus

            // Use "best" matched letter if multiple string letters match the pattern
            char? bestLetter = null;
            char? bestLower = null;
            int? bestLetterIdx = null;
            var bestLetterScore = 0;

            var matchedIndices = new List<int>();

            // Loop over strings
            while (strIdx != strLength)
            {
                var patternChar = patternIdx != patternLength ? pattern[patternIdx] as char? : null;
                var strChar = stringToSearch[strIdx];

                var patternLower = patternChar != null ? char.ToLower((char)patternChar) as char? : null;
                var strLower = char.ToLower(strChar);
                var strUpper = char.ToUpper(strChar);

                var nextMatch = patternChar != null && patternLower == strLower;
                var rematch = bestLetter != null && bestLower == strLower;

                var advanced = nextMatch && bestLetter != null;
                var patternRepeat = bestLetter != null && patternChar != null && bestLower == patternLower;
                if (advanced || patternRepeat)
                {
                    score += bestLetterScore;
                    matchedIndices.Add((int)bestLetterIdx);
                    bestLetter = null;
                    bestLower = null;
                    bestLetterIdx = null;
                    bestLetterScore = 0;
                }

                if (nextMatch || rematch)
                {
                    var newScore = 0;

                    // Apply penalty for each letter before the first pattern match
                    // Note: Math.Max because penalties are negative values. So max is smallest penalty.
                    if (patternIdx == 0)
                    {
                        var penalty = Math.Max(strIdx * leadingLetterPenalty, maxLeadingLetterPenalty);
                        score += penalty;
                    }

                    // Apply bonus for consecutive bonuses
                    if (prevMatched)
                        newScore += adjacencyBonus;

                    // Apply bonus for matches after a separator
                    if (prevSeparator)
                        newScore += separatorBonus;

                    // Apply bonus across camel case boundaries. Includes "clever" isLetter check.
                    if (prevLower && strChar == strUpper && strLower != strUpper)
                        newScore += camelBonus;

                    // Update pattern index IF the next pattern letter was matched
                    if (nextMatch)
                        ++patternIdx;

                    // Update best letter in stringToSearch which may be for a "next" letter or a "rematch"
                    if (newScore >= bestLetterScore)
                    {
                        // Apply penalty for now skipped letter
                        if (bestLetter != null)
                            score += unmatchedLetterPenalty;

                        bestLetter = strChar;
                        bestLower = char.ToLower((char)bestLetter);
                        bestLetterIdx = strIdx;
                        bestLetterScore = newScore;
                    }

                    prevMatched = true;
                }
                else
                {
                    score += unmatchedLetterPenalty;
                    prevMatched = false;
                }

                // Includes "clever" isLetter check.
                prevLower = strChar == strLower && strLower != strUpper;
                prevSeparator = strChar == '_' || strChar == ' ';

                ++strIdx;
            }

            // Apply score for last match
            if (bestLetter != null)
            {
                score += bestLetterScore;
                matchedIndices.Add((int)bestLetterIdx);
            }

            outScore = score;
            return patternIdx == patternLength;
        }
    }
}
