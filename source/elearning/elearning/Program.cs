using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mime;
using System.IO;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using HeyRed.Mime;

namespace elearning
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Login ID:");
            var id = Console.ReadLine();
            Console.Write("Password:");
            var password = Console.ReadLine();
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--log-level=3");
            var driver = new ChromeDriver(options);
            try
            {
                driver.Navigate().GoToUrl("https://elearn.cuhk.edu.hk");
                driver.FindElementByName("username").SendKeys(id);
                driver.FindElementByName("password").SendKeys(password);
                driver.FindElementByXPath("/html/body/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/form/table/tbody/tr[6]/td/input[1]").Submit();
                IWebElement table = new WebDriverWait(driver, TimeSpan.FromMilliseconds(20000)).Until(ExpectedConditions.ElementExists(By.ClassName("portletList-img")));
                var course_li = driver.FindElementsByCssSelector(".portletList-img > li");
                for (int i = 1; i <= course_li.Count(); i++)
                {
                    IWebElement wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(20000)).Until(ExpectedConditions.ElementExists(By.ClassName("portletList-img")));
                    var course_link = driver.FindElementByXPath("//*[@id='_4_1termCourses_noterm']/ul/li[" + i + "]/a");
                    string course_text = course_link.Text;
                    try
                    {
                        course_link.Click();
                        IList<IWebElement> elements = driver.FindElements(By.CssSelector("#courseMenuPalette_contents > li > a[href*=webapps]"));
                        List<string> xpaths = new List<string>();
                        foreach (var element in elements)
                        {
                            xpaths.Add(GetElementXPath(driver, element));
                        }
                        List<string> except_text = new List<string>(new string[] { "Notifications", "Announcements", "Discussion Board", "Notifications", "Email", "Groups", "My Grades" });
                        foreach (var element in xpaths)
                        {

                            if (!except_text.Contains(driver.FindElementByXPath(element).Text))
                            {
                                driver.FindElementByXPath(element).Click();
                                download_file(element, driver, driver.FindElementByXPath(element).Text);
                            }
                        }

                        driver.Navigate().GoToUrl("https://elearn.cuhk.edu.hk");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Some file in " + course_text + "cannot be dowloaded.  Please check.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                driver.Quit();
            }
            driver.Quit();
            Console.WriteLine("Press ENTER to close console......");
            ConsoleKeyInfo keyInfo = Console.ReadKey();
            while (keyInfo.Key != ConsoleKey.Enter)
                keyInfo = Console.ReadKey();
        }
        private static bool IsElementPresent(IWebDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        private static string cookieString(IWebDriver driver)
        {
            var cookies = driver.Manage().Cookies.AllCookies;
            return string.Join("; ", cookies.Select(c => string.Format("{0}={1}", c.Name, c.Value)));
        }
        private static void download_file(string xpath, ChromeDriver driver, string type)
        {
            IList<IWebElement> elements2 = driver.FindElements(By.CssSelector(".item > h3 > a[href*=bbcswebdav]"));

            string header = driver.FindElementById("courseMenu_link").Text;
            string course_code = GetSafeFilename(header.Replace("Old-", ""));
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Course\" + course_code + "\\";
            download_file2(course_code, driver, path + type);

            IList<IWebElement> elements = driver.FindElements(By.CssSelector(".item > h3 > a[href*=webapps]"));
            List<string> xpaths = new List<string>();
            foreach (var element in elements)
            {
                xpaths.Add(GetElementXPath(driver, element));
            }
            foreach (string element in xpaths)
            {
                IWebElement table = new WebDriverWait(driver, TimeSpan.FromMilliseconds(20000)).Until(ExpectedConditions.ElementExists(By.XPath(element)));
                string temp = driver.FindElement(By.XPath(element)).Text;
                if (driver.FindElement(By.XPath(element)).GetAttribute("href").Contains("launch_in_new=true"))
                {
                    download_file3(course_code, driver, path + type, driver.FindElement(By.XPath(element)));
                    continue;
                }
                if (driver.FindElement(By.XPath(element)).GetAttribute("href").Contains("launchPackage"))
                {
                    continue;
                }
                if (isAttribtuePresent(driver.FindElement(By.XPath(element)), "target"))
                {
                    if (driver.FindElement(By.XPath(element)).GetAttribute("target") == "_blank")
                    {
                        continue;
                    }
                }
                driver.FindElement(By.XPath(element)).Click();
                download_file2(course_code, driver, path + type + "\\" + GetSafeFilename(temp));
                driver.FindElement(By.XPath(xpath)).Click();
            }



        }
        private static void download_file2(string course, ChromeDriver driver, string path)
        {
            //string header = driver.FindElementById("courseMenu_link").Text;
            //string course_code = header.Replace("Old-", "").Split('-')[1].Split(' ')[0];
            //path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Course\" + course_code + "\\" + path;
            System.IO.Directory.CreateDirectory(path);
            if (IsElementPresent(driver, By.Id("downloadPanelButton")))
            {
                string outline = driver.FindElementById("downloadPanelButton").GetAttribute("href");
                using (WebClient client = new WebClient())
                {

                    client.Headers[HttpRequestHeader.Cookie] = cookieString(driver);
                    var data = client.DownloadData(outline);
                    var extension = MimeTypesMap.GetExtension(client.ResponseHeaders["Content-Type"]);
                    string fileName = GetSafeFilename(driver.FindElementById("downloadPanelFileName").Text);
                    System.IO.Directory.CreateDirectory(path);
                    File.WriteAllBytes(path + "\\" + System.IO.Path.GetFileNameWithoutExtension(fileName) + "." + extension, data);
                }
            }
            if (IsElementPresent(driver, By.Id("content_listContainer")))
            {
                IList<IWebElement> elements = driver.FindElements(By.CssSelector("a[href*=bbcswebdav]"));
                foreach (var element in elements)
                {
                    download_file3(course, driver, path, element);
                }
            }
        }
        private static void download_file3(string course, ChromeDriver driver, string path, IWebElement element)
        {
            string outline = element.GetAttribute("href");
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.Cookie] = cookieString(driver);
                var data = client.DownloadData(outline);
                var extension = MimeTypesMap.GetExtension(client.ResponseHeaders["Content-Type"]);
                string fileName = GetSafeFilename(element.Text);

                System.IO.Directory.CreateDirectory(path);
                string file_name = path + "\\" + System.IO.Path.GetFileNameWithoutExtension(fileName) + "." + extension;
                if (file_name.Length > 259)
                {
                    Console.WriteLine("Warning: Course:" + course + " file:" + fileName + ". Please download manually");
                    return;
                }
                File.WriteAllBytes(file_name, data);
            }
        }
        public static string GetSafeFilename(string filename)
        {
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }
        public static String GetElementXPath(IWebDriver driver, IWebElement element)
        {
            String javaScript = "function getElementXPath(elt){" +
                                    "var path = \"\";" +
                                    "for (; elt && elt.nodeType == 1; elt = elt.parentNode){" +
                                        "idx = getElementIdx(elt);" +
                                        "xname = elt.tagName;" +
                                        "if (idx > 1){" +
                                            "xname += \"[\" + idx + \"]\";" +
                                        "}" +
                                        "path = \"/\" + xname + path;" +
                                    "}" +
                                    "return path;" +
                                "}" +
                                "function getElementIdx(elt){" +
                                    "var count = 1;" +
                                    "for (var sib = elt.previousSibling; sib ; sib = sib.previousSibling){" +
                                        "if(sib.nodeType == 1 && sib.tagName == elt.tagName){" +
                                            "count++;" +
                                        "}" +
                                    "}" +
                                    "return count;" +
                                "}" +
                                "return getElementXPath(arguments[0]).toLowerCase();";
            return (String)((IJavaScriptExecutor)driver).ExecuteScript(javaScript, element);
        }
        private static Boolean isAttribtuePresent(IWebElement element, String attribute)
        {
            Boolean result = false;
            try
            {
                String value = element.GetAttribute(attribute);
                if (value != null)
                {
                    result = true;
                }
            }
            catch (Exception e) { }

            return result;
        }
    }
}