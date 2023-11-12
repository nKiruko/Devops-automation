using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Reflection;
using System.Threading;
using OpenQA.Selenium.Safari;
using System.Collections.Generic;
using System.Web;
using System.Linq;
using CsvHelper;
using Newtonsoft.Json;
using CsvHelper.Configuration;
using System.Globalization;

namespace Webscraping
{
    public class ScrapingTest
    {
        String test_url_1 = "https://www.youtube.com/";
        String test_url_2 = "https://www.ictjob.be/";
        String test_url_3 = "https://raider.io";
        public IWebDriver driver;

        /* LambdaTest Credentials and Grid URL */
        String username = "nick.bulen";
        String accesskey = "aBs59Xj3d2wh7CMfg1cHzvKdWcqzZ3XhuL6NQIjISVJDMPZg9k";
        String gridURL = "@hub.lambdatest.com/wd/hub";

        [SetUp]
        public void start_Browser()
        {
            /*Local Selenium WebDriver */
            driver = new ChromeDriver();
            ChromeOptions capabilities = new ChromeOptions();
             capabilities.BrowserVersion = "119.0";
             Dictionary<string, object> ltOptions = new Dictionary<string, object>();
             ltOptions.Add("username", username);
             ltOptions.Add("accessKey", accesskey);
             ltOptions.Add("platformName", "Windows 10");
             ltOptions.Add("project", "Webscraper");
             capabilities.AddAdditionalOption("LT:Options", ltOptions);

            //driver = new RemoteWebDriver(new Uri("https://" + username + ":" + accesskey + gridURL), capabilities.ToCapabilities(),TimeSpan.FromSeconds(600));
           
            driver.Manage().Window.Maximize();
           
        }

        [Test(Description = "Web Scraping Youtube Recent Videos"), Order(1)]
        public void YouTubeScraping()
        {
            string csvFilePath = @"C:\CloudStation\Thomas More DI\devops\Videos.csv";
            string jsonFilePath = @"C:\CloudStation\Thomas More DI\devops\Videos.json";
            driver.Url = test_url_1;
            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            driver.FindElement(By.XPath("//*[@id='content']/div[2]/div[6]/div[1]/ytd-button-renderer[2]")).Click();
            Thread.Sleep(5000);

            WebDriverWait waiter = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1500));
            var elementsWithSearchID = waiter.Until((driver) => driver.FindElements(By.Id("search")));
            var search = elementsWithSearchID.Where(e => e.TagName == "input").FirstOrDefault();
            search.SendKeys("world of warcraft");
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='search-icon-legacy']")).Click();
            Thread.Sleep(8000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='filter-button']/ytd-button-renderer/yt-button-shape/button")).Click();
            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='options']/ytd-search-filter-group-renderer[5]/ytd-search-filter-renderer[2]")).Click();
            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='filter-button']/ytd-button-renderer/yt-button-shape/button")).Click();
            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='options']/ytd-search-filter-group-renderer[3]/ytd-search-filter-renderer[2]")).Click();
            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(5000);
           
            /* Once the page has loaded, scroll to the end of the page to load all the videos */
            /* Scroll to the end of the page to load all the videos in the channel */
            /* Reference - https://stackoverflow.com/a/51702698/126105 */
            /* Get scroll height */
            Int64 last_height = (Int64)(((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight"));
            var count = 0;
            while (count < 1)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.documentElement.scrollHeight);");
                /* Wait to load page */
                Thread.Sleep(2000);
                /* Calculate new scroll height and compare with last scroll height */
                Int64 new_height = (Int64)((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight");
                if (new_height == last_height)
                    /* If heights are the same it will exit the function */
                    break;
                last_height = new_height;
                count++;
            }

            By elem_video_link = By.CssSelector("ytd-video-renderer.style-scope.ytd-item-section-renderer");
            ReadOnlyCollection<IWebElement> videos = driver.FindElements(elem_video_link);
            int maxVideos = 5;

            List<videoData> videoDataList = new List<videoData>();

            Console.WriteLine("Total number of videos in " + test_url_1 + " are " + Math.Min(maxVideos, videos.Count));

            for (int i = 0; i < Math.Min(maxVideos, videos.Count); i++)
            {
                IWebElement video = videos[i];
                IWebElement elem_video_title = video.FindElement(By.XPath(".//*[@id='video-title']"));
                string str_title = elem_video_title.Text;

                IWebElement elem_video_views = video.FindElement(By.XPath(".//*[@id='metadata-line']/span[1]"));
                string str_views = elem_video_views.Text;

                IWebElement elem_video_uploader = video.FindElement(By.CssSelector("ytd-channel-name.long-byline.style-scope.ytd-video-renderer"));
                string str_uploader = elem_video_uploader.Text;

                IWebElement elem_video_reldate = video.FindElement(By.XPath(".//*[@id='video-title']"));
                string str_rel = elem_video_reldate.GetAttribute("href");

                Console.WriteLine("******* Video " + (i + 1) + " *******");
                Console.WriteLine("Video Title: " + str_title);
                Console.WriteLine("Video Views: " + str_views);
                Console.WriteLine("Video Uploader: " + str_uploader);
                Console.WriteLine("Video Link: " + str_rel);
                Console.WriteLine("\n");

                videoDataList.Add(new videoData { vidTitel = str_title, View = str_views, Uplaoder = str_uploader, vidLink = str_rel });
            }
            // Include the first Console.WriteLine in both CSV and JSON outputs
            videoDataList.Insert(0, new videoData { vidTitel = "Amount of Videos Scraped " + Math.Min(maxVideos, videos.Count).ToString()});

            Console.WriteLine("Scraping Data from Most Recent Videos on Search World Of Warcraft");

            // Include the last Console.WriteLine in both CSV and JSON outputs
            videoDataList.Add(new videoData { vidTitel = "Scraping Data from Most Recent Videos on Search World Of Warcraft - Finished at " + DateTime.Now.ToString() });

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(videoDataList);
            }

            // Write to JSON
            string json = JsonConvert.SerializeObject(videoDataList, Formatting.Indented);
            File.WriteAllText(jsonFilePath, json);
        }
        public class videoData
        {
            public string vidTitel { get; set; }
            public string View { get; set; }
            public string Uplaoder { get; set; }
            public string vidLink { get; set; }
        }

        [Test(Description = "Web Scraping 4 Python Jobs"), Order(2)]
        public void JobScraping()
        {
            string csvFilePath = @"C:\CloudStation\Thomas More DI\devops\Jobs.csv";
            string jsonFilePath = @"C:\CloudStation\Thomas More DI\devops\Jobs.json";
            driver.Url = test_url_2;
            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            Thread.Sleep(5000);
            WebDriverWait waiter = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1500));
            var elementsWithSearchID = waiter.Until((driver) => driver.FindElements(By.XPath("//*[@id='keywords-input']")));
            var search = elementsWithSearchID.Where(e => e.TagName == "input").FirstOrDefault();
            search.SendKeys("python");
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='main-search-button']")).Click();
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(10000);

            /* Once the page has loaded, scroll to the end of the page to load all the videos */
            /* Scroll to the end of the page to load all the videos in the channel */
            /* Reference - https://stackoverflow.com/a/51702698/126105 */
            /* Get scroll height */
            Int64 last_height = (Int64)(((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight"));
            var count = 0;
            while (count < 1)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.documentElement.scrollHeight);");
                /* Wait to load page */
                Thread.Sleep(2000);
                /* Calculate new scroll height and compare with last scroll height */
                Int64 new_height = (Int64)((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight");
                if (new_height == last_height)
                    /* If heights are the same it will exit the function */
                    break;
                last_height = new_height;
                count++;
            }

            By elem_job_link = By.CssSelector("li.search-item.clearfix");
            ReadOnlyCollection<IWebElement> jobs = driver.FindElements(elem_job_link);
            int maxJobs = 5;

            List<jobsData> jobsDataList = new List<jobsData>();

            Console.WriteLine("Total number of Jobs in " + test_url_2 + " are " + Math.Min(maxJobs, jobs.Count));

            for (int i = 0; i < Math.Min(maxJobs, jobs.Count); i++)
            {
                IWebElement job = jobs[i];
                IWebElement elem_job_title = job.FindElement(By.CssSelector("h2.job-title"));
                string str_title = elem_job_title.Text;

                IWebElement elem_job_bedrijf = job.FindElement(By.CssSelector("span.job-company"));
                string str_bedrijf = elem_job_bedrijf.Text;

                IWebElement elem_job_keyword= job.FindElement(By.CssSelector("span.job-keywords"));
                string str_keyword = elem_job_keyword.Text;

                IWebElement elem_job_location = job.FindElement(By.CssSelector("span.job-location"));
                string str_location = elem_job_location.Text;

                IWebElement elem_job_links = job.FindElement(By.CssSelector("a.job-title.search-item-link"));
                string str_links = elem_job_links.GetAttribute("href");

                Console.WriteLine("******* Job " + (i + 1) + " *******");
                Console.WriteLine("Job Title: " + str_title);
                Console.WriteLine("Bedrijf: " + str_bedrijf);
                Console.WriteLine("Keywords: " + str_keyword);
                Console.WriteLine("Locatie: " + str_location);
                Console.WriteLine("Link: " + str_links);
                Console.WriteLine("\n");

                jobsDataList.Add(new jobsData { Titel = str_title, Bedrijf = str_bedrijf, Keywords = str_keyword, Locatie = str_location, Link = str_links });
            }
            // Include the first Console.WriteLine in both CSV and JSON outputs
            jobsDataList.Insert(0, new jobsData { Titel = "Amount of python Jobs Scraped " + Math.Min(maxJobs, jobs.Count).ToString() });

            Console.WriteLine("Scraping Data from ICTJOBS, 5 Python jobs");

            // Include the last Console.WriteLine in both CSV and JSON outputs
            jobsDataList.Add(new jobsData { Titel = "Scraping Data from ICTJOBS - 5 Python jobs - Finished at " + DateTime.Now.ToString() });

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(jobsDataList);
            }

            // Write to JSON
            string json = JsonConvert.SerializeObject(jobsDataList, Formatting.Indented);
            File.WriteAllText(jsonFilePath, json);
        }
        public class jobsData
        {
            public string Titel { get; set; }
            public string Bedrijf { get; set; }
            public string Keywords { get; set; }
            public string Locatie { get; set; }
            public string Link { get; set; }
        }

        [Test(Description = "Web Scraping My Own Dungeon Timers"), Order(3)]
        public void WoWHeadScraping()
        {
            string csvFilePath = @"C:\CloudStation\Thomas More DI\devops\Dungeons.csv";
            string jsonFilePath = @"C:\CloudStation\Thomas More DI\devops\Dungeons.json";
            driver.Url = test_url_3;
            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            driver.FindElement(By.XPath("//*[@id='qc-cmp2-ui']/div[2]/div/button[2]")).Click();
            Thread.Sleep(5000);
            WebDriverWait waiter = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1500));
            var elementsWithSearchID = waiter.Until((driver) => driver.FindElements(By.XPath("//*[@id='main-content-container']/div[1]/div[2]/div/div/div/div[2]/div/div[1]/input")));
            var search = elementsWithSearchID.Where(e => e.TagName == "input").FirstOrDefault();
            search.SendKeys("Airisu");
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='react-autowhatever-1--item-1']")).Click();
            Thread.Sleep(4000);

            /* Once the page has loaded, scroll to the end of the page to load all the videos */
            /* Scroll to the end of the page to load all the videos in the channel */
            /* Reference - https://stackoverflow.com/a/51702698/126105 */
            /* Get scroll height */
            Int64 last_height = (Int64)(((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight"));
            var count = 0;
            while (count < 2)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.documentElement.scrollHeight);");
                /* Wait to load page */
                Thread.Sleep(2000);
                /* Calculate new scroll height and compare with last scroll height */
                Int64 new_height = (Int64)((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight");
                if (new_height == last_height)
                    /* If heights are the same it will exit the function */
                    break;
                last_height = new_height;
                count++;
            }

            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='content']/div/div[1]/div[2]/div[2]/section[2]/div[1]/div/div[2]/div/ul")).Click();
            Thread.Sleep(2000);
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            driver.FindElement(By.XPath("//*[@id='item_0$Menu']/li[7]/span/div")).Click();
            Thread.Sleep(2000);

            By elem_dungeons_table = By.XPath("//*[@id='content']/div/div[1]/div[2]/div[2]/section[2]/div[4]/div/div[3]/table");
            ReadOnlyCollection<IWebElement> tables = driver.FindElements(elem_dungeons_table);
            int maxDungeons = 8;
            List<dungeonData> dungeonDataList = new List<dungeonData>();
            // Assuming the rows are under a 'tbody' tag
            var rows = tables[0].FindElements(By.XPath(".//tbody/tr"));

            Console.WriteLine("Total number of dungeons in " + test_url_3 + " are " + Math.Min(maxDungeons, rows.Count));

            for (int i = 0; i < Math.Min(maxDungeons, rows.Count); i++)
            {
                var row = rows[i];

                // Assuming the title is in the first cell (td) of each row
                IWebElement elem_dungeon_title = row.FindElement(By.XPath(".//td[1]"));
                string str_title = elem_dungeon_title.Text;

                // Assuming the views are in the second cell (td) of each row
                IWebElement elem_dungeon_timer = row.FindElement(By.XPath(".//td[5]"));
                string str_timer = elem_dungeon_timer.Text;

                Console.WriteLine("******* Dungeon " + (i + 1) + " *******");
                Console.WriteLine("Dungeon: " + str_title);
                Console.WriteLine("Timer: " + str_timer);
                Console.WriteLine("\n");

                dungeonDataList.Add(new dungeonData { Dungeon = str_title, Timer = str_timer });
            }
            // Include the first Console.WriteLine in both CSV and JSON outputs
            dungeonDataList.Insert(0, new dungeonData { Dungeon = "Amount of Dungeons Scraped " + Math.Min(maxDungeons, rows.Count).ToString() });

            Console.WriteLine("Scraping Data from Raider IO My Character and Dungeon Timers");

            // Include the last Console.WriteLine in both CSV and JSON outputs
            dungeonDataList.Add(new dungeonData { Dungeon = "Scraping Data from Raider IO My Character and Dungeon Timers - Finished at "+ DateTime.Now.ToString() });

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(dungeonDataList);
            }

            // Write to JSON
            string json = JsonConvert.SerializeObject(dungeonDataList, Formatting.Indented);
            File.WriteAllText(jsonFilePath, json);
        }
        public class dungeonData
        {
            public string Dungeon { get; set; }
            public string Timer { get; set; }
        }
        [TearDown]
        public void close_Browser()
        {
            driver.Quit();
        }
    }
}