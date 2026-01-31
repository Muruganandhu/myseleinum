using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MySeleniumProject;
using SeleniumExtras.WaitHelpers;
namespace MySeleniumProject
{
    internal class Program
    {
        static List<double> inputValue = new List<double>();
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("Starting Selenium Automation...");
            string downloadPath = Path.Combine(Directory.GetCurrentDirectory(), "downloads");
            if (!Directory.Exists(downloadPath)) Directory.CreateDirectory(downloadPath);

            // Load Config from Environment Variables (Secrets)
            string email = Environment.GetEnvironmentVariable("GOCHART_EMAIL");
            string pass = Environment.GetEnvironmentVariable("GOCHART_PASS");
            string symbols = Environment.GetEnvironmentVariable("SYMBOLS") ?? "NIFTY";

            IWebDriver driver = WebDriverFactory.CreateWebDriver(downloadPath);

            try
            {
                while (true)
                {
                    if (!IsTradingHours())
                    {
                        Console.WriteLine($"[{GetISTTime()}] Outside 9:00-15:30 IST. Sleeping...");
                        await Task.Delay(TimeSpan.FromMinutes(5));
                        continue;
                    }

                    Console.WriteLine($"[{GetISTTime()}] Starting cycle for: {symbols}");

                    // --- Your Existing Logic ---
                    await LoginToGoCharting(driver, email, pass);
                    foreach (var s in symbols.Split(','))
                    {
                        await DownloadExcel(driver, s);
                        await ProcessCSV(downloadPath, s);
                    }

                    Console.WriteLine("Cycle finished. Waiting 60 seconds...");
                    await Task.Delay(TimeSpan.FromMinutes(1));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal Error: {ex.Message}");
            }
            finally
            {
                driver.Quit();
            }
        }

        private static bool IsTradingHours()
        {
            DateTime istNow = GetISTTime();
            if (istNow.DayOfWeek == DayOfWeek.Saturday || istNow.DayOfWeek == DayOfWeek.Sunday) return false;

            TimeSpan start = new TimeSpan(9, 0, 0);
            TimeSpan end = new TimeSpan(15, 30, 0);
            return istNow.TimeOfDay >= start && istNow.TimeOfDay <= end;
        }

        private static DateTime GetISTTime()
        {
            DateTime utcNow = DateTime.UtcNow;
            try
            {
                TimeZoneInfo istZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Kolkata");
                return TimeZoneInfo.ConvertTimeFromUtc(utcNow, istZone);
            }
            catch
            {
                return utcNow.AddHours(5).AddMinutes(30); // Fallback
            }
        }


        private static async Task LoginToGoCharting(IWebDriver driver, string email,string password)
        {

            driver.Navigate().GoToUrl("https://gocharting.com/sign-in");

            
            Console.WriteLine($"{email} {password}");
            try
            {

                driver.FindElement(By.XPath("//input[@type='email']")).SendKeys(email);
                driver.FindElement(By.XPath("//input[@type='password']")).SendKeys(password);
                driver.FindElement(By.XPath("//span[text()='Continue']")).Click();

            }
            catch (Exception ex)
            {

            }
            await Task.Delay(5000); // wait for login to complete
                                    // Console.ReadLine();

        }


        public static IWebElement WaitForUserAvatar(IWebDriver driver)
        {
            By avatarBy = By.CssSelector("img[alt='User Avatar']");

            try
            {
                // First wait (default 10 seconds)
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                return wait.Until(ExpectedConditions.ElementExists(avatarBy));
            }
            catch (WebDriverTimeoutException)
            {
                // Extra wait of 5 seconds
                WebDriverWait extraWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                return extraWait.Until(ExpectedConditions.ElementExists(avatarBy));
            }
        }

        private static async Task DownloadExcel(IWebDriver driver, string Symbol)
        {

            // var txtStocks = driver.FindElement(By.Id("react-select-2-input"));
            await Task.Delay(7000);
            try
            {
                var avatar = WaitForUserAvatar(driver);
                avatar.Click(); // or any action
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");

                // 1. Take the screenshot
                if (driver is ITakesScreenshot ts)
                {
                    var screenshot = ts.GetScreenshot();
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "error_screenshots");
                    if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                    string fileName = $"failure_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    screenshot.SaveAsFile(Path.Combine(path, fileName));
                    Console.WriteLine($"Screenshot saved: {fileName}");
                }
                throw; // Re-throw so GitHub Actions knows the job failed
            }

            try
            {
                driver.Navigate().GoToUrl($"https://gocharting.com/terminal?ticker=NSE:{Symbol}");

                await Task.Delay(7000); // wait for chart to load
                var min3 = driver.FindElement(By.Id("interval-selector"));//("interval-selector-btn"));
                min3.Click();
                var mins3 = driver.FindElement(By.XPath("//*[@title='1 Minute']"));
                mins3.Click();

                Actions actions = new Actions(driver);

                // Press ALT + D
                //actions
                //    .KeyDown(Keys.Alt)
                //    .SendKeys("1")
                //    .KeyUp(Keys.Alt)
                //    .Perform();

                var moreSettings = driver.FindElement(By.Id("more-features-btn"));
                moreSettings.Click();
                var isExcel = driver.FindElement(By.Id("Excel-Download"));
                bool isChecked = isExcel.Selected;
                if (!isChecked)
                {
                    isExcel.Click();
                }
                moreSettings.Click();
                //var downloadBtn = driver.FindElement(By.XPath("//div[@title='Download to Excel']"));
                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                //IWebElement element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("[class*='ExcelContainer']")));
                //element.Click();

                By parentBy = By.CssSelector("div.css-cu9v3.eoj6wn090");
                By childBy = By.CssSelector("div.css-p0jopv.eb64bh74");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                IWebElement secondFromLast = wait.Until(d =>
                {
                    var parent = d.FindElement(parentBy);
                    var items = parent.FindElements(childBy);
                    return items.Count >= 2 ? items[items.Count - 2] : null;
                });

                secondFromLast.Click();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");

                // 1. Take the screenshot
                if (driver is ITakesScreenshot ts)
                {
                    var screenshot = ts.GetScreenshot();
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "error_screenshots");
                    if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                    string fileName = $"failure_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    screenshot.SaveAsFile(Path.Combine(path, fileName));
                    Console.WriteLine($"Screenshot saved: {fileName}");
                }
                throw; // Re-throw so GitHub Actions knows the job failed
            }
            //var downloadBtn = driver.FindElement(By.CssSelector("[class*='ExcelContainer']"));
            //downloadBtn.Click();

            Console.WriteLine("Excel downloaded...");
            //await Task.Delay(3000);
        }

        private static async Task ProcessCSV(string folderPath, string Symbol)

        {
            inputValue.Clear();
            //set timeout for 2 seconds         to ensure file is completely written
            await Task.Delay(2000);

            string filePath = Directory.GetFiles(folderPath, "*.csv")
                                       .OrderByDescending(f => File.GetCreationTime(f))
                                       .FirstOrDefault();

            if (filePath == null)
            {
                Console.WriteLine("CSV file not found!");
                return;
            }

            string[] lines = await File.ReadAllLinesAsync(filePath);
            if (lines.Length < 2) return;

            string head = lines[0];
            bool isEmptyHead = head.Contains("gocharting");
            int index = 0;
            if (isEmptyHead)
            {
                index = 2;
            }

            var result = ConvertCsvToJson(lines, index, Symbol);
            //  var marketData= JsonConvert.DeserializeObject<MarketData>(result);
            await UploadTOAPI(result);
            MoveProcessed(filePath, folderPath);
            Console.WriteLine($"{result}");
        }
        public static List<MarketData> CalculateCumulative(List<MarketData> data)
        {
            double cumVolume = 0;
            double cumDelta = 0;
            double cumBuyVol = 0;
            double cumSellVol = 0;
            double cumOpen = 0;
            double cumHigh = 0;
            double cumLow = 0;
            double cumClose = 0;

            foreach (var item in data)
            {
                cumVolume += item.Volume;
                cumDelta += item.Delta;
                cumBuyVol += item.BuyVolume;
                cumSellVol += item.SellVolume;
                cumOpen += item.Open;
                cumHigh += item.High;
                cumLow += item.Low;
                cumClose += item.Close;
                cumVolume += item.Volume;

                item.CumOpen = cumOpen;
                item.CumHigh = cumHigh;
                item.CumLow = cumLow;
                item.CumClose = cumClose;
                item.CumVolume = cumVolume;

                item.CumVolume = cumVolume;
                item.CumDelta = cumDelta;
                item.CumBuyVolume = cumBuyVol;
                item.CumSellVolume = cumSellVol;
            }

            return data;
        }

        private static string ConvertCsvToJson(string[] lines, int index, string Symbol)
        {
            // Extract headers (first line)
            string[] headers = lines[index].Split(',');
            var dataLines = lines.Skip(index).ToList();
            // List to store JSON objects for each row
            List<Dictionary<string, string>> jsonList = new List<Dictionary<string, string>>();

            //var headers = dataLines[index].Split(','); // TSV, not comma-separated
            var rows = dataLines.Skip(1);

            var dataList = new List<MarketData>();

            foreach (var row in rows)
            {
                var values = row.Split(',');

                if (values.Length < 16) continue;

                decimal _buyVolume = ParseLong(values[11]);
                if (_buyVolume > 0)
                {
                    string cleaned = values[0].Trim('"');
                    var data = new MarketData
                    {
                        Date = cleaned,
                        Symbol = Symbol,
                        Open = ParseDouble(values[1]),
                        High = ParseDouble(values[2]),
                        Low = ParseDouble(values[3]),
                        Close = ParseDouble(values[4]),
                        Volume = ParseLong(values[5]),
                        Open_interest = ParseLong(values[6]),
                        Delta = ParseDouble(values[7]),
                        MaxDelta = ParseDouble(values[8]),
                        MinDelta = ParseDouble(values[9]),
                        CumDelta = ParseDouble(values[10]),
                        BuyVolume = ParseLong(values[11]),
                        SellVolume = ParseLong(values[12]),
                        Vwap = ParseDouble(values[13]),
                        BuyVwap = ParseDouble(values[14]),
                        SellVwap = ParseDouble(values[15])
                    };

                    dataList.Add(data);
                }
            }


            // Iterate over the remaining lines (excluding the header)
            //foreach (string line in lines.Skip(index + 1))
            //{
            //   // Split the line into columns
            //      string[] data = line.Split(',');

            //   // Create a dictionary to hold the key - value pairs(header->data)
            //    Dictionary<string, string> jsonObject = new Dictionary<string, string>();

            //    for (int i = 0; i < headers.Length; i++)
            //    {
            //        // Add each header as key and corresponding value as value
            //        if (i < data.Length)
            //        {
            //            jsonObject[headers[i]] = data[i];
            //        }
            //    }

            //    // Add the dictionary (representing a row) to the JSON list
            //    jsonList.Add(jsonObject);
            //}
            CalculateCumulative(dataList);

            // Serialize the list to JSON format
            return JsonConvert.SerializeObject(dataList, Formatting.Indented);
        }
        private static void MoveProcessed(string path, string folderPath)
        {
            Console.WriteLine("Moving files to processed");
            string destinationDirectory = folderPath + "\\processed";
            if (File.Exists(path))
            {
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }
                string extension = ".csv";
                string uniqueFileName = $"{DateTime.UtcNow.Ticks}{extension}";
                string destinationFilePath = Path.Combine(destinationDirectory, uniqueFileName);
                File.Move(path, destinationFilePath);
            }
        }
        static double ParseDouble(string s)
        {
            string cleaned = s.Trim('"');
            double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double result);
            return result;
        }

        static long ParseLong(string s)
        {
            string cleaned = s.Trim('"');
            long.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out long result);
            return result;
        }
        private static async Task HighlightDataInExcel(string folderPath)
        {
            string filePath = Directory.GetFiles(folderPath, "*.csv")
                                       .OrderByDescending(f => File.GetCreationTime(f))
                                       .FirstOrDefault();

            if (filePath == null) return;

            string excelFilePath = Path.ChangeExtension(filePath, ".xlsx");

            FileInfo fileInfo = new FileInfo(excelFilePath);
            using ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            worksheet.Cells["A1"].LoadFromText(await File.ReadAllTextAsync(filePath), new ExcelTextFormat { Delimiter = ',' });
            // worksheet.DeleteRow(0,2);
            int rows = worksheet.Dimension.Rows;
            int cumDeltaColumn = 0;

            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (worksheet.Cells[1, col].Text == "CUMDelta")
                {
                    cumDeltaColumn = col;
                    break;
                }
            }

            if (cumDeltaColumn == 0) return;

            // Apply color logic
            for (int row = 2; row <= rows; row++)
            {
                if (double.TryParse(worksheet.Cells[row, cumDeltaColumn].Text, out double val))
                {
                    if (val > 0)
                        worksheet.Cells[row, cumDeltaColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, cumDeltaColumn].Style.Fill.BackgroundColor.SetColor(
                        val > 0 ? System.Drawing.Color.LightGreen : System.Drawing.Color.LightCoral);
                }
            }

            await package.SaveAsync();
            Console.WriteLine("Excel file saved with highlights.");
        }

        private static async Task UploadTOAPI(string jsonData)
        {
            var apiUrl = "https://app.villagetrader.org/api/v1/automate/uploadgochart";
            using (var client = new HttpClient())
            {
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                try
                {
                    var response = await client.PostAsync(apiUrl, content);
                    string responseString = await response.Content.ReadAsStringAsync();

                    Console.WriteLine("Response:");
                    Console.WriteLine(responseString);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error:");
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }

}