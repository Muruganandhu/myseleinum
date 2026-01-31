using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Runtime.InteropServices;

public static class WebDriverFactory
{
    public static IWebDriver CreateWebDriver(string downloadPath)
    {
        var options = new ChromeOptions();

        // 1. AUTOMATIC BINARY DETECTION
        string chromePath = GetChromeBinaryPath();
        if (!string.IsNullOrEmpty(chromePath))
        {
            options.BinaryLocation = chromePath;
        }

        // 2. DOCKER & HEADLESS FLAGS
        options.AddArgument("--headless=new");
        options.AddArgument("--no-sandbox");
        options.AddArgument("--disable-dev-shm-usage");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--window-size=1920,1080");

        // 3. DOWNLOAD CONFIG
        options.AddUserProfilePreference("download.default_directory", downloadPath);
        options.AddUserProfilePreference("download.prompt_for_download", false);

        // Selenium 4.11+ automatically handles the Driver/Browser version matching
        return new ChromeDriver(options);
    }

    private static string GetChromeBinaryPath()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            // List of common Windows paths, including the user-specific cache from your error
            string[] paths = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application\chrome.exe"),
                @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                // This covers your specific error path dynamically
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @".cache\selenium\chrome\win64\145.0.7632.26\chrome.exe")
            };

            return paths.FirstOrDefault(File.Exists);
        }

        // Linux path (Docker)
        return "/usr/bin/google-chrome";
    }
}