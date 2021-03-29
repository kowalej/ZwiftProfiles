using CommandLine;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.Threading;

namespace ZwiftProfiles
{
    class Program
    {
        enum Gender
        {
            MALE,
            FEMALE
        }

        enum Browser
        {
            CHROME,
            FIREFOX
        }

        public class Options
        {
            [Option('u', "username", Required = true, HelpText = "Zwift profile username.")]
            public string Username { get; set; }

            [Option('p', "password", Required = true, HelpText = "Zwift profile password.")]
            public string Password { get; set; }

            [Option('c', "centimeters", Required = false, HelpText = "Metric height (in cm).")]
            public float? HeightCm { get; set; }

            [Option('k', "kilograms", Required = false, HelpText = "Metric weight (in kg).")]
            public float? WeightKilos { get; set; }

            [Option('f', "feet", Required = false, HelpText = "Feet component of imperial height.")]
            public int? HeightFeetComponent { get; set; }

            [Option('i', "inches", Required = false, HelpText = "Inch component of imperial height.")]
            public int? HeightInchComponent { get; set; }

            [Option('l', "pounds", Required = false, HelpText = "Imperial weight (in pounds).")]
            public float? WeightPounds { get; set; }

            [Option('g', "gender", Required = true, HelpText = "Gender (m = male, f = female).")]
            public string Gender { get; set; }

            [Option('e', "executable-path", Required = false, HelpText = "Zwift executable path to launch after profile setup.")]
            public string ExecutablePath { get; set; }

            [Option('b', "browser", Required = false, HelpText = "Which browser to use (Chrome or Firefox).")]
            public string Browser { get; set; }
        }

        private static IWebDriver driver;

        static void Main(string[] args)
        {
            Browser browser;
            float heightCm;
            float weightKg;
            Gender gender;

            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
                    // Browser.
                    if (o.Browser.ToLower() == "firefox")
                    {
                        browser = Browser.FIREFOX;
                    }
                    else if (o.Browser.ToLower() == "chrome")
                    {
                        browser = Browser.CHROME;
                    }
                    else
                    {
                        throw new Exception("Invalid browser (enter Firefox or Chrome).");
                    }

                    // Height.
                    if (o.HeightCm.HasValue) // Metric.
                    {
                        if (o.HeightFeetComponent.HasValue || o.HeightInchComponent.HasValue)
                            throw new Exception("Can't set feet and inches when using metric height (cm).");
                        heightCm = o.HeightCm.Value;
                    }
                    else if (o.HeightFeetComponent.HasValue && o.HeightInchComponent.HasValue) // Imperial.
                    {
                        heightCm = (o.HeightFeetComponent.Value * 30.48f) + (o.HeightInchComponent.Value * 2.54f);
                    }
                    else
                    {
                        throw new Exception("No height attributes were set.");
                    }

                    // Weight.
                    if (o.WeightKilos.HasValue) // Metric.
                    {
                        if (o.WeightPounds.HasValue)
                            throw new Exception("Can't set lbs when using metric weight (kgs).");
                        weightKg = o.WeightKilos.Value;
                    }
                    else if (o.WeightPounds.HasValue) // Imperial.
                    {
                        weightKg = o.WeightPounds.Value * 0.453592f;
                    }
                    else
                    {
                        throw new Exception("No weight attributes were set.");
                    }

                    if (o.Gender.ToLower() == "m")
                    {
                        gender = Gender.MALE;
                    }
                    else if (o.Gender.ToLower() == "f")
                    {
                        gender = Gender.FEMALE;
                    }
                    else
                    {
                        throw new Exception("Invalid gender (no spectrum here - enter m or f).");
                    }

                    EditZwiftProfile(browser, o.Username, o.Password, heightCm, weightKg, gender, o.ExecutablePath);
                });
        }

        static void EditZwiftProfile(Browser browser, string username, string password, float heightCm, float weightKg, Gender gender, string zwiftExePath = null)
        {
            // Launch Zwift exe if required.
            if (!string.IsNullOrEmpty(zwiftExePath))
                Process.Start(zwiftExePath);

            // Determine browser to use.
            if (browser == Browser.FIREFOX)
                driver = new FirefoxDriver();
            else if (browser == Browser.CHROME)
                driver = new ChromeDriver();

            // Set url as the profile page (but we will have to login then get redirected).
            driver.Url = "https://my.zwift.com/profile/edit";
            driver.Manage().Window.Maximize();

            // Wait a maximum of 2 seconds while locating elements.
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            // Wait a maximum of 5 seconds for async Javascript calls.
            driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(5);

            // Login to Zwift.
            driver.FindElement(By.Id("username")).SendKeys(username);
            driver.FindElement(By.Id("password")).SendKeys(password);
            driver.FindElement(By.Id("submit-button")).Click();

            // This sequence is to avoid situations where the page doesn't fully load due to some JavaScript glitch.
            // There seems to be a bug on the Zwift website which will cause the async loading to hang and in turn not finish rendering the HTML.
            // The page will only show a loader and not have any of the appropriate fields available.
            // Therefore we will refresh with a small delay afterwards until we detect our input fields are available.
            int retries = 0;
            while (retries < 20)
            {
                try
                {
                    Thread.Sleep(2000);
                    // This is an arbitrary element we know exists in the form.
                    driver.FindElement(By.CssSelector("input[name='units']:checked"));
                    break;
                }
                catch (NoSuchElementException ex)
                {
                    driver.Navigate().Refresh();
                    retries += 1;
                }
            }

            // Click the stupid cookies consent.
            driver.FindElement(By.Id("truste-consent-button")).Click();
            Thread.Sleep(2000);

            // Get current selected units (Imperial / Metric).
            var initialUnits = driver.FindElement(By.CssSelector("input[name='units']:checked"));
            Thread.Sleep(1000);

            // Set profile to metric.
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.querySelector(\"input[name='units'][value='METRIC']\").scrollIntoView()");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("input[name='units'][value='METRIC']")).Click();
            Thread.Sleep(1000);

            // Set metric height (cm).
            js.ExecuteScript("document.querySelector(\"input[name='heightCm']\").scrollIntoView()");
            Thread.Sleep(1000);
            driver.FindElement(By.Name("heightCm")).Clear();
            driver.FindElement(By.Name("heightCm")).SendKeys(Math.Round(heightCm).ToString());
            Debug.WriteLine($"Height entry: {heightCm.ToString()}");

            // Set metric weight (kg).
            js.ExecuteScript("document.querySelector(\"input[name='weightKg']\").scrollIntoView()");
            Thread.Sleep(1000);
            driver.FindElement(By.Name("weightKg")).Clear();
            driver.FindElement(By.Name("weightKg")).SendKeys(Math.Round(weightKg).ToString());
            Debug.WriteLine($"Weight entry: {weightKg.ToString()}");

            // Gender disabled as of 2021-03-22 - fuck off Zwift.
            // // Set gender nth-child = 1 for male, nth-child=2 for female.
            // string nGenderChild = gender == Gender.MALE ? "1" : "2";
            // driver.FindElement(By.CssSelector($".form-radio:nth-child({nGenderChild}) > .dummy")).Click();
            // Debug.WriteLine($"Gender entry: {nGenderChild}");

            // Reset units to initial.
            js.ExecuteScript("document.querySelector(\"input[name='units']:unchecked\").scrollIntoView()");
            Thread.Sleep(1000);
            initialUnits.Click();

            Thread.Sleep(1000);
            // Submit form.
            js.ExecuteScript("document.querySelector(\"//button[text()='Save Changes]\").scrollIntoView()");
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//button[text()='Save Changes']")).Click();

            // Wait for save submission complete.
            Thread.Sleep(2000);

            // End session.
            driver.Quit();
        }
    }
}
