using CommandLine;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
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

            // Login to Zwift.
            driver.FindElement(By.Id("username")).SendKeys(username);
            driver.FindElement(By.Id("password")).SendKeys(password);
            driver.FindElement(By.Id("submit-button")).Click();

            // Get current selected units (Imperial / Metric).
            var se = new SelectElement(driver.FindElement(By.Id("displayUnit")));
            var so = se.SelectedOption;
            
            // Set profile to metric.
            driver.FindElement(By.Id("displayUnit")).Click();
            {
                IWebElement dropdown = driver.FindElement(By.Id("displayUnit"));
                dropdown.FindElement(By.XPath("//option[. = 'Metric']")).Click();
            }
            driver.FindElement(By.Id("displayUnit")).Click();

            // Set metric height (cm).
            driver.FindElement(By.Id("metricHeight")).Clear();
            driver.FindElement(By.Id("metricHeight")).SendKeys(heightCm.ToString());
            Debug.WriteLine($"Height entry: {heightCm.ToString()}");

            // Set metric weight (kg).
            driver.FindElement(By.Id("metricWeight")).Clear();
            driver.FindElement(By.Id("metricWeight")).SendKeys(weightKg.ToString());
            Debug.WriteLine($"Weight entry: {weightKg.ToString()}");

            // Set gender nth-child = 1 for male, nth-child=2 for female.
            string nGenderChild = gender == Gender.MALE ? "1" : "2";
            driver.FindElement(By.CssSelector($".form-radio:nth-child({nGenderChild}) > .dummy")).Click();
            Debug.WriteLine($"Gender entry: {nGenderChild}");

            // Reset units to initial.
            driver.FindElement(By.Id("displayUnit")).Click();
            {
                IWebElement dropdown = driver.FindElement(By.Id("displayUnit"));
                so.Click();
            }
            driver.FindElement(By.Id("displayUnit")).Click();

            // Submit form.
            driver.FindElement(By.CssSelector(".btn-zwift")).Click();

            // Wait for save submission complete.
            Thread.Sleep(2000);

            // End session.
            driver.Quit();
        }
    }
}
