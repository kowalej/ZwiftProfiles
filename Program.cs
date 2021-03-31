using CommandLine;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Diagnostics;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace ZwiftProfiles
{
    class Program
    {
        enum Gender
        {
            MALE,
            FEMALE
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

            [Option('w', "ftp", Required = true, HelpText = "FTP (in watts).")]
            public int? FTPWatts { get; set; }

            [Option('e', "executable-path", Required = false, HelpText = "Zwift executable path to launch after profile setup.")]
            public string ExecutablePath { get; set; }
        }

        static void Main(string[] args)
        {
            float heightCm;
            float weightKg;
            Gender gender;

            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
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

                    EditZwiftProfile(o.Username, o.Password, heightCm, weightKg, gender, o.FTPWatts, o.ExecutablePath);
                });
        }

        static void EditZwiftProfile(string username, string password, float heightCm, float weightKg, Gender gender, int? ftpWatts, string zwiftExePath = null)
        {
            // Launch Zwift exe if required.
            if (!string.IsNullOrEmpty(zwiftExePath))
                Process.Start(zwiftExePath);

            // Create ChromeDriver with ability to capture network logs.
            ChromeOptions options = new ChromeOptions();
            options.SetLoggingPreference("performance", LogLevel.All);
            options.AddUserProfilePreference("intl.accept_languages", "en-US");
            options.AddUserProfilePreference("disable-popup-blocking", "true");
            options.AddArgument("test-type");
            options.AddArgument("--disable-gpu");
            options.AddArgument("no-sandbox");
            options.AddArgument("start-maximized");
            IWebDriver driver = new ChromeDriver(options);

            // Set url as the profile page (but we will have to login then get redirected).
            driver.Url = "https://my.zwift.com/profile/edit";

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
                catch (NoSuchElementException)
                {
                    driver.Navigate().Refresh();
                    retries += 1;
                }
            }

            // Click the stupid cookies consent.
            driver.FindElement(By.Id("truste-consent-button")).Click();
            Thread.Sleep(2000);

            // Change first name to random value so we can submit form.
            var firstNameE = driver.FindElement(By.Name("firstName"));
            string firstName = firstNameE.GetAttribute("value");
            Thread.Sleep(1000);
            firstNameE.SendKeys(Guid.NewGuid().ToString());

            // Submit form (with JS due to glitches with Selenium).
            // This will force a PUT request to update the profile. We can later find this request in the
            // network logs and extract the request parameters (auth header, body, full url, etc).
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.evaluate(\"//button[text()='Save Changes']\", document, null, XPathResult.ANY_TYPE, null).iterateNext().click()");

            // Let request finish.
            Thread.Sleep(4000);
            
            // Extracting the performance logs (i.e. network logs).
            var logs = driver.Manage().Logs.GetLog("performance");

            string url = null;
            string authHeader = null;
            string body = null;

            foreach (var log in logs)
            {
                var json = JsonConvert.DeserializeObject<JObject>(log.Message);
                var urlJ = json.SelectToken("message.params.request.url");
                var method = json.SelectToken("message.params.request.method");

                // Find PUT request to profile.
                if (urlJ != null && ((string)urlJ).Contains("/api/profiles/me") && method != null && ((string)method) == "PUT") {
                    url = (string)urlJ;
                    var authHeaderJ = json.SelectToken("message.params.request.headers.Authorization");
                    var bodyJ = json.SelectToken("message.params.request.postData");
                    // Extract token and body.
                    if (authHeaderJ != null && bodyJ != null) {
                        authHeader = (string) authHeaderJ;
                        body = (string) bodyJ;
                    }
                }
            }

            if (url == null) {
                throw new Exception("Url could not be found! Exiting immediately.");
            }

            if (authHeader == null) {
                throw new Exception("Auth token could not be found! Exiting immediately.");
            }
            
            if (body == null) {
                throw new Exception("Body could not be found! Exiting immediately.");
            }

            // Modify the original request JSON.
            var bodyM = JsonConvert.DeserializeObject<dynamic>(body);
            JsonConvert.SerializeObject(bodyM);
            bodyM.firstName = firstName; // Fixes name
            bodyM.height = (int)(heightCm * 10); // mm
            bodyM.weight = (int)(weightKg * 1000); // grams
            bodyM.male = gender == Gender.MALE ? true : false;
            if (ftpWatts.HasValue) {
                bodyM.ftp = ftpWatts;
            }
            
            // Replay the request with the modified values.
            var client = new RestClient(url);
            var request = new RestRequest();
            request.Body = new RequestBody("application/json", "body", bodyM.ToString());
            request.AddHeader("Authorization", authHeader);
            request.AddHeader("Content-Type", "application/json");
            var response = client.Put(request);
            if (!response.IsSuccessful) {
                Debug.Write(response);
                throw new Exception("Profile update request not successful.");
            }

            // End session.
            driver.Quit();
        }
    }
}
