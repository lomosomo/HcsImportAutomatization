using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using CommandLine;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace HcsImportAutomatization
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<CommandLineOptions>(args)
                .WithParsed(Execute);
        }

        public static void Execute(CommandLineOptions options)
        {
            var dryRun = !string.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings["DryRun"]) || options.DryRun;
            var importFile = !string.IsNullOrWhiteSpace(options.FileToImport)
                ? options.FileToImport
                : System.Configuration.ConfigurationManager.AppSettings["ImportFile"];

            if (string.IsNullOrWhiteSpace(importFile))
            {
                

                Console.WriteLine("No import file specified, check config file and command line parameters");
            }

            if (!File.Exists(importFile))
            {
                Console.WriteLine($"Import file not found: {importFile}");
                return;
            }

            Console.WriteLine($"Import file: {importFile}");

            if (dryRun)
            {
                Console.WriteLine("DryRun: import will not actually be performed");
            }

            var fileFullPath = CopyFileLocally(importFile);

            var driver = new ChromeDriver();
            driver.Manage().Window.Size = new Size(1024, 768);
            driver.Url = "https://hosted-comm-service.a1.net/";

            // remember handle of main window
            var mainWindowHandle = driver.CurrentWindowHandle;


            // Login
            var uernameElement = driver.FindElement(By.Id("txtUserName"));
            var passwordElement = driver.FindElement(By.Id("txtPassword"));
            var loginButton = driver.FindElement(By.Id("btnLogin"));
            uernameElement.SendKeys("211@61651");
            passwordElement.SendKeys("d9Avi!1hJon#s");
            loginButton.Click();

            // Navigate to external phone book in tree
            driver.SwitchTo().Frame(1);
            driver.SwitchTo().Frame(0);
            driver.FindElement(By.PartialLinkText("Dr.med. Anita Mang")).Click();
            driver.FindElement(By.PartialLinkText("Telefonbücher")).Click();
            driver.FindElement(By.PartialLinkText("Externes Telefonbuch")).Click();

            // Navigate to right side and select import
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().Frame(3);
            driver.SwitchTo().Frame(1);
            driver.FindElement(By.Id("ibImport")).Click();

            // Select file in popup window and generat preview
            var newWindowHandle = driver.WindowHandles.Where(s => s != mainWindowHandle).First();
            driver.SwitchTo().Window(newWindowHandle);
            driver.FindElement(By.Id("uploadFile")).SendKeys(fileFullPath);
            driver.FindElement(By.Id("ibPreview")).Click();
            Thread.Sleep(5000);
            driver.Manage().Window.Maximize();
            driver.FindElement(By.Id("chbOverWrite")).Click();
            
            // Import and wait for result
            if (!options.DryRun)
            {
                driver.FindElement(By.Id("ibImport")).Click();
                WaitForAlert(driver); // Wait till messgebox pops up and accept it
                driver.SwitchTo().Alert().Accept();
            }

            driver.FindElement(By.Id("ibClose")).Click();
            driver.SwitchTo().Window(mainWindowHandle);
            driver.Close();
            driver.Quit();
            Console.WriteLine("Run completed");
        }

        public static string CopyFileLocally(string fileFullPath)
        {
            var tempFile = Path.Combine(Path.GetTempPath(), "HcS Import File.csv");
            File.Copy(fileFullPath, tempFile, true);
            return tempFile;
        }

    

        public static void WaitForAlert(WebDriver driver)
        {
            int i = 0;
            while (i++ < 5)
            {
                try
                {
                    var alert = driver.SwitchTo().Alert();
                    break;
                }
                catch (NoAlertPresentException)
                {
                    Thread.Sleep(1000);
                    continue;
                }
            }
        }

    }
}
