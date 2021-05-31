using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.IO;

namespace Test_7zip_Desktop_App
{
    public class Tests7zip
    {
        private const string AppiumServerUrl = "http://[::1]:4723/wd/hub";
        private const string AppForTesting = @"C:\Users\7-Zip\7zFM.exe";
        private const string Path7zip = @"C:\Users\7-Zip\";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> driverDesktop;
        private string workDir;

        [OneTimeSetUp]
        public void Setup()
        {
            var appiumOptions = new AppiumOptions() { PlatformName = "windows" };
            appiumOptions.AddAdditionalCapability("app",AppForTesting);
            driver = new WindowsDriver<WindowsElement>(new Uri(AppiumServerUrl), appiumOptions);

            var appiumOptionsDesktop = new AppiumOptions() { PlatformName = "windows" };
            appiumOptionsDesktop.AddAdditionalCapability("app", "Root");
            driverDesktop = new WindowsDriver<WindowsElement>(new Uri(AppiumServerUrl), appiumOptionsDesktop);



            workDir = Directory.GetCurrentDirectory() + @"\workdir";
            if (Directory.Exists(workDir))
               Directory.Delete(workDir, true);
               Directory.CreateDirectory(workDir);
            
        }

        [Test]
        public void Test7zip()
        {
            //Find the Textboxpanel and send the directory path
            var textBoxPanel = driver.FindElementByClassName("Edit");
            textBoxPanel.SendKeys(Path7zip);
            textBoxPanel.SendKeys(Keys.Enter);

            var listFilesToZip = driver.FindElementByClassName("SysListView32");
            listFilesToZip.SendKeys(Keys.Control + "a");

            var buttonAddToArchieve = driver.FindElementByName("Add");
            buttonAddToArchieve.Click();

            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));

            // Create the archive

            var createArchiveWindow = driverDesktop.FindElementByName("Add to Archive");
            var textBoxArchiveName = createArchiveWindow.FindElementByXPath("/Window/ComboBox/Edit[@Name ='Archive:']");
            string archiveFileName = workDir + @"\" + "archive.7z";
            textBoxArchiveName.SendKeys(archiveFileName);

            var textBoxArchiveFormat = createArchiveWindow.FindElementByXPath
                ("/Window/ComboBox[@Name ='Archive format:']");
            textBoxArchiveFormat.SendKeys("7z");

            var textBoxLevelOfCompression = createArchiveWindow.FindElementByXPath
               ("/Window/ComboBox[@Name ='Compression level:']");
            textBoxLevelOfCompression.SendKeys("Ultra");

            var textBoxCompressionMethod = createArchiveWindow.FindElementByXPath
               ("/Window/ComboBox[@Name ='Compression method:']");
            textBoxCompressionMethod.SendKeys(Keys.Home);

            var textBoxDictionarySize = createArchiveWindow.FindElementByXPath
               ("/Window/ComboBox[@Name ='Dictionary size:']");
            textBoxDictionarySize.SendKeys(Keys.End);

            var textBoxWordSize = createArchiveWindow.FindElementByXPath
              ("/Window/ComboBox[@Name ='Word size:']");
            textBoxWordSize.SendKeys(Keys.End);

            var buttonCreateArchiveOk = createArchiveWindow.FindElementByXPath
              ("/Window/Button[@Name ='OK']");
            buttonCreateArchiveOk.Click();

            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(2));

            // Extract the archive

            textBoxPanel.SendKeys(archiveFileName + Keys.Enter);
            var buttonExtract= driver.FindElementByName("Extract");
            buttonExtract.Click();

            var buttonExtractOk = driver.FindElementByName("OK");
            buttonExtractOk.Click();

            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(2));

            //Assert whether the files are equal

            FileAssert.AreEqual(workDir + @"\7zFM.exe", Path7zip + @"\7zFM.exe");

        }

        [OneTimeTearDown]
        public void ShutDown()
        {
          driver.Quit();
        }
    }
}
