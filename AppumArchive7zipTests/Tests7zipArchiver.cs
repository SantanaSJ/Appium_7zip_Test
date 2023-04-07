using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace AppumArchive7zipTests
{
    public class Tests7zipArchiver
    {

        private const string AppiumUriString = "http://127.0.0.1:4723/wd/hub";
        private const string ZipLocation = @"C:\Program Files\7-Zip\7zFM.exe";
        private const string tempDirectory = @"C:\temp";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> driverArchiveWindow;
        private AppiumOptions options;
        private AppiumOptions optionsArchiveWindow;

        [SetUp]
        public void Setup()
        {
            this.options = new AppiumOptions() { PlatformName = "Windows" };
            this.options.AddAdditionalCapability("app", ZipLocation);
            this.driver = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), this.options);
            this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            this.optionsArchiveWindow = new AppiumOptions() { PlatformName = "Windows" };
            this.optionsArchiveWindow.AddAdditionalCapability("app", "Root");
            this.driverArchiveWindow = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), this.optionsArchiveWindow);
            this.driverArchiveWindow.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
        
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, true);
            }

            Directory.CreateDirectory(tempDirectory);
        
        }

        [TearDown]
        public void TearDown()
        {
            this.driver.Quit();
        }

        [Test]
        public void Test1()
        {
            WindowsElement inputFilePath = this.driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            inputFilePath.SendKeys(@"C:\Program Files\7-Zip\" + Keys.Enter);

            WindowsElement listFiles = this.driver.FindElementByClassName("SysListView32");
            listFiles.SendKeys(Keys.Control + "a");

            WindowsElement buttonAdd = this.driver.FindElementByName("Add");
            buttonAdd.Click();

            WindowsElement windowArchive = this.driverArchiveWindow.FindElementByName("Add to Archive");
            
            AppiumWebElement inputArchivePath = windowArchive.FindElementByClassName("ComboBox");
            inputArchivePath.SendKeys(tempDirectory + @"\archve.7z");

            WindowsElement dropdownFieldArchiveFormat = this.driverArchiveWindow.FindElementByName("Archive format:");
            dropdownFieldArchiveFormat.SendKeys("7z");


            WindowsElement dropdownFieldCompressionLevel = this.driverArchiveWindow.FindElementByName("Compression level:");
            dropdownFieldCompressionLevel.SendKeys("Ultra");

            WindowsElement dropdownFieldCompressionMethod = this.driverArchiveWindow.FindElementByName("Compression method:");
            dropdownFieldCompressionMethod.SendKeys("LZMA2");

            WindowsElement OKButton = this.driverArchiveWindow.FindElementByName("OK");
            OKButton.Click();

            Thread.Sleep(2000);
            inputFilePath.SendKeys(tempDirectory + @"\archve.7z" + Keys.Enter);

            

            WindowsElement buttonExtract = this.driver.FindElementByName("Extract");
            buttonExtract.Click();

            WindowsElement inputFieldCopyTo = this.driver.FindElementByName("Copy to:");
            inputFieldCopyTo.SendKeys(tempDirectory + Keys.Enter);
            Thread.Sleep(1000);

            FileAssert.AreEqual(ZipLocation, tempDirectory + @"\7zFM.exe");
        }
    }
}