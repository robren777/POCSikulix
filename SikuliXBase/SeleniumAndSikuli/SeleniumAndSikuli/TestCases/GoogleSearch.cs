using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using Sikuli4Net.sikuli_UTIL;
using Sikuli4Net.sikuli_REST;
using OpenQA.Selenium.Chrome;
using System.Threading;
using SikuliSharp;

namespace SeleniumAndSikuli
{
    [TestClass]
    public class GoogleSearch
    {
        private IWebDriver webDriver = null;
        private APILauncher launcher = new APILauncher(true);


        [TestMethod]
        public void GoogleSearchTest()
        {
            launcher.Start();
            GetWebDriver("https://www.google.com/");

            Screen screen = new Screen();
            Pattern pattern_SearchText = new Pattern(@"c:\users\stefani\documents\visual studio 2017\Projects\SeleniumAndSikuli\SeleniumAndSikuli\Images\SearchText.PNG");
            Pattern pattern_SearchButton = new Pattern(@"c:\users\stefani\documents\visual studio 2017\Projects\SeleniumAndSikuli\SeleniumAndSikuli\Images\SearchButton.PNG");
            Pattern pattern_GoogleIcon = new Pattern(@"c:\users\stefani\documents\visual studio 2017\Projects\SeleniumAndSikuli\SeleniumAndSikuli\Images\Google.PNG");

            screen.Wait(pattern_SearchText, 250);
            screen.Type(pattern_SearchText, "Test", KeyModifier.NONE);
            screen.Wait(pattern_GoogleIcon, 200);
            screen.Click(pattern_GoogleIcon);

            screen.Wait(pattern_SearchButton, 200);
            screen.Click(pattern_SearchButton);

            Thread.Sleep(200);
            webDriver.Quit();

            launcher.Stop();
        }
        
        [TestMethod]
        public void ClickDesktopIcon()
        {
            launcher.Start();
            Screen screen = new Screen();
            //Pattern pattern_FolderIcon = new Pattern("path/FolderIcon.PNG");
            //Pattern pattern_FolderIcon = new Pattern("C:\\Users\\osvaldo.zamora\\Documents\\GitHub\\SikuliXBase\\SeleniumAndSikuli\\SeleniumAndSikuli\\Images\\InicioWindows.png");
            Pattern pattern_FolderIcon = new Pattern("C:\\SikuliX\\Images\\InicioWindows.png"); 
            screen.Wait(pattern_FolderIcon, 500);
            screen.DoubleClick(pattern_FolderIcon);
            launcher.Stop();
        }
       
        private void GetWebDriver(string url)
        {
            webDriver = new ChromeDriver();
            webDriver.Manage().Window.Maximize();
            webDriver.Navigate().GoToUrl(url);
            Thread.Sleep(1500);
        }
        
    
}
}
