using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Configuration;
using NUnit.Framework;
using System.Text;
using System.Linq;

namespace DemoUITests
{
    public static class HiglightElementWhileClicking
    {
        public static void ClickAndHighlight(this IWebElement self, int delay, string highlightStyleJS = null)
        {
            self.Click();
            if (highlightStyleJS != null)
            {
                WebDriver.HiglightElement(self, delay, highlightStyleJS);
            }
            else
            {
                WebDriver.HiglightElement(self, delay);
            }
        }
    }

    public class WebDriver
    {
        private const string highlightStyleJavaScript = @"arguments[0].style.cssText = ""border-width: 4px; border-style: solid; border-color: yellow; background-color: rgba(255, 255, 0, 0.3)"";";
        private const string clearHighlightStyleJavaScript = "arguments[0].setAttribute('style', \"\");";
        

        public static IWebDriver Instance { get; set; }
      
        public static string BaseUrl => ConfigurationManager.AppSettings["URL"];

        private static Dictionary<int, int> RandomNumbers { get; set; }
        private static Random RandomGenerator { get; set; }
        private static object lockObject = new object();

        public static WebDriverWait Wait { get; set; }



        public static void Initialize()
        {
            RandomGenerator = new Random();
            RandomNumbers = new Dictionary<int, int>();
            NewDriver();
        }

        public static void NewDriver()
        {
            Close();
            
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            Instance = new ChromeDriver(options);
            Wait = new WebDriverWait(Instance, TimeSpan.FromSeconds(5));

        }

        public static void HiglightElement(IWebElement element, int delay, string highlightStyleJS = highlightStyleJavaScript)
        {
            if (element != null)
            {
                var jsDriver = (IJavaScriptExecutor)Instance;
                jsDriver.ExecuteScript(highlightStyleJS, new object[] { element });
                Thread.Sleep(delay);
                jsDriver.ExecuteScript(clearHighlightStyleJavaScript, element);
            }
        }

        public static IWebElement FindAndHighlightElement(By by, int delay)
        {
            var element = Instance.FindElement(by);
            HiglightElement(element, delay);
            return element;
        }


        public static int GetPersistentRandomNumber(int index, int minValue = 1, int maxValue = 10000)
        {
            if (RandomNumbers.ContainsKey(index))
            {
                return RandomNumbers[index];
            }
            else
            {
                lock (lockObject)
                {
                    RandomNumbers[index] = RandomGenerator.Next(minValue, maxValue);
                }
                return RandomNumbers[index];
            }
        }

       

        public static void Close()
        {
            if (Instance != null)
            {
                Instance.Quit();
            }
            Instance = null;
        }

        public static string ClearHighlightStyleJavaScript1 { get; set; }


        public static void AttemptAndWait(int maxWaitTime, int stepTime, Action action)
        {
            var currentWaitTime = 0;
            while (true)
            {
                try
                {
                    action();
                    return;
                }
                catch
                {
                    if (currentWaitTime >= maxWaitTime)
                    {
                        throw;
                    }
                    Thread.Sleep(stepTime);
                    currentWaitTime += stepTime;
                }
            }
        }

        public static bool IsElementPresent(By by)
        {
            try
            {
                Instance.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public static string Capture(IWebDriver driver, string screenShotName)
        {
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            Screenshot screenshot = ts.GetScreenshot();
            string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string finalpth = pth.Substring(0, pth.LastIndexOf("bin")) + "Reports\\" + screenShotName + ".png";
            string localpath = new Uri(finalpth).LocalPath;
            screenshot.SaveAsFile(localpath, OpenQA.Selenium.ScreenshotImageFormat.Png);
            return localpath;
        }

        public static void ElementVisibleCheck()
        {
            //TODO: Implement method to replace obsolete expected condition method
            //TODO: think about creating new class with expected conditions that are obsolete
        }

        public static bool acceptNextAlert = true;
        public bool IsAlertPresent()
        {
            try
            {
                Instance.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }

        public static string CloseAlertAndGetItsText()
        {
            try
            {
                IAlert alert = Instance.SwitchTo().Alert();
                string alertText = alert.Text;
                if (acceptNextAlert)
                {
                    alert.Accept();
                }
                else
                {
                    alert.Dismiss();
                }
                return alertText;
            }
            finally
            {
                acceptNextAlert = true;
            }
        }

        public static IWebElement FindElementByDataBind(IWebElement element, string dataBindProperty, string binding = "text")
        {
            return element.FindElement(By.CssSelector(string.Format("[data-bind*='{0}: {1}']", binding, dataBindProperty)));
        }

        //This method makes Switch To Current Window makes much easier 
        public static void SwitchtoCurrentWindow()
        {
            Instance.SwitchTo().Window(Instance.WindowHandles.Last());
        }
        public static void FillOutField(IWebElement element, string FieldValue)
        {
            element.Click();
            element.Clear();
            element.SendKeys(FieldValue);
        }

        private static StringBuilder verificationErrors;
        public static void AssertOnPageMethod(string MyText, IWebElement element)
        {
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(MyText, element.Text);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        public static void AssertNotEqualOnPageMethod(string MyText, IWebElement element)
        {
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual(MyText, element.Text);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }
    }
}



