using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace DemoUITests.POM
{
    class HomePage : WebDriver
    {
        //login-logout
        public static IWebElement LogInOption => Instance.FindElement(By.LinkText("Log in"));
        public static IWebElement LogOutOption => Instance.FindElement(By.LinkText("Log out"));
        public static IWebElement EmailField => Instance.FindElement(By.Id("Email"));
        public static IWebElement PasswordField => Instance.FindElement(By.Id("Password"));
        public static IWebElement LogInButton => Instance.FindElement(By.CssSelector("input.button-1.login-button"));


        public static By txtEmail = By.Id("Email");
        public static void PopulateEmail(String email)
        {
            Instance.FindElement(txtEmail).ClickAndHighlight(1500);
            Instance.FindElement(txtEmail).SendKeys(email);
        }

        public static By txtPassword = By.Id("Password");
        public static void PopulatePassword(String pwd)
        {
            Instance.FindElement(txtPassword).ClickAndHighlight(1500);
            Instance.FindElement(txtPassword).SendKeys(pwd);
        }

        //header links
        public static IWebElement AccountLink => Instance.FindElement(By.CssSelector("a.account"));
        public static IWebElement MyAccountOption => Instance.FindElement(By.LinkText("My account"));
        public static IWebElement ContactUsOption => Instance.FindElement(By.LinkText("Contact us"));
        public static IWebElement RegisterOption => Instance.FindElement(By.LinkText("Register"));
        public static IWebElement CartOption => Instance.FindElement(By.CssSelector("span.cart-label"));
        public static IWebElement SearchOption => Instance.FindElement(By.LinkText("Search"));
        public static IWebElement WishlistOption => Instance.FindElement(By.CssSelector("a.ico-wishlist > span.cart-label"));



        //main menu links
        public static IWebElement ComputersLink => Instance.FindElement(By.LinkText("Computers"));
        public static IWebElement DigitalDownloadsLink => Instance.FindElement(By.LinkText("Digital downloads"));
        public static IWebElement ElectronicsLink => Instance.FindElement(By.XPath("(//a[contains(@href, '/electronics')])[3]"));
        public static IWebElement JewelryOption => Instance.FindElement(By.XPath("(//a[contains(@href, '/jewelry')])[3]"));

        public static void LoginAs (string Username, string Password)
        {
            PopulateEmail(Username);
            PopulatePassword(Password);
            LogInButton.Click();
        }
    }
}
