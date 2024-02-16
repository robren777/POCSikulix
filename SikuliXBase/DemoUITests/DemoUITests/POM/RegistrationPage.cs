using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace DemoUITests.POM
{
    class RegistrationPage : WebDriver
    {
        public static IWebElement FemaleRadioButton => Instance.FindElement(By.Id("gender-female"));
        public static IWebElement MaleRadioButton => Instance.FindElement(By.Id("gender-male"));
        public static IWebElement FirstName => Instance.FindElement(By.Id("FirstName"));
        public static IWebElement LastName => Instance.FindElement(By.Id("LastName"));
        public static IWebElement Email => Instance.FindElement(By.Id("Email"));
        public static IWebElement Password => Instance.FindElement(By.Id("Password"));
        public static IWebElement ConfirmPassword => Instance.FindElement(By.Id("ConfirmPassword"));
        public static IWebElement RegisterButton => Instance.FindElement(By.Id("register-button"));
        public static IWebElement RegisterConfirmation => Instance.FindElement(By.ClassName("result"));
        //(By.CssSelector("div.result"))
        public static IWebElement ContinueAfterRegisterButton => Instance.FindElement(By.CssSelector("input.button-1.register-continue-button"));
    }
}
