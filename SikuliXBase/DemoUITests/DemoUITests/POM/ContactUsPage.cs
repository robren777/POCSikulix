using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace DemoUITests.POM
{
   public class ContactUsPage : WebDriver
    {
        public static IWebElement ContactUs_YourName => Instance.FindElement(By.Id("FullName"));
        public static IWebElement ContactUs_Email => Instance.FindElement(By.Id("Email"));
        public static IWebElement ContactUs_Enquiry => Instance.FindElement(By.Id("Enquiry"));
        public static IWebElement SendEmailButton => Instance.FindElement(By.Name("send-email"));
    }
}
