using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace DemoUITests.POM 
{
    public class ManageAccountPage : WebDriver
    {
        public static IWebElement MyAccount_Addresses => Instance.FindElement(By.LinkText("Addresses"));
        public static IWebElement AddNewAddressButton => Instance.FindElement(By.CssSelector("input.button-1.add-address-button"));
        public static IWebElement Address_FirstName => Instance.FindElement(By.Id("Address_FirstName"));
        public static IWebElement Address_LastName => Instance.FindElement(By.Id("Address_LastName"));
        public static IWebElement Address_Email => Instance.FindElement(By.Id("Address_Email"));
        public static IWebElement Address_Company => Instance.FindElement(By.Id("Address_Company"));
        public static IWebElement Address_Country => Instance.FindElement(By.Id("Address_CountryId"));
        public static IWebElement Address_City => Instance.FindElement(By.Id("Address_City"));
        public static IWebElement Address_Address1 => Instance.FindElement(By.Id("Address_Address1"));
        public static IWebElement Address_PostalCode => Instance.FindElement(By.Id("Address_ZipPostalCode"));
        public static IWebElement Address_PhoneNum => Instance.FindElement(By.Id("Address_PhoneNumber"));
        public static IWebElement SaveAddressButton => Instance.FindElement(By.CssSelector("input.button-1.save-address-button"));
        public static IWebElement AddressElementOnList => Instance.FindElement(By.CssSelector("div.section.address-item > div.title > strong"));
        public static IWebElement EditAddressButton => Instance.FindElement(By.CssSelector("input.button-2.edit-address-button"));
        public static IWebElement DeleteAddressButton => Instance.FindElement(By.CssSelector("input.button-2.delete-address-button"));
        public static IWebElement NoAddressNote => Instance.FindElement(By.CssSelector("div.address-list"));
    }
}
