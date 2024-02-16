using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace DemoUITests.POM
{
    class SearchPage : WebDriver
    {
        public static IWebElement SearchPageTitle => Instance.FindElement(By.CssSelector("h1"));
        public static IWebElement SearchKeywordInputField => Instance.FindElement(By.Id("Q"));
        public static IWebElement SearchButton => Instance.FindElement(By.CssSelector("input.button-1.search-button"));
        public static IWebElement AdvancedSearchOption => Instance.FindElement(By.Id("As"));
        public static IWebElement AutomaticallySearchSubCategories => Instance.FindElement(By.Id("Isc"));
        public static IWebElement SearchInProductDescriptions => Instance.FindElement(By.Id("Sid"));
        public static IWebElement SearchCategory => Instance.FindElement(By.Id("Cid"));
        public static IWebElement SearchManufacturer => Instance.FindElement(By.Id("Mid"));

        //validations and notes
        public static IWebElement SearchValidationNote => Instance.FindElement(By.CssSelector("strong.warning"));
        public static IWebElement NoSearchResults => Instance.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='Search In product descriptions'])[1]/following::strong[1]"));

        //searched items
        public static By SearchedItem = By.XPath("//img[@alt='Picture of Build your own cheap computer']");
        public static By SelectedItemJewelery = By.XPath("//img[@alt='Picture of Black & White Diamond Heart']");


        /* list of page components */
        /* list of page actions */
        /* list of keywords */
    }
}
