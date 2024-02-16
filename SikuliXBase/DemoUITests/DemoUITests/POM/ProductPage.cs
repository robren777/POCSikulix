using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace DemoUITests.POM
{
    class ProductPage : WebDriver
    {
        public static IWebElement DesktopComputers => Instance.FindElement(By.CssSelector("img[alt=\"Picture for category Desktops\"]"));
        public static IWebElement CellPhonesElectronics => Instance.FindElement(By.CssSelector("img[alt=\"Picture for category Cell phones\"]"));
        public static IWebElement Product => Instance.FindElement(By.CssSelector("img[alt=\"Picture of Build your own cheap computer\"]"));
        public static IWebElement AddToChartButton => Instance.FindElement(By.Id("add-to-cart-button-72"));
        public static IWebElement AddToCartNotification => Instance.FindElement(By.XPath("//div[@id='bar-notification']/p"));
        public static By AddedToCartNotification = By.XPath("//div[@id='bar-notification']/p");
        public static IWebElement Cart_TermOfService => Instance.FindElement(By.Id("termsofservice"));
        public static IWebElement Cart_Checkout => Instance.FindElement(By.Id("checkout"));
        public static IWebElement ProductName => Instance.FindElement(By.CssSelector("h1"));
        public static IWebElement AddToWishlistButton => Instance.FindElement(By.Id("add-to-wishlist-button-14"));
        public static IWebElement WislistItem => Instance.FindElement(By.CssSelector("td.product > a"));
        public static IWebElement RemoveFromWishlist => Instance.FindElement(By.Name("removefromcart"));
        public static IWebElement UpdateWishlist => Instance.FindElement(By.Name("updatecart"));
        public static IWebElement EmptyWishlistNote => Instance.FindElement(By.CssSelector("div.wishlist-content"));


    }
}
