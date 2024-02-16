using System;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using AventStack.ExtentReports;
using DemoUITests.TestData;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;


namespace DemoUITests
{
    [TestClass]
    public class Tests : BaseTestClass
    {
        private StringBuilder verificationErrors;

        [TestMethod]
        public void Registration()
        {
            // var i
            var rand = new Random();
            var i = rand.Next(1, 100000);
            string Email = "test" + i + "@test.eu";
            Init();
            test = extent.CreateTest("Registration");
            test.AssignCategory("Regression", "Functional");
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            POM.HomePage.RegisterOption.Click();
            POM.RegistrationPage.FemaleRadioButton.Click();
            POM.RegistrationPage.FirstName.ClickAndHighlight(1000);
            POM.RegistrationPage.FirstName.SendKeys("Name");
            POM.RegistrationPage.LastName.ClickAndHighlight(1500);
            POM.RegistrationPage.LastName.SendKeys("Lastname");
            POM.RegistrationPage.Email.Click();
            POM.RegistrationPage.Email.SendKeys(Email);
            POM.RegistrationPage.Password.ClickAndHighlight(1500);
            POM.RegistrationPage.Password.SendKeys(ConfigurationManager.AppSettings["Pass"]);
            POM.RegistrationPage.ConfirmPassword.ClickAndHighlight(1500);
            POM.RegistrationPage.ConfirmPassword.SendKeys(ConfigurationManager.AppSettings["Pass"]);
            POM.RegistrationPage.RegisterButton.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.RegistrationCompletedNote, POM.RegistrationPage.RegisterConfirmation.Text);
                test.Log(Status.Info, "Registration was successfull");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.RegistrationPage.ContinueAfterRegisterButton.Click();
            Thread.Sleep(2000);
            //check if user logged in
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(Email, POM.HomePage.AccountLink.Text);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.HomePage.LogOutOption.Click();
            Cleanup();
           
        }

        [TestMethod]
        public void LogInLogOut()
        {

            Init();
            test = extent.CreateTest("Log in - Log out");
            test.AssignCategory("Regression", "Functional");
            test.AssignAuthor("smajic");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            POM.HomePage.LogInOption.Click();
            WebDriver.FillOutField(POM.HomePage.EmailField, ConfigurationManager.AppSettings["User"]);
            POM.HomePage.PopulatePassword(ConfigurationManager.AppSettings["Pass"]);
            POM.HomePage.LogInButton.Click();
            Thread.Sleep(2000);
            //check if user logged in
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(ConfigurationManager.AppSettings["User"], POM.HomePage.AccountLink.Text);
                test.Log(Status.Info, "User is logged in");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Thread.Sleep(2000);
            POM.HomePage.LogOutOption.Click();
            Cleanup();

        }

        [TestMethod]
        public void Navigation()
        {
            Init();
            test = extent.CreateTest("Navigation");
            test.AssignCategory("Functional", "Regression");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            POM.HomePage.ComputersLink.Click();
            POM.ProductPage.DesktopComputers.Click();
            POM.HomePage.DigitalDownloadsLink.Click();
            POM.HomePage.ElectronicsLink.Click();
            POM.ProductPage.CellPhonesElectronics.Click();
            Thread.Sleep(2000);
            Cleanup();
        }

        [TestMethod]
        public void AddToChart()
        {
            Init();
            test = extent.CreateTest("Add to chart");
            test.AssignCategory("Regression", "Functional");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            POM.HomePage.ComputersLink.Click();
            POM.ProductPage.DesktopComputers.Click();
            POM.ProductPage.Product.Click();
            Thread.Sleep(1000);
            var AddToChartButton = POM.ProductPage.AddToChartButton;
            Actions actions = new Actions(WebDriver.Instance);
            actions.MoveToElement(AddToChartButton);
            actions.Perform();
            POM.ProductPage.AddToChartButton.Click();
            try
            {
                WebDriverWait wait = new WebDriverWait(WebDriver.Instance, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementIsVisible(POM.ProductPage.AddedToCartNotification));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.AddedToCartNote, POM.ProductPage.AddToCartNotification.Text);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Thread.Sleep(5000);
            POM.HomePage.CartOption.Click();
            Thread.Sleep(1000);
            POM.ProductPage.Cart_TermOfService.Click();
            POM.ProductPage.Cart_Checkout.Click();
            Cleanup();
        }

        [TestMethod]
        public void NegativeTestCase()
        {
            Init();
            test = extent.CreateTest("Negative Test Case");
            test.AssignCategory("New feature", "Negative test");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            test.Info("Test should try to click on an item that does not exists and fail");
            if (WebDriver.IsElementPresent(By.LinkText("NotExistingItem")))
            {
                test.Log(Status.Pass, "Item found");
            }
            else
            {
                test.Log(Status.Fail, "Item not found");
            }
            
            Cleanup();
        }

        [TestMethod]
        public void ContactUs()
        {
            // var i
            var rand = new Random();
            var i = rand.Next(1, 100000);
            string Email = "contact" + i + "@test.eu";
            Init();
            test = extent.CreateTest("Contact us");
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.ContactUsOption.Click();
            WebDriver.FillOutField(POM.ContactUsPage.ContactUs_YourName, "Contact name");
            WebDriver.FillOutField(POM.ContactUsPage.ContactUs_Email, Email);
            WebDriver.FillOutField(POM.ContactUsPage.ContactUs_Enquiry, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
            POM.ContactUsPage.SendEmailButton.Click();
            Cleanup();
        }

        [TestMethod]
        public void Search()
        {
            Init();
            test = extent.CreateTest("Search");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.SearchOption.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Search", POM.SearchPage.SearchPageTitle.Text);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.SearchPage.SearchKeywordInputField.Click();
            POM.SearchPage.SearchKeywordInputField.SendKeys("computer");
            POM.SearchPage.SearchButton.Click();
            if (WebDriver.IsElementPresent(POM.SearchPage.SearchedItem))
            {
                test.Log(Status.Pass, "Item found, search is working");
            }
            else
            {
                test.Log(Status.Fail, "Item not found, please check your search or search criteria");
            }
            Cleanup();
        }

        [TestMethod]
        public void AdvancedSearchCheckValidationNote()
        {
            Init();
            test = extent.CreateTest("Search - validation note for term minimum lenght");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.SearchOption.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Search", POM.SearchPage.SearchPageTitle.Text);
                test.Log(Status.Info, "On Search page");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.SearchPage.SearchKeywordInputField.Click();
            POM.SearchPage.SearchKeywordInputField.SendKeys("co");
            POM.SearchPage.AdvancedSearchOption.Click();
            POM.SearchPage.AutomaticallySearchSubCategories.Click();
            POM.SearchPage.SearchInProductDescriptions.Click();
            POM.SearchPage.SearchButton.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.SearchTermValidationNote, POM.SearchPage.SearchValidationNote.Text);
                test.Log(Status.Info, "Validation note is shown");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Cleanup();
        }

        [TestMethod]
        public void AdvancedSearchNoResults()
        {
            Init();
            test = extent.CreateTest("Advanced search without results");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.SearchOption.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Search", POM.SearchPage.SearchPageTitle.Text);
                test.Log(Status.Info, "On Search page");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.SearchPage.SearchKeywordInputField.Click();
            POM.SearchPage.SearchKeywordInputField.Clear();
            POM.SearchPage.SearchKeywordInputField.SendKeys("computer");
            POM.SearchPage.AdvancedSearchOption.Click();
            POM.SearchPage.SearchInProductDescriptions.Click();
            POM.SearchPage.AutomaticallySearchSubCategories.Click();
            POM.SearchPage.SearchCategory.Click();
            new SelectElement(POM.SearchPage.SearchCategory).SelectByText("Computers");
            POM.SearchPage.SearchCategory.Click();
            POM.SearchPage.SearchManufacturer.Click();
            new SelectElement(POM.SearchPage.SearchManufacturer).SelectByText("Tricentis");
            POM.SearchPage.SearchManufacturer.Click();
            POM.SearchPage.SearchInProductDescriptions.Click();
            POM.SearchPage.SearchButton.Click();

            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.NoProductsFoundNote, POM.SearchPage.NoSearchResults.Text);
                test.Log(Status.Info, "No result note is shown");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Cleanup();
        }

        [TestMethod]
        public void AdvancedSearch()
        {
            Init();
            test = extent.CreateTest("Advanced search with result");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.SearchOption.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Search",POM.SearchPage.SearchPageTitle.Text);
                test.Log(Status.Info, "On Search page");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.SearchPage.SearchKeywordInputField.Click();
            POM.SearchPage.SearchKeywordInputField.Clear();
            POM.SearchPage.SearchKeywordInputField.SendKeys("computer");
            POM.SearchPage.AdvancedSearchOption.Click();
            POM.SearchPage.AutomaticallySearchSubCategories.Click();
            POM.SearchPage.SearchCategory.Click();
            new SelectElement(POM.SearchPage.SearchCategory).SelectByText("Computers");
            POM.SearchPage.SearchCategory.Click();
            POM.SearchPage.SearchInProductDescriptions.Click();
            POM.SearchPage.SearchButton.Click();

            if (WebDriver.IsElementPresent(POM.SearchPage.SearchedItem))
            {
                test.Log(Status.Pass, "Item found");
            }
            else
            {
                test.Log(Status.Fail, "Item not found, please check your search or search criteria");
            }
            Cleanup();
        }

        [TestMethod]
        public void AddAndRemoveFromWishlist()
        {
            Init();
            test = extent.CreateTest("Add and remove item from wishlist");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            POM.HomePage.JewelryOption.Click();
       
            WebDriver.Instance.FindElement(POM.SearchPage.SelectedItemJewelery).Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Black & White Diamond Heart", POM.ProductPage.ProductName.Text);
                test.Log(Status.Info, "Correct item opened");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.ProductPage.AddToWishlistButton.Click();
            try
            {
                WebDriverWait wait = new WebDriverWait(WebDriver.Instance, TimeSpan.FromSeconds(3));
                wait.Until(ExpectedConditions.ElementIsVisible(POM.ProductPage.AddedToCartNotification));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.AddedToWishlistNote, WebDriver.Instance.FindElement(POM.ProductPage.AddedToCartNotification).Text);
                test.Log(Status.Info, "Note for adding item to wishlist appeared");

            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.HomePage.WishlistOption.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Black & White Diamond Heart", POM.ProductPage.WislistItem.Text);
                test.Log(Status.Info, "Correct item in the wishlist");

            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            POM.ProductPage.RemoveFromWishlist.Click();
            POM.ProductPage.UpdateWishlist.Click();
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.EmptyWishlistNote, POM.ProductPage.EmptyWishlistNote.Text);
                test.Log(Status.Info, "Wishlist is empty");

            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Cleanup();
        }

        [TestMethod]
        public void AccountManagement_Address()
        {
            var rand = new Random();
            var i = rand.Next(1, 100000);
            Init();
            test = extent.CreateTest("Account management - address");
            verificationErrors = new StringBuilder();
            WebDriver.Instance.Navigate().GoToUrl(WebDriver.BaseUrl);
            Thread.Sleep(2000);
            //login
            POM.HomePage.LogInOption.Click();
            POM.HomePage.LoginAs(ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]);   
            // add address
            POM.HomePage.MyAccountOption.Click();
            POM.ManageAccountPage.MyAccount_Addresses.Click();
            POM.ManageAccountPage.AddNewAddressButton.Click();
            WebDriver.FillOutField(POM.ManageAccountPage.Address_FirstName, "Name");
            WebDriver.FillOutField(POM.ManageAccountPage.Address_LastName, "LastName");
            WebDriver.FillOutField(POM.ManageAccountPage.Address_Email, ConfigurationManager.AppSettings["User"]);
            WebDriver.FillOutField(POM.ManageAccountPage.Address_Company, "My company" + i);
            POM.ManageAccountPage.Address_Country.Click();
            new SelectElement(POM.ManageAccountPage.Address_Country).SelectByText("Argentina");
            POM.ManageAccountPage.Address_Country.Click();
            WebDriver.FillOutField(POM.ManageAccountPage.Address_City, "Bounes Aires");
            WebDriver.FillOutField(POM.ManageAccountPage.Address_Address1, "Ovo je moja adresa" + i);
            WebDriver.FillOutField(POM.ManageAccountPage.Address_PostalCode, "3658AMz" + i);
            WebDriver.FillOutField(POM.ManageAccountPage.Address_PhoneNum, "178596933666");
            POM.ManageAccountPage.SaveAddressButton.Click();
            //verify created address
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Name LastName", POM.ManageAccountPage.AddressElementOnList.Text);
                test.Log(Status.Info, "Address created");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            //verify edit button is visible
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Edit", POM.ManageAccountPage.EditAddressButton.GetAttribute("value"));
                test.Log(Status.Info, "Edit button is visible");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            //edit address
            POM.ManageAccountPage.EditAddressButton.Click();
            WebDriver.FillOutField(POM.ManageAccountPage.Address_PostalCode, "56329" + i);
            POM.ManageAccountPage.SaveAddressButton.Click();
            //delete address
            POM.ManageAccountPage.DeleteAddressButton.Click();
            WebDriver.acceptNextAlert = true;
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(Regex.IsMatch(WebDriver.CloseAlertAndGetItsText(), POM.ValidationsAndNotes.DeleteAddressConfirmation));
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(POM.ValidationsAndNotes.NoAddressesNote, POM.ManageAccountPage.NoAddressNote.Text);
                test.Log(Status.Info, "There are no available addresses, address successfully deleted");
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
            Cleanup();
            
        }
            
        }
    }

