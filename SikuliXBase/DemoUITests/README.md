## UI AUTOMATION TEST EXAMPLES
UI test examples made with C#, NUnit, Selenium and Extent reports.

### Getting started
The created tests cover the basic functionalities of Tricentis Demo WebShop which is available on the URL: http://demowebshop.tricentis.com/. 
The URL is configurable in Environments.config.
In the POM folder the IWebElements for each tested site are written down. 
The additional methods for the WebDriver are inside WebDriver.cs while the BaseTestClass.cs covers the base test setup as well as the Extent report options.
The generated reports are saved in folder Reports.

### Built with
* [NUnit](https://github.com/nunit/docs/wiki) - Test framework for .NET 
* [Extent Report](http://extentreports.com/docs/versions/3/net/) - Reporting tool
* [Selenium Webdriver](https://www.seleniumhq.org/projects/webdriver/) - Selenium Webdriver API, version 3.14.0 for .NET bindings for Selenium Webdriver API