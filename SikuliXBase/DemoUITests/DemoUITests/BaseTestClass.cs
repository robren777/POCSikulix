using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using System;
using System.Net;



namespace DemoUITests
{
    public class BaseTestClass
    {
        public static ExtentHtmlReporter HtmlReporter = new ExtentHtmlReporter(@"..\..\Reports\" + DateTime.Now.ToString("dd-MMM-yyyy-hh-mm-ss") + "_ExtentReport.html");
        public static ExtentReports extent = new ExtentReports();
        public ExtentTest test;
       

        [SetUp]
        public void Init()
        {
            //system info
            string hostname = Dns.GetHostName();
            OperatingSystem os = Environment.OSVersion;
            extent.AddSystemInfo("Operating system: ", os.ToString());
            extent.AddSystemInfo("Host: ", hostname);
            extent.AddSystemInfo("Browser: ", "Google Chrome");

            extent.AttachReporter(HtmlReporter);
            HtmlReporter.LoadConfig(@"..\..\extent-config.xml");
            HtmlReporter.Configuration().JS = "$('.test.warning').each(function() { $(this).addClass('pass').removeClass('warning'); }); $('.test-status.warning').each(function() { $(this).addClass('pass').removeClass('warning').text('pass'); });$('.tests-quick-view .status.warning').each(function() { $(this).addClass('pass').removeClass('warning').text('PASS'); }); testSetChart();";

            WebDriver.Initialize();
        }

        [TearDown]
        public void Cleanup()
        {
            WebDriver.Close();
            var status = TestContext.CurrentContext.Result.Outcome.Status;
            
            var stacktrace = string.IsNullOrEmpty(TestContext.CurrentContext.Result.StackTrace)
                    ? ""
                    : string.Format("{0}", TestContext.CurrentContext.Result.StackTrace);
            
            Status logstatus;

            switch (status)
            {
                case TestStatus.Failed:
                    logstatus = Status.Fail;
                    break;
                //case TestStatus.Inconclusive:
                //    logstatus = Status.Warning;
                //    break;
                case TestStatus.Skipped:
                    logstatus = Status.Skip;
                    break;
                default:
                    logstatus = Status.Pass;
                    break;
            }

            test.Log(logstatus, "Test ended with " + logstatus + stacktrace);
            extent.Flush();
  
        }
    }
}
