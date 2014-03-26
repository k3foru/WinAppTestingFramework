using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;

namespace EllisWinAppTest.SearchTests
{
    [CodedUITest]
    public class AdvanceSearchTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
            App = EllisHome.LaunchEllisAsCSRUser();
            Thread.Sleep(5000);
            //App.SetFocus();
        }

        [TestMethod]
        public void CustomerAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CustomerAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                Console.WriteLine(dataRow.ItemArray[24]);
                SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
                CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow);
                CustomerAdvanceSearchWindow.ClickOnSearchButton();
                Assert.IsTrue(CustomerAdvanceSearchWindow.VerifySearchResultsWindowDisplayed(),
                    "Search validation error window was not displayed");
                Playback.Wait(3000);
                CustomerAdvanceSearchWindow.CloseSearchResultsWindow();
            }
            Cleanup();
        }

        [TestMethod]
        public void VerifyCustomerSearchValidationErrorDisplayedTest()
        {
            Initialize();
            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.ClickOnSearchButton();
            Playback.Wait(2000);
            Assert.IsTrue(CustomerAdvanceSearchWindow.VerifySearchValidationWindowDisplayed(),
                "Search validation error window was not displayed");
            Playback.Wait(2000);
            CustomerAdvanceSearchWindow.ClickValidationOk();
            CustomerAdvanceSearchWindow.CloseSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        public void WorkerAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAdvancedSearch);
            foreach (var dataRow in datarows)
            {
             
                SearchWindow.SelectSearchElements(null, "Worker", SearchWindow.SearchTypeConstants.Advanced);
                WorkerAdvancedSearchWindow.EnterWorkerAdvancedSearchData(dataRow);
                WorkerAdvancedSearchWindow.ClickSearchBtn();
                Playback.Wait(3000);
                var summary = WorkerSummaryWindow.ClickOnCloseBtn();
                Assert.IsTrue(summary,"Worker results window not displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void LockoutAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.LockOutAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                SearchWindow.SelectSearchElements(null, "Lockout", SearchWindow.SearchTypeConstants.Advanced);
                NotificationAdvanceSearchWindow.EnterNotificationAdvancedSearchData(dataRow);
                NotificationAdvanceSearchWindow.ClickSearchBtn();
                Playback.Wait(3000);
            }
            Cleanup();
        }

        [TestMethod]
        public void BillingAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                //Console.WriteLine(dataRow.ItemArray[3]);
                SearchWindow.SelectSearchElements(null, "BillingLineItem", SearchWindow.SearchTypeConstants.Advanced);
                ARAdvancedSearchWindow.EnterBillingSearchData(dataRow);
                ARAdvancedSearchWindow.ClickSearchBtn();
                Playback.Wait(3000);
            }
            Cleanup();
        }

        [TestMethod]
        public void CreditCardAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                //Console.WriteLine(dataRow.ItemArray[3]);
                SearchWindow.SelectSearchElements(null, "CreditCard", SearchWindow.SearchTypeConstants.Advanced);
                ARAdvancedSearchWindow.EnterCreditCardSearchData(dataRow);
                ARAdvancedSearchWindow.ClickSearchBtn();
                Playback.Wait(3000);
            }
            Cleanup();
        }

        [TestMethod]
        public void CustomerCalendarFilterTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Customers");
            LandingPage.ClickOnCalendarButton(LandingPage.LandingPageControls.Advanced);
            LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedFromDate, "01012014");
            LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedToDate, "02272014");
            LandingPage.ClickOnCalendarClient();

            CustomerProfileWindow.SelectFirstCustomerFromTable();
            CustomerProfileWindow.VerifyProfileDefaults();
            CustomerProfile.CloseCustomerProfile();

            Cleanup();
        }

        //[TestMethod]
        //public void InvoiceAdvanceSearch()
        //{
        //    Initialize();
        //    var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CustomerAdvancedSearch);
        //    foreach (var dataRow in datarows)
        //    {
        //        Console.WriteLine(dataRow.ItemArray[24]);
        //        SearchWindow.SelectSearchElements(null, "Invoices", SearchWindow.SearchTypeConstants.Advanced);
        //        //    CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //        //    Playback.Wait(3000);


        //    }
        //    Cleanup();
        //}

        //[TestMethod]
        //public void InvoiceTransactionAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("ITransactions"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("ITransactions");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void InvoiceRelationShipsAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("IRelationShips"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("IRelationShips");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void JobOrderAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("JobOrder"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("JobOrder");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void QuoteAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("Quote"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("Quote");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void DispatchAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("Dispatch"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("Dispatch");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void WorkTicketAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("WorkTicket"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("WorkTicket");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        //[TestMethod]
        //public void CheckRegisterAdvanceSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        if (dataRow.ItemArray[4].ToString().Equals("CheckRegister"))
        //        {
        //            Console.WriteLine(dataRow.ItemArray[3]);
        //            CustomerAdvanceSearchWindow.AdvanceSearch("CheckRegister");
        //            CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow.ItemArray[5].ToString());
        //            Playback.Wait(3000);


        //        }
        //    }
        //    Cleanup();
        //}
        
        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}