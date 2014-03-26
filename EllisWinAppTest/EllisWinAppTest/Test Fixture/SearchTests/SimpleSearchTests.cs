using System;
using System.Collections.Generic;
using System.Data;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EllisWinAppTest.Windows.EllisWindow;

namespace EllisWinAppTest.SearchTests
{
    [CodedUITest]
    public class SimpleSearchTest : AppContext
    {
        public IEnumerable<DataRow> Initialize()
        {
            WindowsActions.KillEllisProcesses();
            //App = EllisHome.LaunchEllisAsDiffUser();
            App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.SimpleSearch);
            App.SetFocus();
            return datarows;
        }

        [TestMethod]
        public void CustomerSimpleSearch()
        {
            var datarows = Initialize();
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[4].ToString().Equals("Customer"))
                {
                    Console.WriteLine(dataRow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(dataRow.ItemArray[5].ToString(), "Customer",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);
                    Globals.CustomerName = dataRow.ItemArray[6].ToString();

                    Assert.IsTrue(SimpleSearchWindow.VerifyDisplayedResults(Globals.CustomerName),
                        "Displayed results does not contain: " + dataRow.ItemArray[5]);
                    CustomerProfile.CloseCustomerProfile();
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void WorkerSimpleSearch()
        {
            var datarows = Initialize();
            foreach (var dataRow in datarows)
            {
                string TypeOne = dataRow.ItemArray[4].ToString();

                if (TypeOne.Equals("Worker"))
                {
                    Console.WriteLine(dataRow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(dataRow.ItemArray[5].ToString(), TypeOne,
                        SearchWindow.SearchTypeConstants.Simple);
                    Globals.WorkerName = dataRow.ItemArray[6].ToString();

                    WorkerProfile.SelectWorkerFromResultsWindow();

                    //Assert.IsTrue(SimpleSearchWindow.VerifyWorkerDisplayedResults(Globals.WorkerName), "Displayed results does not contain: " + dataRow.ItemArray[5]);
                    //WorkerProfile.CloseWorkerProfile();
                    //WorkerProfile.CloseResultsWindow();
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void NotificationSimpleSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("Lockout"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Lockout",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void BillingLineItemSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("BillingLineItem"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "BillingLineItem",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void InvoicesSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("Invoices"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Invoices",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void CreditCardSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("CreditCard"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "CreditCard",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void InvoiceTransactionsSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("ITransactions"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "ITransactions",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void InvoiceRelationShipsSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("IRelationShips"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "IRelationShips",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void JobOrderSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("JobOrder"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "JobOrder",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void QuoteSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("Quote"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Quote",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void DispatchSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("Dispatch"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Dispatch",
                        SearchWindow.SearchTypeConstants.Simple);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void WorkTicketSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("WorkTicket"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "WorkTicket",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        [TestMethod]
        public void CheckRegisterSearch()
        {
            var datarows = Initialize();


            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[4].ToString().Equals("CheckRegister"))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "CheckRegister",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);

                    //Assert.IsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
                    //SimpleSearchWindow.ClickRefineSearchClose();
                    SimpleSearchWindow.ClickResultWindowClose();
                    Playback.Wait(5000);
                }
            }
            Cleanup();
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}