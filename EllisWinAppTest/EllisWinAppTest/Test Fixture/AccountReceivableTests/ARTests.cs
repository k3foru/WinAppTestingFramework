using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.AccountsReceivableWindow;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.AccountReceivableTests
{
    [CodedUITest]
    public class AccountReceivableTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
        }

        [TestMethod]
        public void VerifyARLandingPageDefaults()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            Assert.IsTrue(ARWindow.VerifyInvoices("All"), "UnPaid invoices are not displayed as All");
            Assert.IsTrue(ARWindow.VerifyMyOrg("Is Collecting"), "My Org is not equal to Is Collecting on landing page");
            Assert.IsTrue(ARWindow.VerifyOverDueDisplayed(), "Over dues are not displayed");
            Assert.IsTrue(ARWindow.VerifyCustomerProfileWindowDisplayedWhenCustomerNumberClicked(),
                "Customer profile page is not displayed when customer on landing page is clicked");

            Cleanup();
        }

        [TestMethod]
        public void VerifyARRescheduleNewNoteTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            RightClick.Reschedule();
            Assert.IsTrue(ARWindow.VerifyCustomerCollectionsWindowDisplayed(),
                "Customer collections window is not displayed");

            ARWindow.AddNewNoteToUnpaidInvoice();
            Assert.IsTrue(ARWindow.VerifyNewNoteAdded(), "New note is not added to the unpaid invoice");
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void VerifyARCancelCallbackNewNoteTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            RightClick.CallBack();
            Assert.IsTrue(ARWindow.AddNoteWindowDisplayed(), "Add Note window is not displayed");

            ARWindow.AddNewNoteToCancelCallback();
            Assert.IsTrue(ARWindow.VerifyNewNoteAdded(), "New note is not added to the cancel callback");
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void VerifyARCustomerProfileDefaultsTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            ARWindow.SelectCustomerCollectionFromLandingPage();
            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            Assert.IsTrue(CustomerProfileWindow.VerifyNoteTypeOptionsDisplayed(), "Note type options are not displayed.");

            CustomerProfileWindow.SelectTab(CCTabConstants.Notes);
            ARWindow.ClickAddNoteButton();

            Assert.IsTrue(ARWindow.VerifyBalanceDueAmountDisplayed(), "Balance amount is not displayed");
            Assert.IsTrue(ARWindow.VerifyProfileWindowDisplayedWhenCusNameLinkClicked(),
                "Customer profile window is not displayed when customer name link is clicked");
            CustomerProfileWindow.CloseCustomerProfileWindow();

            Assert.IsTrue(ARWindow.VerifyProfileWindowDisplayedWhenCusNumberLinkClicked(),
                "Customer profile window is not displayed when customer number link is clicked");
            CustomerProfileWindow.CloseCustomerProfileWindow();

            Cleanup();
        }

        [TestMethod]
        public void LockboxInvoiceSearchTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();

            LandingPage.ClickOnCalendarButton(LandingPage.LandingPageControls.Advanced);
            LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedFromDate, "11/16/2009");
            LandingPage.ClickDateTextbox(LandingPage.LandingPageControls.AdvancedToDate);
            Playback.Wait(2000);

            ARWindow.SelectFirstCustomerInvoiceFromTable();
            ARWindow.SelectFirstRemittenceFromTable();
            ARWindow.ClickOpeninNewWindowButton();

            Assert.IsTrue(ARWindow.VerifyPaymentInvoiceWindowDisplayed(),
                "Payment profile invoice window is not displayed when clicked on Open In New Window");
            Assert.IsTrue(ARWindow.VerifyRemainingAmountDisplayedOnWindow(),
                "Remaining Amount is not displayed on window");

            ARWindow.ClosePaymentInvoiceWindow();
            TitlebarActions.ClickClose((WinWindow) ARWindow.GetPaymentLockboxWindowProperties());
            TitlebarActions.ClickClose((WinWindow) ARWindow.GetPaymentProfileWindowProperties());

            Cleanup();
        }

        [TestMethod]
        public void ARApplyTransactionTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();

            LandingPage.SelectCustomerInvoicesFromNavigationExplorer();

            Assert.IsTrue(ARWindow.VerifyCustomerInvGridDisplayed(), "Customer Invoice Grid is not displayed");
        }

        [TestMethod]
        public void ARDefaultToOwnCostDMUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsDMUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Assert.IsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("1926 - NW Empire"),
                "Default district is not equal to 1926 - NW Empire");

            Cleanup();
        }

        [TestMethod]
        public void ARDefaultToAllCorporateARRUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARRUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Assert.IsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        public void ARDefaultToAllCorporateAVPUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsAVPUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "Invoices", SearchWindow.SearchTypeConstants.Advanced);

            Assert.IsTrue(ARAdvancedSearchWindow.VerifyInvoicingOrganizationIsNull(),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        public void ARDefaultToAllCorporateNABSUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Assert.IsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        public void ARDefaultToAllCorporateARMUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Assert.IsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        private static void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}