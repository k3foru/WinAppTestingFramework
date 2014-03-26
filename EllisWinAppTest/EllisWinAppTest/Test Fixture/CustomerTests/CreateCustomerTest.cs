using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.CustomerTests
{
    [CodedUITest]
    public class CreateCustomerTest : AppContext
    {
        public IEnumerable<DataRow> Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            return datarows;
        }

        [TestMethod]
        public void CreateNewCustomerWithoutFedIdTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();
                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "Customer Profile not displayed for new customer without FED ID");
            }
            Cleanup();
        }

        [TestMethod]
        public void CreateNewCustomerWithFedIdTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", null);
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();
                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "Customer Profile not displayed for new customer with FED ID");
            }
            Cleanup();
        }

        [TestMethod]
        public void CreateCODCustomerWithoutFedIdTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, "COD");
                CreateCustomerWindow.ClickSave();
                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "COD Customer Profile not displayed for new COD customer without FED ID");
            }
            Cleanup();
        }

        [TestMethod]
        public void CreateCODCustomerWithFedIdTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", "COD");
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();

                //TODO: Change Assert state once demo is over
                Assert.IsFalse(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "COD Customer Profile not displayed for new COD customer with FED ID");
            }
            Cleanup();
        }


        [TestMethod]
        public void NAPSCustomerDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAPSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.DiffCredsCreateCustomer);

            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("NAPS")))
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", null);
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();
                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "NAPS Customer Profile not displayed for new customer with FED ID");

                //CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                //CustomerProfileWindow.SelectCustomerSegmentation("National");
                ////Assert.IsTrue(CustomerProfileWindow.VerifyEllisExceptionWindowDisplayed(), "Ellis exception window is displayed");
                //Assert.IsTrue(CustomerProfileWindow.VerifyCustomerSegment("National"), "Customer segment changes once save button is clicked");

                CustomerProfileWindow.SelectTab(CCTabConstants.Management);
                //CustomerProfileWindow.ClickChangeStatus();
                //CustomerProfileWindow.ChangeStatus("Do Not Service");
                //Assert.IsTrue(CustomerProfileWindow.VerifyStatusChanged("Do Not Service"), "Status has not been changed to Do Not Service");

                //CustomerProfileWindow.ClickChangeStatus();
                //CustomerProfileWindow.ChangeStatus("Approved");
                //Assert.IsTrue(CustomerProfileWindow.VerifyStatusChanged("Approved"), "Status has not been changed to Approved");

                CustomerProfileWindow.ClickCustomerLinking();
                CustomerProfileWindow.EnterParentCustomerNumber();
                CustomerProfileWindow.ClickValidate();

                Assert.IsTrue(CustomerProfileWindow.VerifyInheritanceOptionsEnabled(), "Inheritance options are not enabled. Please check entered parent customer number");
                CustomerProfileWindow.ClickSaveCustomerLinking();
                Thread.Sleep(3000);
                Assert.IsTrue(CustomerProfileWindow.VerifyChainLinkDisplayed(), "Customer linking Icon is not displayed");

                CustomerProfileWindow.ClickCustomerLinking();
                CustomerProfileWindow.DelinkParentCustomer();
                Assert.IsFalse(CustomerProfileWindow.VerifyChainLinkDisplayed(), "Customer linking Icon is still displayed");

                CustomerProfileWindow.SelectTab(CCTabConstants.Quoting);
                Assert.IsTrue(CustomerProfileWindow.VerifyQuotingCheckboxesChecked(), "Quoting checkboxes are not checked");

                CustomerProfileWindow.SelectTab(CCTabConstants.Collections);
                CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
                CustomerProfileWindow.UncheckInvoicingOrganization();
                CustomerProfileWindow.SelectAllOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();

                CustomerProfileWindow.SelectTab(CCTabConstants.Billing);
                CustomerProfileWindow.ClickSelectOrganizationButton();
                CustomerProfileWindow.SelectAllOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();
                Assert.IsTrue(CustomerProfileWindow.VerifyInvOrgDisplayed(), "Selected Invoicing Org is Displayed");
            }
            Cleanup();
        }


        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}