using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.CustomerTests
{
    [CodedUITest]
    public class CustomerProfileTest : AppContext
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

        public void VerifyCustomerProfileDefaults()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                Assert.IsTrue(CustomerProfileWindow.VerifyProfileDefaults(),
                    "Profile displayed is not that of the customer selected");
            }
        }

        [TestMethod]
        public void VerifyNABSCustomerProfileDefaults()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                Assert.IsTrue(CustomerProfileWindow.VerifyProfileDefaults(),
                    "Profile displayed is not that of the customer selected");
            }
        }

        [TestMethod]
        public void CreateNewCustomerContactTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

                var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
                foreach (var contactData in datarows2)
                {
                    CustomerProfileWindow.ClickAddContactButton();
                    CustomerProfileWindow.EnterContactDetails(contactData);
                    CustomerProfileWindow.ClickSave();
                    Assert.IsTrue(CustomerProfileWindow.VerifyNewContactAdded(contactData.ItemArray[3].ToString()),
                        "New Contact information is not added to the customer profile");
                    break;
                }
                break;
            }
            Cleanup();
        }

        [TestMethod]
        public void CreateNewNABSCustomerContactTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

                var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
                foreach (var contactData in datarows2)
                {
                    CustomerProfileWindow.ClickAddContactButton();
                    CustomerProfileWindow.EnterContactDetails(contactData);
                    CustomerProfileWindow.ClickSave();
                    Assert.IsTrue(CustomerProfileWindow.VerifyNewContactAdded(contactData.ItemArray[3].ToString()),
                        "New Contact information is not added to the customer profile");
                    break;
                }
                break;
            }
            Cleanup();
        }

        [TestMethod]
        public void EditCustomerContactsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

            var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var contactData in datarows2)
            {
                CustomerProfileWindow.SelectStatusFilter("All");
                CustomerProfileWindow.SelectFirstCustomerContactFromTable();
                CustomerProfileWindow.EnterContactDetails(contactData);
                CustomerProfileWindow.ClickSave();
                break;
            }

            CustomerProfileWindow.SelectStatusFilter("Active");
            Assert.IsTrue(CustomerProfileWindow.VerifyNewContactAdded(Globals.CustomerContact),
                "Customer contact was not updated");

            Cleanup();
        }

        [TestMethod]
        public void EditActiveContactsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

            CustomerProfileWindow.SelectStatusFilter("Active");
            CustomerProfileWindow.SelectFirstCustomerContactFromTable();
            Assert.IsFalse(CustomerProfileWindow.VerifyActiveCheckDisabled(), "Active check is not disabled");

            Cleanup();
        }

        [TestMethod]
        public void AddWorkerCompToCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            Cleanup();
        }

        [TestMethod]
        public void AddWorkerCompToNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            Cleanup();
        }

        [TestMethod]
        public void AddWorkerCompToTACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            Cleanup();
        }

        [TestMethod]
        public void AddWorkerCompToRACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsRAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.ClickAssociatedAddCode();

            Assert.IsTrue(CustomerProfileWindow.VerifyAssociatedCodeTypeSetToPrimary(), "Associated code type is not set to Primary");
            
            Assert.IsTrue(CustomerProfileWindow.VerifyModifierSetToOne(),
                "Modifier is not set to one");

            CustomerProfileWindow.ClickAddModifierWindow();
            Assert.IsTrue(CustomerProfileWindow.VerifyAddModifierWindowDisplayed(), "Add modifier window is not displayed");

            CustomerProfileWindow.CloseAddModifierWindow();
            TitlebarActions.ClickClose((WinWindow) CustomerProfileWindow.GetEditCompCodeModifierWindowProperties());
            CustomerProfileWindow.ClickAssociatedCodeCancel();

            Cleanup();
        }

        [TestMethod]
        public void AddWorkerOrganizationToCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            if (CustomerProfileWindow.VerifySelectOrganizationButtonEnabled())
            {
                CustomerProfileWindow.ClickSelectOrganizationButton();
                CustomerProfileWindow.SelectOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();

                Assert.IsTrue(CustomerProfileWindow.VerifyOrganizationDetailsUpdated(),
                    "Organization Details were not added successfully");
            }

            Cleanup();
        }

        [TestMethod]
        public void VerifyQuotingRulesForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Quoting);

            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.ManageQuotes),
                "Manage Quotes option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.PredefinedContact),
                "Predefined Contract option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.RateAgreementWindow),
                "Rate Agreement option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.OverRideBranchThreshold),
                "Over Ride branch threshold option is enabled.");

            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Save),
                "Save button is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Cancel),
                "Cancel button is enabled.");
            Cleanup();
        }

        [TestMethod]
        public void VerifyQuotingRulesForNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Quoting);

            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.ManageQuotes),
                "Manage Quotes option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.PredefinedContact),
                "Predefined Contract option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.RateAgreementWindow),
                "Rate Agreement option is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.OverRideBranchThreshold),
                "Over Ride branch threshold option is enabled.");

            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Save),
                "Save button is enabled.");
            Assert.IsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Cancel),
                "Cancel button is enabled.");
            Cleanup();
        }

        [TestMethod]
        public void ValidateAndAddNewNoteForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            var notes = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var note in notes)
            {
                CustomerProfileWindow.ClickAddNewNoteButton(note.ItemArray[3].ToString());
                CustomerProfileWindow.AddNewNote(note.ItemArray[4].ToString());
                CustomerProfileWindow.ClickOnSaveAndCloseButton();

                Assert.IsTrue(CustomerProfileWindow.VerifyNewNoteDisplayedInGrid(note.ItemArray[4].ToString()),
                    "New Note is not added to the customer profile");
            }

            Cleanup();
        }

        [TestMethod]
        public void ValidateAndAddNewNoteForNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            var notes = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var note in notes)
            {
                CustomerProfileWindow.ClickAddNewNoteButton(note.ItemArray[3].ToString());
                CustomerProfileWindow.AddNewNote(note.ItemArray[4].ToString());
                CustomerProfileWindow.ClickOnSaveAndCloseButton();

                Assert.IsTrue(CustomerProfileWindow.VerifyNewNoteDisplayedInGrid(note.ItemArray[4].ToString()),
                    "New Note is not added to the customer profile");
            }

            Cleanup();
        }

        [TestMethod]
        public void ValidateNotesForGoAndClearButtonInCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            CustomerProfileWindow.EnterDate("From", "01/01/2013");
            CustomerProfileWindow.EnterDate("To", "02/02/2013");

            CustomerProfileWindow.ClickGoButton();
            Assert.IsTrue(CustomerProfileWindow.VerifyEmptyGridDisplayed(), "GO button did not function as expected");

            CustomerProfileWindow.ClickClearButton();
            Playback.Wait(200);
            CustomerProfileWindow.ClickGoButton();
            Assert.IsFalse(CustomerProfileWindow.VerifyEmptyGridDisplayed(), "Clear button did not function as expected");

            Cleanup();
        }

        [TestMethod]
        public void ValidateCreditHistoryLimitAndDetailsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.CreditInfo);
            CustomerProfileWindow.ClickCreditInfoInnerTab(CCTabConstants.CreditLimit);

            Assert.IsTrue(CustomerProfileWindow.VerifyCreditStatus("Acceptable"),
                "Credit Status is not displayed as Acceptable. Displayed Status is - " + Globals.Temp);
            Assert.IsTrue(CustomerProfileWindow.VerifyCreditTerms(),
                "Credit terms are not equal to either Credit Or COD. Displayed credit term is - " + Globals.Temp);
            Assert.IsTrue(CustomerProfileWindow.VerifyCurrentCreditLimitDisplayed(),
                "Current Credit Limit is not displayed");
            Assert.IsTrue(CustomerProfileWindow.VerifyCreditLimitNotesDisplayed(),
                "Credit Limit Notes grid is not displayed");

            CustomerProfileWindow.ClickCreditInfoInnerTab(CCTabConstants.CreditDetails);

            Assert.IsTrue(CustomerProfileWindow.VerifyCreditStatusDropDown("Acceptable"),
                "Credit Status is not displayed as Acceptable. Displayed Status is - " + Globals.Temp);
            Assert.IsTrue(CustomerProfileWindow.VerifyCreditTermsDropDown(),
                "Credit terms are not equal to either Credit Or COD. Displayed credit term is - " + Globals.Temp);
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Credit Application"),
                "Credit On Application File is not checked");
            Assert.IsTrue(CustomerProfileWindow.VerifyDateSignedOnDisplayed(), "Date signed on was not displayed");
            Assert.IsTrue(CustomerProfileWindow.VerifyBusinessLicensesDisplayed(),
                "Business Licences were not displayed");

            Cleanup();
        }

        [TestMethod]
        public void BillingInvoicingManagementForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.CheckSingleOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                "Billing coordinator displayed is not the selected one.");

            CustomerProfileWindow.ClickCancel();
            Cleanup();
        }

        [TestMethod]
        public void BillingInvoicingManagementForNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.CheckSingleOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                "Billing coordinator displayed is not the selected one.");

            CustomerProfileWindow.ClickCancel();
            Cleanup();
        }

        [TestMethod]
        public void BillingInvoicingManagementForTACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.CheckSingleOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                "Billing coordinator displayed is not the selected one.");

            CustomerProfileWindow.ClickCancel();
            Cleanup();
        }

        [TestMethod]
        public void BillingInvoiceDeliveryOptionsForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            Assert.IsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("EBilling"),
                "Customer have opted not to use e-billing");

            Cleanup();
        }

        [TestMethod]
        public void BillingInvoicingPreferencesTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.CheckConsolidatedInvoicing();
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Consolidated Invoicing"),
                "Consolidated invoicing option is not checked");

            CustomerProfileWindow.CheckOvertimeBilling();
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Overtime Billing"),
                " Over time billing option is not checked");

            Cleanup();
        }

        [TestMethod]
        public void BillingTaxExemptionsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickViewTaxExceptions();
            Assert.IsTrue(CustomerProfileWindow.VerifyExceptionStatusSetToAll(),
                "Exception status is not set to ALL. It is currently set to - " + Globals.Temp);
            CustomerProfileWindow.ClickExcemptionWindowCancel();

            Cleanup();
        }

        [TestMethod]
        public void TABillingTaxExemptionsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Simple);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickViewTaxExceptions();
            Assert.IsTrue(CustomerProfileWindow.VerifyExceptionStatusSetToAll(),
                "Exception status is not set to ALL. It is currently set to - " + Globals.Temp);
            CustomerProfileWindow.ClickExcemptionWindowCancel();

            Cleanup();
        }

        [TestMethod]
        public void SelectCollectionsOrganizationTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            Assert.IsTrue(CustomerProfileWindow.VerifyAgingThesholdsSectionDisplayed(),
                "Aging thresholds section not displayed as per design");

            CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
            CustomerProfileWindow.UncheckInvoicingOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyPrimaryOrgUnitDisplayed(),
                "Primary Organization is not displayed as selected");
            CustomerProfileWindow.ClickCancel();

            Cleanup();
        }

        [TestMethod]
        public void SelectNABSCollectionsOrganizationTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
            CustomerProfileWindow.UncheckInvoicingOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyPrimaryOrgUnitDisplayed(),
                "Primary Organization is not displayed as selected");
            CustomerProfileWindow.ClickCancel();

            Cleanup();
        }

        [TestMethod]
        public void SelectCollectionsCOITest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            CustomerProfileWindow.ClickPrintCOIButton();
            Playback.Wait(3000);
            CustomerProfileWindow.ClickPrintButton();
            Playback.Wait(3000);

            Assert.IsTrue(CustomerProfileWindow.VerifyPrintPreviewWindowDisplayed(),
                "Print Preview window is not displayed");
            CustomerProfileWindow.ClosePPWindow();
            Cleanup();
        }

        [TestMethod]
        public void ValidateCollectionsFinanceChargesTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            Assert.IsTrue(CustomerProfileWindow.SelectFinanceCharges("Yes"), "Radio buttons cannot be checked");

            Cleanup();
        }

        [TestMethod]
        public void ServicingTabCustomerServiceManagementTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

            Assert.IsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                "Account manager checkbox is enabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                "Assigned to branch is disabled");
            Assert.IsTrue(CustomerProfileWindow.SelectServiceCoordinator("Test EllisCSR"),
                "Service Coordinator is not enabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                "Authorized to order checkbox is disabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                "Alert user for servicing rules is disabled");
            Assert.IsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                "User is permitted to access e-commerce site");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                "Branch permitted check is enabled");

            CustomerProfileWindow.ClickCancel();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        public void NABSServicingTabCustomerServiceManagementTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                "Account manager checkbox is disabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                "Assigned to branch is disabled");
            Assert.IsTrue(CustomerProfileWindow.VerifyServiceCoordinatorEnabled(),
                "Service Coordinator is not enabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                "Authorized to order checkbox is disabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                "Alert user for servicing rules is disabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                "User is not permitted to access e-commerce site");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                "Branch permitted check is enabled");

            CustomerProfileWindow.ClickCancel();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        public void ServicingTabCommonRequirementsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Purchase Order Required"),
                "Purchase order required checkbox is diabled");
            Assert.IsFalse(CustomerProfileWindow.SetPOFormat("pdf"), "Not able to set Po format text");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Drug Test"),
                "Drug test required is not enabled");
            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Background Check"),
                "Background check checkbox is disabled");

            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        public void RAServicingTabCommonRequirementsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsRAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

            Assert.IsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Bahavioral Survey"),
                "Bahavioral Survey Required checkbox is diabled");
            
            CustomerProfileWindow.CloseCustomerProfileWindow();
            TitlebarActions.ClickClose(SearchWindow.GetCustomerSearchResultsWindowProperties());
            Cleanup();
        }

        [TestMethod]
        public void ServicingTabDocumentsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.Documents);

            CustomerProfileWindow.ClickAddFileButton();
            CustomerProfileWindow.FindCustomerDocumentOnFile();
            CustomerProfileWindow.ClickCloseOnCustomerDOF();
            CustomerProfileWindow.ClickCloseOnDOF();

            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(), "Quote status is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
            Assert.IsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                "Pricing Matrix Grid is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                "Quote History Grid is not displayed or is null");

            CustomerProfileWindow.CloseQuoteProfileWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerCreateNewQuoteTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.ClickCreateNewQuoteButton();
            CustomerProfileWindow.CloseWarningWindow();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("Test")))
            {
                CustomerProfileWindow.CreateNewQuote(datarow);

                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "Customer profile window is not displayed");
                break;
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerCopyQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.ClickCopyQuote();

            Playback.Wait(2000);
            CustomerProfileWindow.CloseWarningWindow(); 
            Assert.IsTrue(CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed());

            CustomerProfileWindow.ClickChangeQuoteCustomerName();
            CustomerProfileWindow.SelectNewCustomerNameAfterSearch("Test");

            Assert.IsTrue(CustomerProfileWindow.VerifyNoQuotingAllowedWindowDisplayed(), 
                "Verify quoteing allowed window not displayed for new customer by name - TEST");
            CustomerProfileWindow.CloseNoQuotingAllowedWindow();
           
            TitlebarActions.ClickClose(GlobalWindows.QuoteProfileWindow);
            CustomerProfileWindow.CloseUnsavedChangesWindow();

            Cleanup();
        }

        [TestMethod]
        public void NABSCustomerQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(), "Quote status is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
            Assert.IsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                "Pricing Matrix Grid is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                "Quote History Grid is not displayed or is null");

            CustomerProfileWindow.CloseQuoteProfileWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void NABSCustomerCreateNewQuoteTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewQuoteButton();
            CustomerProfileWindow.CloseWarningWindow();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("NABS")))
            {
                CustomerProfileWindow.CreateNewQuote(datarow);

                Assert.IsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "Customer profile window is not displayed");
                break;
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void NABSCustomerCopyQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.ClickCopyQuote();

            Playback.Wait(2000);
            CustomerProfileWindow.CloseWarningWindow();
            Assert.IsTrue(CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed());

            CustomerProfileWindow.ClickChangeQuoteCustomerName();
            CustomerProfileWindow.SelectNewCustomerNameAfterSearch("Test");

            Assert.IsTrue(CustomerProfileWindow.VerifyNoQuotingAllowedWindowDisplayed(),
                "Verify quoteing allowed window not displayed for new customer by name - TEST");
            CustomerProfileWindow.CloseNoQuotingAllowedWindow();

            TitlebarActions.ClickClose(GlobalWindows.QuoteProfileWindow);
            CustomerProfileWindow.CloseUnsavedChangesWindow();

            Cleanup();
        }

        [TestMethod]
        public void TACustomerQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(), "Quote status is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
            Assert.IsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                "Pricing Matrix Grid is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
            Assert.IsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                "Quote History Grid is not displayed or is null");

            CustomerProfileWindow.CloseQuoteProfileWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerVerifyCreateJobOrderWindowDisplayedTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();
            Assert.IsTrue(JobOrderWindow.VerifyCreateJobOrderWindowDisplayed(),
                "Job Order window is not displayed when clicked on Create job Order button");
        }

        [TestMethod]
        public void CustomerJobOrdersWorkflowTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);

            CustomerProfileWindow.EnterJobOrderSearchCriteria();
            CustomerProfileWindow.SelectFirstJobOrderFromResults();

            Assert.IsTrue(CustomerProfileWindow.VerifyBasicJobInfoDisplayed(), "Basic job info is not displayed");

            TitlebarActions.ClickClose((WinWindow) OpenJobOrder.JobOrderProfileWindowProperties());
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerJOAddNewJobSiteWorkflowTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
            CustomerProfileWindow.SelectJOSubTab(CCTabConstants.JobSitesTab);

            CustomerProfileWindow.ClickAddNewJobSiteButton();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateJobSite);
            foreach (var datarow in datarows)
            {
                CustomerProfileWindow.CreateNewJobSite(datarow);
                Assert.IsTrue(CustomerProfileWindow.VerifyNewJobSiteCreated(datarow, Globals.Temp),
                    "New Job site is not created");
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void TACustomerJOAddNewJobSiteWorkflowTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
            CustomerProfileWindow.SelectJOSubTab(CCTabConstants.JobSitesTab);

            CustomerProfileWindow.ClickAddNewJobSiteButton();

            Assert.IsTrue(CustomerProfileWindow.VerifyAddNewJobSiteWindowDispllayed(), "Add new job site is not displayed");
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void VerifyCustomerJOWorkersTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
            CustomerProfileWindow.SelectJOSubTab(CCTabConstants.Workers);

            Assert.IsTrue(CustomerProfileWindow.VerifyWorkersDisplayed(), "workers table is not displayed");
        }

        [TestMethod]
        public void CustomerCreateNewJobOrderTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();

            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);

            // Find Customer Window
            //var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows)
            {
                //Console.WriteLine(dataRow.ItemArray[24]);
                JobOrderWindow.EnterJobOrderData(dataRow);
                JobOrderWindow.ClickOnButton("Search");

                Playback.Wait(3000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnButton("Continue");
                //JobOrderWindow.ClickOnContinueBtn();

                // Find Quote Tab/Window
                Playback.Wait(3000);
                JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                //JobOrderFindQuoteWindow.ClickOnButton("GO");

                Playback.Wait(2000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnContinueBtn();

                // Enter Basic Job Order Details
                BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                BasicJobInformationWindow.ClickOnContinueBtn();

                // Enter Schedule And Additional Charges Details
                ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                // Enter Order Notes in Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                // Focus back to Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                //Enter data in Requirements window
                RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                RequirementsWindow.ClickOnButton("Continue >");
                Playback.Wait(3000);

                //Enter data in Pre-Qualifying Requirements Window
                PreQualifyingQuestionsWindow.ClickonSaveButton();
                PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                Playback.Wait(3000);
                PreQualifyingQuestionsWindow.HandleAlertWindow();

                Assert.IsTrue(PreQualifyingQuestionsWindow.HandleAlertWindow(), "Job order not saved successfully");
            }
        }

        [TestMethod]
        public void NABSCustomerCreateNewJobOrderTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();

            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);

            // Find Customer Window
            //var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows.Where(dataRow => dataRow.ItemArray[4].ToString().Contains("NABS")))
            {
                //Console.WriteLine(dataRow.ItemArray[24]);
                JobOrderWindow.EnterJobOrderData(dataRow);
                JobOrderWindow.ClickOnButton("Search");

                Playback.Wait(3000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnButton("Continue");
                //JobOrderWindow.ClickOnContinueBtn();

                // Find Quote Tab/Window
                Playback.Wait(3000);
                JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                //JobOrderFindQuoteWindow.ClickOnButton("GO");

                Playback.Wait(2000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnContinueBtn();

                // Enter Basic Job Order Details
                BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                BasicJobInformationWindow.ClickOnContinueBtn();

                // Enter Schedule And Additional Charges Details
                ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                // Enter Order Notes in Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                // Focus back to Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                //Enter data in Requirements window
                RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                RequirementsWindow.ClickOnButton("Continue >");
                Playback.Wait(3000);

                //Enter data in Pre-Qualifying Requirements Window
                PreQualifyingQuestionsWindow.ClickonSaveButton();
                PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                Playback.Wait(3000);
                PreQualifyingQuestionsWindow.HandleAlertWindow();

                Assert.IsTrue(PreQualifyingQuestionsWindow.HandleAlertWindow(), "Job order not saved successfully");
            }
        }

        [TestMethod]
        public void CustomerInvoicesTabTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoicesTab);

            CustomerProfileWindow.SearchAndSelectFirstInvoiceFromGrid();
            Assert.IsTrue(CustomerProfileWindow.VerifyCustomerInvoiceNumberDisplayed(Globals.Temp),
                "Invoice number is not displayed on the customer invoice window");

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.TransactionHistory);
            CustomerProfileWindow.SearchForTransactionHistory();

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.DispatchLineItemDetail);
            CustomerProfileWindow.ClickOnJobOrderNumberLink();
            Assert.IsTrue(CustomerProfileWindow.VerifyJobOrderWindowDisplayed());

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.InvoiceSummary);
            Assert.IsTrue(CustomerProfileWindow.VerifyBalanceDueDisplayed(), "Balance due is not displayed");

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.JobOrderDetails);
            CustomerProfileWindow.ClickOnJobOrderNumber();
            Assert.IsTrue(CustomerProfileWindow.VerifyJobOrderWindowDisplayed(), "Job order window is not displayed");

            TitlebarActions.ClickClose(GlobalWindows.JobOrderProfileWindow);

            if (CustomerProfileWindow.ApplyCcPaymentEnabled())
            {
                CustomerProfileWindow.ClickApplyCCPayment();
                Assert.IsTrue(CustomerProfileWindow.VerifyApplyCCPaymentWindowDisplayed(),
                    "Apply CC payment window is not displayed.");
                CustomerProfileWindow.CloseApplyCCPaymentWindow();

            }
           
            CustomerProfileWindow.ClickOnReprintOriginalButton();
            Assert.IsTrue(CustomerProfileWindow.VerifyPrintDialogWindowDisplayed(),
                "Print window is not displayed.");
            Actions.SendAltF4();
            //TitlebarActions.ClickClose(GlobalWindows.DialogWindow);

            CustomerProfileWindow.ClickOnPrintAdjustedButton();
            Assert.IsFalse(CustomerProfileWindow.VerifyEllisExceptionWindowDisplayed(),
                "Ellis Exception Window is displayed displayed for Print Adjusted action.");
            Assert.IsTrue(CustomerProfileWindow.VerifyPrintDialogWindowDisplayed(),
                "Print window is not displayed.");
            Actions.SendAltF4();
            //TitlebarActions.ClickClose(GlobalWindows.DialogWindow);

            CustomerProfileWindow.CloseCustomerInvoiceWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        public void CustomerInvoicesPaymentsAndCreditsTabTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.PaymentsTab);

            CustomerProfileWindow.SearchForAllInvoices();
            CustomerProfileWindow.VerifySearchResultsDisplayed();

            Assert.IsTrue(CustomerProfileWindow.VerifySearchResultsDisplayed(), "Search results are not displayed");

            CustomerProfileWindow.ClickClearButton();
            CustomerProfileWindow.ClickHistoryFilter();
            Assert.IsFalse(CustomerProfileWindow.VerifySearchResultsDisplayed(), "Search results are displayed");

            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        private void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}