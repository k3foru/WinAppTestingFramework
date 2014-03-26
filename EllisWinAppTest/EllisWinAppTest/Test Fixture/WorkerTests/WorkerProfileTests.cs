using System.Data;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using EllisWinAppTest.Windows.WorkerWindow;
using EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows;
using EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.WorkerTests
{
    [CodedUITest]
    public class WorkerProfileTests : AppContext
    {

        #region Initialize Tests
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        #endregion

        #region Summary Tab Tests

        [TestMethod]
        public void VerifyWorkerSummaryTab()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(5000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Summary Tab not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void ClickOnCloseBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Summary Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void ClickOnChangeStatusBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Summary Tab not Displayed");
            WorkerSummaryWindow.ClickOnChangeStatusBtn();
            Assert.IsTrue(WorkerChangeStatusWindow.VerifyChangeStatusWindowDisplayed(), "Change Status Window Not Displayed");
            WorkerChangeStatusWindow.ClickOnCancelBtnStatusWindow();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSummaryWindow.VerifyAlertPopUpDisplayed(), "Alert Pop Up Not Displayed");
            WorkerSummaryWindow.CloseAlertPopUp();
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Workers Landing Page Not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void ClickOnPrintReportBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Summary Tab not Displayed");
            WorkerSummaryWindow.ClickOnPrintReportBtn();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClosePrintWindow();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void VerifyDropDownValuesChangeStatusWindow()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Summary Tab Not Displayed");
            WorkerSummaryWindow.ClickOnChangeStatusBtn();
            Assert.IsTrue(WorkerChangeStatusWindow.VerifyChangeStatusWindowDisplayed(), "Change Status Window Not Displayed");
            WorkerChangeStatusWindow.ClickOnStatusDropDown();
            Playback.Wait(2000);
            WorkerChangeStatusWindow.ClickOnCancelBtnStatusWindow();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSummaryWindow.VerifyAlertPopUpDisplayed(), "Alert Pop Up Not Displayed");
            WorkerSummaryWindow.CloseAlertPopUp();
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Workers Landing Page Not Displayed");
            Cleanup();
        }

        #endregion

        #region Profile Details Tab Tests

        [TestMethod]
        public void VerifyWorkerProfileFromActiveWorkers()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(5000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerProfileTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Playback.Wait(5000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerIdentityTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Identity Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerIdentity);

            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterdataInIdentity(datarow);

            }
            WorkerProfileDetailsWindow.ClickOnSaveBtnIdentity();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Identity Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            Cleanup();

        }

        [TestMethod]
        public void VerifyWorkerContactsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails, "Contacts");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Contacts Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerContacts);
            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterDataInContactTabs(datarow);
            }
            WorkerProfileDetailsWindow.ClickOnSaveBtnContacts();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Contacts Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerAddressTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails, "Addresses");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Addresses Tab not Displayed");

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAddress);
            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterDataInAddressTab(datarow);
            }
            WorkerProfileDetailsWindow.ClickOnSaveBtnAddress();
            Assert.IsTrue(WorkerGeoCodeWindow.VerifyGeoCodeWindowDisplayed(),"Geo Code Window Not Displayed");
            WorkerGeoCodeWindow.ClickOnOkBtn();
            Assert.IsTrue(WorkerVertexGeoCodeWindow.VerifyWorkerVertexGeoCodeWindowDisplayed()," Worker Vertex Geo Code Window Not Displayed");
            WorkerVertexGeoCodeWindow.ClickOnOkBtn();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Addresses Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            
            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerTempToHireTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                "Temp-to-Hire");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Temp-to-Hire Tab not Displayed");
            WorkerProfileDetailsWindow.ClickonSearchBtnTemp();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyCustomerSearchPopUpDisplayed(),"Customer Search PopUp not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerTemp);
            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterDatainCustomerSearch(datarow);
                WorkerProfileDetailsWindow.ClickonSearchBtnCustomer();
                Playback.Wait(2000);
                WorkerProfileDetailsWindow.SelectCustomer();
                WorkerProfileDetailsWindow.EnterdataInTemp(datarow);
            }

            WorkerProfileDetailsWindow.ClickonAddBtnTemp();
            WorkerProfileDetailsWindow.ClickonSaveBtnTemp();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Temp-to-Hire Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerAvailabilityTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                "Availability");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                "Workers Profile Availability Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerTemp);
            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterdataInAvailability(datarow);
                WorkerProfileDetailsWindow.ClickonAddBtnAvail();
            }
            WorkerProfileDetailsWindow.ClickonSaveBtnAvail();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Availability Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerVerificationTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),"Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                "Verification");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Verification Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerVerification);
            foreach (var datarow in datarows)
            {
                WorkerProfileDetailsWindow.EnterDatainVerification(datarow);
            }
            WorkerProfileDetailsWindow.ClickOnSaveBtnVerification();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Playback.Wait(2000);
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Verification Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerPaymentMethodTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                "Payment Method");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Payment Method Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerVerification);

            foreach (var datarow in datarows)
            {
                SelectRadioButton.Selection(datarow.ItemArray[18].ToString());
                WorkerProfileDetailsWindow.ClickOnEditBankDetailsBtn();
                Assert.IsTrue(WorkerProfileDetailsWindow.VerifyBankPopUpDisplayed(),"Bank PopUp Not Displayed");
                WorkerProfileDetailsWindow.EnterDataInBankPopUp(datarow);
            }
            WorkerProfileDetailsWindow.ClickOnSaveandCloseBtnBankPopUp();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyBankConfirmationPopUpDisplayed(),"Bank Confirmation Pop Up Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBankConfirmation();
            WorkerProfileDetailsWindow.ClickOnSaveBtnPayment();
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Payment Method Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerStatusTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Tab not Displayed");
            WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails, "Status");
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Status Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Withholdings Tab Tests

        [TestMethod]
        public void VerifyWorkerWithholdingsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Withholdings);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Withholdings Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerWithholdings);
            foreach (var datarow in datarows)
            {
                WorkerWithHoldingsWindow.EnterDataInWithholdings(datarow);
            }
            WorkerWithHoldingsWindow.ClickOnSaveBtnWithholdings();
            Playback.Wait(2000);
            WorkerWithHoldingsWindow.ClickOnSaveBtnWithholdings();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Withholdings Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Garnishments Tab Tests

        [TestMethod]
        public void VerifyWorkerTransactionHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(4);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Garnishment Tab not Displayed");
            WorkerGarnishmentsWindow.SelectTransactionHistoryTab();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Garnishment Select Transaction History Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerGarnishment);
            foreach (var datarow in datarows)
            {
                WorkerGarnishmentsWindow.EnterDataInTransactionHistoryTab(datarow);
            }
            WorkerGarnishmentsWindow.ClickonGoBtn();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerExistingOrdersTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(4);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Garnishment Existing Orders Tab not Displayed");
            Playback.Wait(2000);

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerGarnishment);
            foreach (var datarow in datarows)
            {
                WorkerGarnishmentsWindow.SelectDataInComboBox(datarow);
            }

            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Skills Tab Tests

        [TestMethod]
        public void VerifyWorkerJobSkillsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(5);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void AddSkillsForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(5);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            WorkerSkillsWindow.ClickOnAddorUpdateBtn();
            Assert.IsTrue(WorkerSkillsWindow.VerifyAddWorkerSkillsWindowDisplayed(),"Add Worker Skills Window Not Displayed");

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                WorkerSkillsWindow.EnterDataAddWorkerSkills(datarow);
            }
            WorkerSkillsWindow.ClickOnAddSelectedBtn();
            Playback.Wait(2000);
            WorkerSkillsWindow.ClickOnSaveBtn();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            WorkerSkillsWindow.SelectDutyChkBox();
            WorkerSkillsWindow.SelectVehicleChkBox();
            WorkerSkillsWindow.ClickonSaveBtninSkillsTab();
            Playback.Wait(3000);
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerSkillsWindow.ClickonOkBtn();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void AddLicensesForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(5);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datrow in datarows)
            {
                WorkerSkillsWindow.EnterLicenseData(datrow);
            }
            WorkerSkillsWindow.SelectDutyChkBox();
            WorkerSkillsWindow.SelectVehicleChkBox();
            WorkerSkillsWindow.ClickonSaveBtninSkillsTab();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
            WorkerSkillsWindow.ClickonOkBtn();
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Skills Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Work and Payment History Tab Tests

        [TestMethod]
        public void VerifyWorkerWorkHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(6);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Work and Payment History Tab not Displayed");
            WorkerWorkandPaymentHistoryWindow.SelectWorkHistoryInDropDown();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerPayment);
            foreach (var datarow in datarows)
            {
                WorkerWorkandPaymentHistoryWindow.EnterDataInWorkerHistoryTab(datarow);
            }
            WorkerWorkandPaymentHistoryWindow.ClickOnGoBtn();
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void VerifyWorkerPaymentHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(6);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Work and Payment History Tab not Displayed");
            WorkerWorkandPaymentHistoryWindow.SelectPaymentHistoryInDropDown();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerPayment);
            foreach (var datarow in datarows)
            {
                WorkerWorkandPaymentHistoryWindow.EnterDataInPaymentHistoryTab(datarow);
            }
            WorkerWorkandPaymentHistoryWindow.ClickOnGoBtnPayment();
            //WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Ratings and Notes Tab Tests

        [TestMethod]
        public void VerifyWorkerRatingsandNotesTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(7);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Ratings and Notes Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void AddRatingsForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(7);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Ratings and Notes Tab not Displayed");
            WorkerRatingsandNotesWindow.SelectRatingsInComboBox();
            WorkerRatingsandNotesWindow.ClickAddRatingsBtn();
            Assert.IsTrue(WorkerRatingsandNotesWindow.VerifyRatingsWindowDisplayed(),"Ratings Window Not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                WorkerRatingsandNotesWindow.EnterdataInRatingsWindow(datarow);
            }
            WorkerRatingsandNotesWindow.ClickSubmitRatings();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void AddNotesForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(7);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Ratings and Notes Tab not Displayed");
            WorkerRatingsandNotesWindow.SelectNotesInComboBox();
            WorkerRatingsandNotesWindow.ClickAddNotesBtn();
            Assert.IsTrue(WorkerRatingsandNotesWindow.VerifyNotesWindowDisplayed(),"Notes Window Not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                WorkerRatingsandNotesWindow.EnterdataInNotesWindow(datarow);
            }
            WorkerRatingsandNotesWindow.ClickSubmitnotes();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void BrowseForCustomerInRatingsWindow()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(7);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Ratings and Notes Tab not Displayed");
            WorkerRatingsandNotesWindow.SelectRatingsInComboBox();
            WorkerRatingsandNotesWindow.ClickAddRatingsBtn();
            Assert.IsTrue(WorkerRatingsandNotesWindow.VerifyRatingsWindowDisplayed(),"Ratings Window Not Displayed");
            WorkerRatingsandNotesWindow.ClickOnCustomerBrowseBtn();
            Assert.IsTrue(WorkerRatingsandNotesWindow.VerifyCustomerSearchWindowDisplayed(),"Customer Search Window Not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                WorkerRatingsandNotesWindow.EnterDataCustomerSearchWindow(datarow);
            }
            WorkerRatingsandNotesWindow.ClickOnCustomerNoBtn();
            Playback.Wait(2000);
            WorkerRatingsandNotesWindow.ClickOnCloseBtn();
            WorkerRatingsandNotesWindow.ClickCancelRatings();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void BrowseForCustomerInNotesWindow()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(7);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Ratings and Notes Tab not Displayed");
            WorkerRatingsandNotesWindow.SelectNotesInComboBox();
            WorkerRatingsandNotesWindow.ClickAddNotesBtn();
            Assert.IsTrue(WorkerRatingsandNotesWindow.VerifyCustomerSearchWindowDisplayed(),"Customer Search Window Not Displayed");
            WorkerRatingsandNotesWindow.ClickOnBrowseBtn();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                WorkerRatingsandNotesWindow.EnterDataCustomerSearchWindow(datarow);
            }
            WorkerRatingsandNotesWindow.ClickOnCustomerNoBtn();
            Playback.Wait(2000);
            WorkerRatingsandNotesWindow.ClickOnCloseBtn();
            WorkerRatingsandNotesWindow.ClickCancelnotes();
            Playback.Wait(2000);
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        #endregion

        #region Survey Tab Tests

        [TestMethod]
        public void VerifyWorkerSurveyTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            WorkerSummaryWindow.SelectWorkerFromTable();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Survey);
            Playback.Wait(2000);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Survey Tab not Displayed");
            WorkerSummaryWindow.ClickOnCloseBtn();
            Assert.IsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");

            Cleanup();
        }

        [TestMethod]
        public void RaChangeWorkerStatus()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsRAUser();

            SearchWindow.SelectSearchElements(null, "Worker", SearchWindow.SearchTypeConstants.Advanced);
            WorkerAdvancedSearchWindow.EnterWorkerNameAsSearchData("TEST");
            WorkerAdvancedSearchWindow.ClickSearchBtn();
            WorkerAdvancedSearchWindow.SelectWorkerfromSearchResults();
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Window not Displayed");
            WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Survey);
            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Survey not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAdvancedSearch);
            foreach (var datarow in datarows)
            {
                WorkerSurveyWindow.EnterDatainSsn(datarow);
                WorkerSurveyWindow.ClickOnSearchBtn();
                WorkerSurveyWindow.EnterDataInSurveyGrid();
                WorkerSurveyWindow.ClickOnUpdateBtn();
                WorkerSurveyWindow.EnterNotes(datarow);
                WorkerSurveyWindow.ClickOnSaveandCLoseBtn();
            }

            Assert.IsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Worker Profile Window Not Displayed");

            Cleanup();

        }

        #endregion

        #region CleanUp Test

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

        #endregion
    }
}