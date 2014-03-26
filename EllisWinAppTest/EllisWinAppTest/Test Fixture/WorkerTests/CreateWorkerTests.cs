using System.Runtime.CompilerServices;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
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
    public class CreateWorkerTests : AppContext
    {
        #region Initialize Tests
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        #endregion

        #region Create Applicant Test Methods


        [TestMethod]
        public void CreateInactiveQualifiedWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {

                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.ClickonAddSelectedBtn();
                    WorkerJobSkillsWindow.ClickonSaveBtn();
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonContinueBtn();
                }
            }
            var confirmation = WorkerConfirmApplicantElgibiltyWindow.ClickOnYesBtn();
            Assert.IsTrue(confirmation, "Confirmation Not Displayed");
            Cleanup();
        }

        [TestMethod]
        public void WorkerMismatchedSsn()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Mismatched SSN")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnOverideBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.EnterDatainTescor();
                    var survey = WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonContinueBtnTescor();
                    Assert.IsTrue(survey, "Survey Window Not Displayed");
                }

            }
            Cleanup();
        }

        [TestMethod]
        public void PrintNonQualifiedWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Not Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnOverideBtn();
                    WorkerAlreadyExistWindow.ClickonContinueBtnTelephone();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    var survey = WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonPrintBtn();
                    Playback.Wait(1000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClosePrintWindow();

                    Assert.IsTrue(survey, "Survey Window Not Displayed");
                }
            }
            Cleanup();

        }

        [TestMethod]
        public void CloseNonQualifiedWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Not Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnOverideBtn();
                    WorkerAlreadyExistWindow.ClickonContinueBtnTelephone();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    var survey = WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonCloseBtn();

                    Assert.IsTrue(survey, "Survey Window Not Displayed");
                }
            }
            Cleanup();

        }
        #endregion

        #region Update Applicant Test Methods

        [TestMethod]
        public void UpdateWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Active Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    var update = WorkerAlreadyExistWindow.ClickOnUpdateProfileBtn();

                    Assert.IsTrue(update, " Update Profile Window Not Displayed");
                }
            }
            Cleanup();
        }

        #endregion

        #region Override Applicant Test Methods

        [TestMethod]
        public void OverrideWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Active Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    var overRide = WorkerAlreadyExistWindow.ClickOnOverideBtn();

                    Assert.IsTrue(overRide, "Override Window Not Displayed");
                }
            }
            Cleanup();
        }

        #endregion

        #region Cancel Applicant Test Methods

        [TestMethod]
        public void CancelWorkerAddressData()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnCancelBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();

                Assert.IsTrue(confirmation, "Confirmation Pop up Not Displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void CancelWorkerEmploymentEligibilityData()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnCancelBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();

                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void ClickBackBtnWorkerEmploymentEligibilityData()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnBackBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Window Not Displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void ClickCancelBtnAlertWindow()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnBackBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnCancelBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void CancelWorkerW5Screen()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    var withHoldings = WorkerWithholdings.ClickOnCancelBtnW5();
                    Assert.IsTrue(withHoldings, "WithHoldings Window not Displayed");

                }

            }

            Cleanup();
        }

        [TestMethod]
        public void CancelWorkerWithHoldingsScreen()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnCancelBtn();

                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");

            }

            Cleanup();
        }

        [TestMethod]
        public void CancelWorkerJobSkills()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.ClickonAddSelectedBtn();
                    WorkerJobSkillsWindow.ClickonSaveBtn();
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonBackBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Window Not Displayed");
            }

            Cleanup();
        }

        [TestMethod]
        public void CancelWorkerConfirmationScreen()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.ClickonAddSelectedBtn();
                    WorkerJobSkillsWindow.ClickonSaveBtn();
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonContinueBtn();
                }
                WorkerConfirmApplicantElgibiltyWindow.ClickOnNoBtn();
                WorkerIdentityWindow.ClickOnOkBtnPopUp();
                WorkerAddressWindow.EnterDataInRejectPopUp();
                WorkerAddressWindow.ClickOnBackBtnReject();
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");
            }

            Cleanup();
        }


        #endregion

        #region Reject Applicant Test Methods

        [TestMethod]
        public void RejectWorker()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnRejectBtn();
                }
                WorkerIdentityWindow.ClickOnOkBtnPopUp();
                WorkerAddressWindow.EnterDataInRejectPopUp();
                var address = WorkerAddressWindow.ClickOnDoneBtnReject();
                Assert.IsTrue(address, "Address Window Not Displayed");
            }
            Cleanup();
        }

        [TestMethod]
        public void RejectWorkerVerification()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnRejectBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");

            }
            Cleanup();
        }

        [TestMethod]
        public void RejectWorkerW5Screen()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnRejectBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");

            }

            Cleanup();
        }

        [TestMethod]
        public void RejectWorkerJobSkills()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.ClickonAddSelectedBtn();
                    WorkerJobSkillsWindow.ClickonSaveBtn();
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonRejectBtn();
                }
                var confirmation = WorkerIdentityWindow.ClickOnOkBtnPopUp();
                Assert.IsTrue(confirmation, "Confirmation Pop Up Not Displayed");
            }

            Cleanup();
        }


        [TestMethod]
        public void RejectWorkerConfirmationScreen()
        {
            Initialize();

            WorkerIdentityWindow.ClickOnCreateApplicant();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Inactive Qualified")
                {
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(2000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.ClickonAddSelectedBtn();
                    WorkerJobSkillsWindow.ClickonSaveBtn();
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonContinueBtn();
                }
                WorkerConfirmApplicantElgibiltyWindow.ClickOnNoBtn();
                WorkerIdentityWindow.ClickOnOkBtnPopUp();
                WorkerAddressWindow.EnterDataInRejectPopUp();
                var address = WorkerAddressWindow.ClickOnDoneBtnReject();
                Assert.IsTrue(address, "Address Window Not Displayed");
            }

            Cleanup();
        }

        #endregion

       
        #region Cleanup Tests

        public
            void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

        #endregion

    }
}
