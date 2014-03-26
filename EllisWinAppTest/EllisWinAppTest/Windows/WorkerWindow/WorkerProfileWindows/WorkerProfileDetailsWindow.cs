using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerProfileDetailsWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        private static UITestControl GetWorkerProfilePopUpProperties()
        {
            var workerPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile" });
            return workerPopUp;
        }

        private static UITestControl GetBankDetailsPopUpProperties()
        {
            var bankPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Bank Details-" + Globals.WorkerName });
            return bankPopUp;
        }

        private static UITestControl GetCustomerSearchPopUpProperties()
        {
            var customerPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Customer Search" });
            return customerPopUp;
        }

        private static UITestControl GetBankConfirmationPopUpProperties()
        {
            var bankConfirmation = App.Container.SearchFor<WinWindow>(new { Name = "Bank Details" });
            return bankConfirmation;
        }

        #endregion

        #region Bank Confirmation PopUp Methods

        public static bool VerifyBankConfirmationPopUpDisplayed()
        {
            var bankConfirmation = GetBankConfirmationPopUpProperties();
            if (bankConfirmation.Exists)
            {
                return true;
            }
            return false;
        }

        public static bool ClickOnOkBankConfirmation()
        {
            var bankConfirmation = GetBankConfirmationPopUpProperties();
            if (bankConfirmation.Exists)
            {
                MouseActions.ClickButton(bankConfirmation, "_OKButton");
                return true;
            }
            return false;
        }

        #endregion

        #region Customer Search PopUp Methods

        public static bool VerifyCustomerSearchPopUpDisplayed()
        {
            var customerPopUp = GetCustomerSearchPopUpProperties();
            if (customerPopUp.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool ClickonSearchBtnCustomer()
        {
            var customerPopUp = GetCustomerSearchPopUpProperties();
            if (customerPopUp.Exists)
            {
                MouseActions.ClickButton(customerPopUp,CustomerSearchConstants.SearchBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonCloseBtnCustomer()
        {
            var customerPopUp = GetCustomerSearchPopUpProperties();
            if (customerPopUp.Exists)
            {
                MouseActions.ClickButton(customerPopUp, CustomerSearchConstants.CloseBtn);
                return true;
            }
            return false;
        }

        public static bool EnterDatainCustomerSearch(DataRow data)
        {
            var customerPopUp = GetCustomerSearchPopUpProperties();
            if (customerPopUp.Exists)
            {
                var custmerNumber = Actions.GetWindowChild(customerPopUp, CustomerSearchConstants.CustomerNumber);
                Actions.SetText(custmerNumber,data.ItemArray[3].ToString());
                return true;
            }
            return false;
        }

        public static bool SelectCustomer()
        {
            var customerPopUp = GetCustomerSearchPopUpProperties();
            if (customerPopUp.Exists)
            {
                var cell = TableActions.SelectCellFromTable(customerPopUp, CustomerSearchConstants.CustomersGrid,
                    "CustomerDomain row 1", "Customer Name - Number");
                cell.SetFocus();
                MouseActions.DoubleClick(cell);
                return true;
            }
            return false;
        }

        #endregion

        #region Worker PopUp Methods

        public static bool ClickOnOkBtnPopUp()
        {
            var workerPopUp = GetWorkerProfilePopUpProperties();
            if (workerPopUp.Exists)
            {
                var okBtn = Actions.GetWindowChild(workerPopUp, "_OKButton");
                Mouse.Click(okBtn);
                return true;
            }
            return false;

        }

        public static bool VerifyWorkerPopUpDisplayed()
        {
            var workerPopUp = GetWorkerProfilePopUpProperties();
            if (workerPopUp.Enabled)
            {
                return true;
            }
            return false;
        }


        #endregion

        #region Bank PopUp Methods

        public static bool VerifyBankPopUpDisplayed()
        {
            var bankPopUp = GetBankDetailsPopUpProperties();
            if (bankPopUp.Enabled)
            {
                return true;
            }
            return false;
        }

        public static void EnterDataInBankPopUp(DataRow data)
        {
            var bankPopUp = GetBankDetailsPopUpProperties();

            var bankName = Actions.GetWindowChild(bankPopUp, WorkerBankDetailsPopUpConstants.BankName);
            Actions.SetText(bankName, data.ItemArray[14].ToString());

            //SelectAccountType(data.ItemArray[17].ToString());

            var rtn = Actions.GetWindowChild(bankPopUp, WorkerBankDetailsPopUpConstants.Rtn);
            Actions.SetText(rtn, data.ItemArray[15].ToString());

            var accountNumber = Actions.GetWindowChild(bankPopUp, WorkerBankDetailsPopUpConstants.AccountNumber);
            Actions.SetText(accountNumber, data.ItemArray[16].ToString());
        }

        public static bool ClickOnCanelBtnBankPopUp()
        {
            var bankPopUp = GetBankDetailsPopUpProperties();
            if (bankPopUp.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(bankPopUp, WorkerBankDetailsPopUpConstants.Cancel);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;

        }

        public static bool ClickOnSaveandCloseBtnBankPopUp()
        {
            var bankPopUp = GetBankDetailsPopUpProperties();
            if (bankPopUp.Exists)
            {
                var saveBtn = Actions.GetWindowChild(bankPopUp, WorkerBankDetailsPopUpConstants.SaveandCloseBtn);
                Mouse.Click(saveBtn);
                return true;
            }
            return false;
        }

        public static void SelectAccountType(string data)
        {
            string control;
            var bankPopUp = GetBankDetailsPopUpProperties();
            switch (data)
            {
                case "Checking":
                    control = "Checking";
                    break;
                case "Savings":
                    control = "Savings";
                    break;
                default:
                    control = "Checking";
                    break;
            }
            Actions.SelectRadioButton(bankPopUp, control);
        }

        #endregion

        #region Identity Tab Methods

        public static void EnterdataInIdentity(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var suffix = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.Suffix);
            DropDownActions.SelectDropdownByText(suffix, data.ItemArray[3].ToString());

            var gender = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.Gender);
            DropDownActions.SelectDropdownByText(gender, data.ItemArray[4].ToString());

            var race = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.Race);
            DropDownActions.SelectDropdownByText(race, data.ItemArray[5].ToString());

            var mInitial = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.MiddleInitial);
            Actions.SetText(mInitial, data.ItemArray[6].ToString());

            var nName = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.NickName);
            Actions.SetText(nName, data.ItemArray[7].ToString());
        }

        public static bool ClickOnSaveBtnIdentity()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var saveBtn = Actions.GetWindowChild(workerProfileWindow, WorkerIdentityTabConstants.Save);
                Mouse.Click(saveBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Contacts Tab Methods

        public static void EnterDataInContactTabs(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var email = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.WorkerEmail);
            Actions.SetText(email, data.ItemArray[3].ToString());

            var pPhone = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.PrimaryPhoneNumber);
            pPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(pPhone, data.ItemArray[4].ToString());

            var pType = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.PrimaryContactType);
            DropDownActions.SelectDropdownByText(pType, data.ItemArray[6].ToString());

            var pTime = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.PrimaryContactTime);
            DropDownActions.SelectDropdownByText(pTime, data.ItemArray[7].ToString());

            var pExt = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.PrimaryExtension);
            Actions.SetText(pExt, data.ItemArray[5].ToString());

            var sPhone = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.SecondaryPhoneNumber);
            sPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(sPhone, data.ItemArray[8].ToString());

            var sType = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.MobileContactType);
            DropDownActions.SelectDropdownByText(sType, data.ItemArray[10].ToString());

            var sTime = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.MobileContactTime);
            DropDownActions.SelectDropdownByText(sTime, data.ItemArray[11].ToString());

            var sExt = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.MobileExtension);
            Actions.SetText(sExt, data.ItemArray[9].ToString());

            var ePhone = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.EmergencyContactNumber);
            ePhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(ePhone, data.ItemArray[12].ToString());

            var eName = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.EmergencyContactName);
            Actions.SetText(eName, data.ItemArray[14].ToString());

            var eRelation = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.ContactRelationship);
            DropDownActions.SelectDropdownByText(eRelation, data.ItemArray[15].ToString());

            var eExt = Actions.GetWindowChild(workerProfileWindow, WorkerContactsTabConstants.EmergencyExtension);
            Actions.SetText(eExt, data.ItemArray[13].ToString());

            
        }

        public static bool ClickOnSaveBtnContacts()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerContactsTabConstants.Save);
                return true;
            }
            return false;
        }

        #endregion

        #region Address Tab Methods

        public static void EnterDataInAddressTab(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            
            var rAddress = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.ResidenceAddressLine1);
            Actions.SetText(rAddress, data.ItemArray[3].ToString());

            var rSuiteNo = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.ResidenceAddressLine2);
            Actions.SetText(rSuiteNo, data.ItemArray[4].ToString());

            var rState = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.ResidenceAddressState);
            DropDownActions.SelectDropdownByText(rState, data.ItemArray[5].ToString());

            var rZip = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.ResidenceAddressZip);
            DropDownActions.SelectDropdownByText(rZip, data.ItemArray[6].ToString());

            var rCity = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.ResidenceAddressCity);
            DropDownActions.SelectDropdownByText(rCity, data.ItemArray[7].ToString());

            var chkBox = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.RcheckBox);
            Actions.SetCheckBox((WinCheckBox)chkBox,data.ItemArray[13].ToString());

            var mAddress = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.MailingAddressLine1);
            Actions.SetText(mAddress, data.ItemArray[8].ToString());

            var mSuiteNo = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.MailingAddressLine2);
            Actions.SetText(mSuiteNo, data.ItemArray[9].ToString());

            var mState = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.MailingAddressState);
            DropDownActions.SelectDropdownByText(mState, data.ItemArray[10].ToString());

            var mZip = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.MailingAddressZip);
            DropDownActions.SelectDropdownByText(mZip, data.ItemArray[11].ToString());

            var mCity = Actions.GetWindowChild(workerProfileWindow, WorkerAddressTabConstants.MailingAddressCity);
            DropDownActions.SelectDropdownByText(mCity, data.ItemArray[12].ToString());
        }

        public static bool ClickOnSaveBtnAddress()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerAddressTabConstants.Save);
                return true;
            }
            return false;
        }

        #endregion

        #region Temp-to-Hire Tab Methods

        public static bool ClickonSaveBtnTemp()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow,WorkerTempToHireTabConstants.Save);
                return true;
            }
            return false;
        }

        public static bool ClickonAddBtnTemp()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerTempToHireTabConstants.Add);
                return true;
            }
            return false;
        }

        public static bool ClickonSearchBtnTemp()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerTempToHireTabConstants.Browse);
                return true;
            }
            return false;
        }

        public static void EnterdataInTemp(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var startDate = Actions.GetWindowChild(workerProfileWindow, WorkerTempToHireTabConstants.StartDate1);
            DropDownActions.SelectDropdownByText(startDate, data.ItemArray[4].ToString());

            var endDate = Actions.GetWindowChild(workerProfileWindow, WorkerTempToHireTabConstants.EndDate1);
            DropDownActions.SelectDropdownByText(endDate, data.ItemArray[5].ToString());
        }

       
        #endregion

        #region Avilability Tab Methods

        public static bool ClickonSaveBtnAvail()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerAvailabilityTabConstants.Save);
                return true;
            }
            return false;
        }

        public static bool ClickonAddBtnAvail()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerAvailabilityTabConstants.Add);
                return true;
            }
            return false;
        }

        public static void EnterdataInAvailability(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var startDate = Actions.GetWindowChild(workerProfileWindow, WorkerAvailabilityTabConstants.FromDate);
            DropDownActions.SelectDropdownByText(startDate, data.ItemArray[4].ToString());

            var endDate = Actions.GetWindowChild(workerProfileWindow, WorkerAvailabilityTabConstants.ToDate);
            DropDownActions.SelectDropdownByText(endDate, data.ItemArray[5].ToString());

            var reason = Actions.GetWindowChild(workerProfileWindow,
                WorkerAvailabilityTabConstants.WorkerUnavailabilityReason);
            DropDownActions.SelectDropdownByText(reason,data.ItemArray[6].ToString());
        }

        #endregion

        #region Verification Tab Methods

        public static bool ClickOnSaveBtnVerification()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var saveBtn = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.Save);
                MouseActions.Click(saveBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnPrintBtnVerification()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var printBtn = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.Print);
                Mouse.Click(printBtn);
                return true;
            }
            return false;
        }

        public static void EnterDatainVerification(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var i9Date = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.I9CompletionDate);
            //var comboBox = i9Date.Container.SearchFor<WinComboBox>(new {Name = ""})
            i9Date.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[3].ToString());
            //DropDownActions.SelectDropdownByText(i9Date,data.ItemArray[3].ToString());


            var dOb = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.Dob);
            dOb.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[4].ToString());
            //DropDownActions.SelectDropdownByText(dOb, data.ItemArray[4].ToString());

            var cStatus = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.CitizenshipStatus);
            //DropDownActions.SelectDropdownByText(cStatus, data.ItemArray[5].ToString());
            cStatus.SetFocus();
            SendKeys.SendWait(data.ItemArray[5].ToString());

            var dTitle = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.DocumentTitle);
            //DropDownActions.SelectDropdownByText(dTitle, data.ItemArray[6].ToString());
            dTitle.SetFocus();
            Actions.SendText(data.ItemArray[6].ToString());

            var iAuthority = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.IssuingAuthority);
            Actions.SetText(iAuthority, data.ItemArray[7].ToString());

            var dNumber = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.DocumentNumber);
            Actions.SetText(dNumber, data.ItemArray[8].ToString());

            SendKeys.SendWait("{TAB}");
            var workerPopUp = GetWorkerProfilePopUpProperties();
            if (workerPopUp.Exists)
            {
                var oKBtn = Actions.GetWindowChild(workerPopUp, "_OKButton");
                Mouse.Click(oKBtn);
            }

            var eDate = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.ExpirationDate);
            //DropDownActions.SelectDropdownByText(eDate, data.ItemArray[9].ToString());
            eDate.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[9].ToString());

            var dTitleC = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.DocumentTitleListC);
            //DropDownActions.SelectDropdownByText(dTitleC, data.ItemArray[10].ToString());
            dTitleC.SetFocus();
            Actions.SendText(data.ItemArray[10].ToString());

            var iAuthorityC = Actions.GetWindowChild(workerProfileWindow,
                WorkerVerificationTabConstants.IssuingAuthorityListC);
            Actions.SetText(iAuthorityC, data.ItemArray[11].ToString());

            var dNumberC = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.DocumentNumberC);
            Actions.SetText(dNumberC, data.ItemArray[12].ToString());

            var eDateC = Actions.GetWindowChild(workerProfileWindow, WorkerVerificationTabConstants.ExpirationDateListC);
            //DropDownActions.SelectDropdownByText(eDateC, data.ItemArray[13].ToString());
            eDateC.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[13].ToString());
        }

        #endregion

        #region Payment Tab Methods

        public static bool ClickOnSaveBtnPayment()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var savBtn = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentMethodTabConstants.Save);
                Mouse.Click(savBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnEditBankDetailsBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var editBtn = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentMethodTabConstants.EditBankDetails);
                Mouse.Click(editBtn);
                return true;
            }
            return false;
        }

        public static void SelectPaymentMethod(string data)
        {
            string control;
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            switch (data)
            {
                case "Paycard":
                    control = "Paycard";
                    break;
                case "Direct Deposit":
                    control = "Direct Deposit";
                    break;
                case "Check":
                    control = "Check";
                    break;
                case "Voucher":
                    control = "Voucher";
                    break;
                default:
                    control = "Paycard";
                    break;
            }
            //Actions.SelectRadioButton(workerProfileWindow, control);

            var rbtControl = workerProfileWindow.Container.SearchFor<WinRadioButton>(new { Name = "Check" });
            var children = rbtControl.GetChildren();

            foreach (var child in children)
            {
                if (child.Name.Equals(control))
                {
                    child.SetProperty("Selected", true);
                }
            }
        }

        #endregion

        #region Controls

        public class WorkerIdentityTabConstants
        {
            public const string SSN = "mskSSN";
            public const string NickName = "txtNickName";
            public const string FirstName = "txtFirstName";
            public const string MiddleInitial = "txtMiddleInitial";
            public const string LastName = "txtLastName";
            public const string Gender = "cmbGender";
            public const string Suffix = "cmbSuffix";
            public const string Race = "cmbRace";
            public const string Save = "btnSaveIdentity";
        }

        public class WorkerContactsTabConstants
        {
            public const string WorkerEmail = "txtWorkerEmail";
            public const string EmergencyExtension = "txtEmergencyExtension";
            public const string PrimaryPhoneNumber = "mskPrimaryPhoneNumber";
            public const string SecondaryPhoneNumber = "mskMobilePhoneNumber";
            public const string EmergencyContactNumber = "mskEmergencyContactNumber";
            public const string EmergencyContactName = "txtEmergencyContactName";
            public const string PrimaryContactTime = "cmbPrimaryContactTime";
            public const string PrimaryContactType = "cmbPrimaryContactType";
            public const string ContactRelationship = "cmbContactRelationship";
            public const string PrimaryExtension = "txtPrimaryExtension";
            public const string PrimaryStatus = "txtPrimaryStatus";
            public const string PrimaryOptIn = "txtPrimaryOptIn";
            public const string PrimaryLocationTracking = "cmbPrimaryLocationTracking";
            public const string PrimaryJobAlert = "cmbPrimaryJobAlert";
            public const string MobileContactType = "cmbMobileContactType";
            public const string MobileContactTime = "cmbMobileContactTime";
            public const string MobileExtension = "txtMobileExtension";
            public const string Save = "btnSave";
        }

        public class WorkerAddressTabConstants
        {
            public const string ResidenceAddressLine1 = "_txtResidenceAddressLine1";
            public const string ResidenceAddressLine2 = "_txtResidenceAddressLine2";
            public const string ResidenceAddressState = "_cmbResidenceAddressState";
            public const string ResidenceAddressZip = "_cmbResidenceAddressZip";
            public const string ResidenceAddressCity = "_cmbResidenceAddressCity";
            public const string MailingAddressLine1 = "_txtMailingAddressLine1";
            public const string MailingAddressLine2 = "_txtMailingAddressLine2";
            public const string MailingAddressState = "_cmbMailingAddressState";
            public const string MailingAddressZip = "_cmbMailingAddressZip";
            public const string MailingAddressCity = "_cmbMailingAddressCity";
            public const string Save = "_btnSaveAddresses";
            public const string RcheckBox = "_chkbxMailingAddressSameAsResidence";
        }

        public class WorkerTempToHireTabConstants
        {
            public const string CompanyName = "txtCompanyName";
            public const string EndDate1 = "dteEndDate1";
            public const string StartDate1 = "dteStartDate1";
            public const string TempToHireGrid = "grdTemptoHire";
            public const string Save = "btnSaveHire";
            public const string Add = "btnAddToHire";
            public const string Browse = "btnCustomerSearch";
        }

        public class WorkerAvailabilityTabConstants
        {
            public const string Reason = "txtReason";
            public const string Miles = "txtMiles";
            public const string DontMessageTime = "txtDontMessageTime";
            public const string ToDate = "calToDate";
            public const string FromDate = "calFromDate";
            public const string WorkerUnavailabilityReason = "cmbWorkerUnavailabilityReason";
            public const string UnavailabiltyReasonsGrid = "grdUnavailabiltyReasons";
            public const string Save = "btnSaveAvailability";
            public const string Add = "btnAdd";
        }

        public class WorkerVerificationTabConstants
        {
            public const string EverifyDate = "dteEverifyDate";
            public const string EverifyStatus = "cmbEverifyStatus";
            public const string EverifyNumber = "txtEverifyNumber";
            public const string DelayReason = "txtDelayReason";
            public const string I9CompletionDate = "dteI9Completion";
            public const string MaidenName = "txtMaidenName";
            public const string Dob = "dteDOB";
            public const string CitizenshipStatus = "cmbStatus";
            public const string DocumentTitle = "cmbDocumentTitle";
            public const string IssuingAuthority = "txtIssuingAuthority";
            public const string DocumentNumber = "txtDocument";
            public const string ExpirationDate = "dteExpirationDate";
            public const string DocumentTitleListC = "cmbDocumentTitleListC";
            public const string IssuingAuthorityListC = "txtIssuingAuthorityListC";
            public const string DocumentNumberC = "txtDocumentListC";
            public const string ExpirationDateListC = "dteExpirationDateListC";
            public const string Save = "btnSaveVerify";
            public const string Print = "btnPrintI9";
        }

        public class WorkerPaymentMethodTabConstants
        {
            public const string Save = "btnSavePayment";
            public const string EditBankDetails = "btnEdit";
        }

        public class WorkerStatusTabConstants
        {
            public const string ProfileStatusDetailsGrid = "grdProfileStatusDetails";
        }

        public class WorkerBankDetailsPopUpConstants
        {
            public const string BankName = "txtBankName";
            public const string Rtn = "txtRTN";
            public const string AccountNumber = "txtAccountNumber";
            public const string SaveandCloseBtn = "btnSave";
            public const string Cancel = "btnCancel";
        }

        public class CustomerSearchConstants
        {
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string SearchBtn = "btnSearchCustomer";
            public const string CloseBtn = "btnClose";
            public const string CustomersGrid = "grdCustomers";
            
        }

        #endregion
    }
}