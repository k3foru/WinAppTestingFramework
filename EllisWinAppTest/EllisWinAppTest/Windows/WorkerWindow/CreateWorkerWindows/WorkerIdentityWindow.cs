using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Web; 

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerIdentityWindow : AppContext
    {
    
        #region Window Properties

        private static UITestControl GetCreateApplicantWindowProperties()
        {
            var applicantWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Identity"});
            return applicantWindow;
        }

        private static UITestControl DuplicateEmailAddressPopUpProperties()
        {
            var duplicateEmailAddressPopUp = App.Container.SearchFor<WinWindow>(new {Name = "Duplicate Email"});
            return duplicateEmailAddressPopUp;
        }

        private static UITestControl AlertPopUpProperties()
        {
            var popup = App.Container.SearchFor<WinWindow>(new {Name = "Alert"});
            return popup;
        }

        #endregion

        #region Identity Tab Methods

        public static void ClickOnCreateApplicant()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.File });
            var worker = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Worker });
            var applicant = worker.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.CApplicant });

            MouseActions.Click(file);
            MouseActions.Click(worker);
            MouseActions.Click(applicant);
        }

        public static bool ClickOnContinueBtn()
        {
            var applicantWindow = GetCreateApplicantWindowProperties();
            if (applicantWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(applicantWindow, IWorkerConstants.ContinueBtn);
                Mouse.Click(continueBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var applicantWindow = GetCreateApplicantWindowProperties();
            if (applicantWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(applicantWindow, IWorkerConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOkBtnDuplicate()
        {
            var duplicateEmailAddressPopUp = DuplicateEmailAddressPopUpProperties();
            if (duplicateEmailAddressPopUp.Exists)
            {
                var okBtn = Actions.GetWindowChild(duplicateEmailAddressPopUp, IWorkerConstants.OkBtn);
                Mouse.Click(okBtn);
                return true;
            }
            return false;
        }

        public static bool EnterApplicantData(DataRow data)
        {
            var applicantWindow = GetCreateApplicantWindowProperties();
            if (applicantWindow.Exists)
            {
                var firstName = Actions.GetWindowChild(applicantWindow, IWorkerConstants.FirstName);
                Actions.SetText(firstName, data.ItemArray[3].ToString());

                var middleInitial = Actions.GetWindowChild(applicantWindow, IWorkerConstants.MiddleInitial);
                Actions.SetText(middleInitial, data.ItemArray[4].ToString());

                var lastName = Actions.GetWindowChild(applicantWindow, IWorkerConstants.LastName);
                Actions.SetText(lastName, data.ItemArray[5].ToString());

                var ssn = Actions.GetWindowChild(applicantWindow, IWorkerConstants.SSN);
                Actions.SetText(ssn, data.ItemArray[6].ToString());

                var phone = Actions.GetWindowChild(applicantWindow, IWorkerConstants.PrimaryPhone);
                Actions.SetText(phone, data.ItemArray[7].ToString());

                var contactType = Actions.GetWindowChild(applicantWindow, IWorkerConstants.ContactType);
                DropDownActions.SelectDropdownByText(contactType, data.ItemArray[9].ToString());

                var email = Actions.GetWindowChild(applicantWindow, IWorkerConstants.Email);
                Actions.SetText(email, data.ItemArray[8].ToString());

                var laborReady = Actions.GetWindowChild(applicantWindow, IWorkerConstants.LaborReady);
                //DropDownActions.SelectDropdownByText(laborReady, data.ItemArray[10].ToString());
                laborReady.SetFocus();
                SendKeys.SendWait(data.ItemArray[10].ToString());

                return true;
            }
            return false;
           
        }

        
        #endregion

        #region Alert PopUp Methods

        public static bool ClickOnOkBtnPopUp()
        {
            var popup = AlertPopUpProperties();
            if (popup.Exists)
            {
                MouseActions.ClickButton(popup, "_OKButton");
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtnPopUp()
        {
            var popup = AlertPopUpProperties();
            if (popup.Exists)
            {
                MouseActions.ClickButton(popup, "_CancelButton");
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class IWorkerConstants
        {
            public const string FirstName = "txtFirstName";
            public const string MiddleInitial = "txtMiddleInitial";
            public const string LastName = "txtLastName";
            public const string SSN = "mskSSN";
            public const string PrimaryPhone = "mskPhone";
            public const string Email = "txtJobSiteEmail";
            public const string ContactType = "cmbMobileContactType";
            public const string LaborReady = "cmbLaborReady";
            public const string CancelBtn = "btnCancel";
            public const string ContinueBtn = "btnContinue";
            public const string OkBtn = "btnOk";
        }

        #endregion
    }
}

