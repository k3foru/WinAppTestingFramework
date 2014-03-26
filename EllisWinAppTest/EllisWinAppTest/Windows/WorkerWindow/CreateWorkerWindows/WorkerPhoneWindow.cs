using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerPhoneWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerPhoneWindowProperties()
        {
            var phoneWindow =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"});//{Name = "Phone-" + Globals.WorkerName});
            return phoneWindow;
        }

        #endregion

        #region Phone Window Methods

        public static bool EnterPhoneData(DataRow data)
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                var email = Actions.GetWindowChild(phoneWindow, PWorkerConstants.Email);
            Actions.SetText(email, data.ItemArray[8].ToString());

            var pPhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PPhoneNumber);
            pPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(pPhone, data.ItemArray[26].ToString());

            var pType = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PContactType);
            DropDownActions.SelectDropdownByText(pType, data.ItemArray[28].ToString());

            var pTime = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PContactTime);
            DropDownActions.SelectDropdownByText(pTime, data.ItemArray[29].ToString());

            var pExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PExt);
            Actions.SetText(pExt, data.ItemArray[27].ToString()); 

            var sPhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SPhoneNumber);
            sPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(sPhone, data.ItemArray[30].ToString());

            var sType = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SContactType);
            DropDownActions.SelectDropdownByText(sType, data.ItemArray[32].ToString());

            var sTime = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SContactTime);
            DropDownActions.SelectDropdownByText(sTime, data.ItemArray[33].ToString());

            var sExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SExt);
            Actions.SetText(sExt, data.ItemArray[31].ToString());

            var ePhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EPhoneNumber);
            ePhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(ePhone, data.ItemArray[34].ToString());

            var eName = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EContactName);
            Actions.SetText(eName, data.ItemArray[36].ToString());

            var eRelation = Actions.GetWindowChild(phoneWindow, PWorkerConstants.ContactRelationship);
            DropDownActions.SelectDropdownByText(eRelation, data.ItemArray[37].ToString());

            var eExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EExt);
            Actions.SetText(eExt, data.ItemArray[35].ToString());

            var chkBox = Actions.GetWindowChild(phoneWindow, PWorkerConstants.Application);
                Actions.SetCheckBox((WinCheckBox) chkBox, "TRUE");
                return true;
            }

            return false;

            
        }

       
        public static bool ClickOnContinueBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(phoneWindow, PWorkerConstants.ContinueBtn);
                MouseActions.Click(continueBtn);
                return true;
            }
           return false;
            
        }

        public static bool ClickOnRejectBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                var rejectBtn = Actions.GetWindowChild(phoneWindow, PWorkerConstants.RejectBtn);
                MouseActions.Click(rejectBtn);
                return true;
            }
           return false; 
        }

        public static bool ClickOnBackBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(phoneWindow, PWorkerConstants.BackBtn);
                MouseActions.Click(backBtn);
                return true;
            }
           return false; 
        }

        public static bool ClickOnCancelBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(phoneWindow, PWorkerConstants.CancelBtn);
                MouseActions.Click(cancelBtn);
                return true;
            }
           return false; 
        }

        #endregion

        #region Controls

        private class PWorkerConstants
        {
            public const string Email = "txtWorkerEmail";
            public const string PPhoneNumber = "mskPrimaryPhoneNumber";
            public const string PExt = "txtPrimaryExtension";
            public const string PContactType = "cmbPrimaryContactType";
            public const string PContactTime = "cmbPrimaryContactTime";
            public const string SPhoneNumber = "mskMobilePhoneNumber";
            public const string SExt = "txtMobileExtension";
            public const string SContactType = "cmbMobileContactType";
            public const string SContactTime = "cmbMobileContactTime";
            public const string EPhoneNumber = "mskEmergencyContactNumber";
            public const string EExt = "txtEmergencyExtension";
            public const string EContactName = "txtEmergencyContactName";
            public const string ContactRelationship = "cmbContactRelationship";
            public const string Application = "chkApplication";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "btnBack";
        }

        #endregion
    }
}