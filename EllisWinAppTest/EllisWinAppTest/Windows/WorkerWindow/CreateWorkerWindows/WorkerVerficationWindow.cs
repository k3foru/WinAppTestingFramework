using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerVerficationWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerVerificationWindowProperties()
        {
            var vWorkerWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Verification" });
            return vWorkerWindow;
        }

        #endregion

        #region Verification Methods

        public static bool EnterVerificationData(DataRow data)
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                var i9Date = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.I9CompletionDt);
                i9Date.SetFocus();
                SendKeys.SendWait(data.ItemArray[39].ToString());

                var maidenName = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.MaidenName);
                Actions.SetText(maidenName, data.ItemArray[40].ToString());

                var dOb = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.Dob);
                dOb.SetFocus();
                SendKeys.SendWait("^(A)");
                SendKeys.SendWait("{DEL}");
                SendKeys.SendWait(data.ItemArray[41].ToString());

                var aTitle = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.ATitle);
                aTitle.SetFocus();
                SendKeys.SendWait(data.ItemArray[48].ToString());

                var aAuthority = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.AAuthority);
                Actions.SetText(aAuthority, data.ItemArray[49].ToString());

                var aDocument = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.ADocument);
                Actions.SetText(aDocument, data.ItemArray[50].ToString());
                SendKeys.SendWait("{TAB}");
                Playback.Wait(2000);
                SendKeys.SendWait(" ");

                var aExpiration = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.AExpiryDt);
                aExpiration.SetFocus();
                SendKeys.SendWait(data.ItemArray[51].ToString());

                return true;
            }
            return false;
           
        }

        public static bool ClickOnContinueBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.ContinueBtn);
                Mouse.Click(continueBtn);
                SendKeys.SendWait(" ");
                Playback.Wait(2000);
                return true;
            }
            return false;
        }

        public static bool ClickOnRejectBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                var rejectBtn = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.RejectBtn);
                Mouse.Click(rejectBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
           return false; 
        }

        public static bool ClickOnBackBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.BackBtn);
                Mouse.Click(backBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class VWorkerConstants
        {
            public const string I9CompletionDt = "dteI9Completion";
            public const string MaidenName = "txtMaidenName";
            public const string Dob = "dteDOB";
            public const string StatusRadioBtn = "optListValues";
            public const string DocumentsRadioBtn = "optDocumentList";
            public const string ATitle = "cmbTitle";
            public const string AAuthority = "txtAuthority";
            public const string ADocument = "txtDocument";
            public const string AExpiryDt = "dteExpDate";
            public const string CTitle = "cmbListCDocumentTitle";
            public const string CAuthority = "txtListCIssuingAuthority";
            public const string CDocument = "txtListCDocumentNumber";
            public const string CExpiryDt = "dteListCExpireDate";
            public const string Alien = "txtAlienResident";
            public const string AlienAuthorized = "txtAlienAuthorized";
            public const string I94 = "txtAdmission";
            public const string WorkUntilDt = "dteWorkUntil";
            public const string CancelBtn = "btnCancel";
            public const string ContinueBtn = "btnContinue";
            public const string BackBtn = "btnBack";
            public const string RejectBtn = "btnReject";
        }

        #endregion
    }
}