using System.Data;
using System.Reflection;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerWithholdings : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerWithholdingsWindowProperties()
        {
            var wWithholdingsWindow =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"});//{ Name = "Withholdings-" + Globals.WorkerName });
            return wWithholdingsWindow;
        }

        private static UITestControl GetW5PopUpProperties()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            var w5PopUp = wWithholdingsWindow.Container.SearchFor<WinWindow>(new { Name = "New Applicant" });
            return w5PopUp;
        }

        #endregion

        #region Withholdings Methods

        public static bool ClickOnContinueBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.ContinueBtn);
                Mouse.Click(continueBtn);
                return true;
            }
           return false; 
        }

        public static bool ClickOnRejectBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {
                var rejectBtn = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.RejectBtn);
                Mouse.Click(rejectBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnBackBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.BackBtn);
                Mouse.Click(backBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        public static bool EnterDataInWithholdings(DataRow data)
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {

                var lastName = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.LastName);
                lastName.SetFocus();
                SendKeys.SendWait(data.ItemArray[56].ToString());

                var martialStatus = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.MaritalStatus);
                martialStatus.SetFocus();
                SendKeys.SendWait(data.ItemArray[57].ToString());
                SendKeys.SendWait("{TAB}");

                for (int i = 0; i < 10; i++)
                {

                    if (martialStatus.Name.Equals(data.ItemArray[57].ToString()))
                    {
                        i = 101;
                    }
                    else
                    {
                        martialStatus.SetFocus();
                        SendKeys.SendWait(data.ItemArray[57].ToString());
                        SendKeys.SendWait("{TAB}");
                    }

                }


                var allowances = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.Allowances);
                allowances.SetFocus();
                SendKeys.SendWait("^(A)");
                SendKeys.SendWait("{DEL}");
                Actions.SetText(allowances, data.ItemArray[58].ToString());

                var additional = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.AdditionalAmount);
                additional.SetFocus();
                SendKeys.SendWait("^(A)");
                SendKeys.SendWait("{DEL}");
                Actions.SetText(additional, data.ItemArray[59].ToString());

                var exempt = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.Exempt);
                exempt.SetFocus();
                SendKeys.SendWait(data.ItemArray[60].ToString());

                var w5ChkBox = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.W5FormRadio);
                Actions.SetCheckBox((WinCheckBox) w5ChkBox, data.ItemArray[64].ToString());

                var w5PopUp = GetW5PopUpProperties();
                if (w5PopUp.Exists)
                {
                    var eligible = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Eligible);
                    eligible.SetFocus();
                    SendKeys.SendWait(data.ItemArray[61].ToString());

                    var filingStatus = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Status);
                    filingStatus.SetFocus();
                    SendKeys.SendWait(data.ItemArray[62].ToString());

                    var w5Spouse = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.SpouseW5);
                    w5Spouse.SetFocus();
                    SendKeys.SendWait(data.ItemArray[63].ToString());

                    MouseActions.ClickButton(w5PopUp, Ww5PopupConstants.DoneBtn);
                }

                var wTocChkBox = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.WotcChk);
                Actions.SetCheckBox((WinCheckBox) wTocChkBox, data.ItemArray[65].ToString());

                return true;
            }
            return false;
        }


        #endregion

        #region W5 PopUp Methods

        public static bool ClickOnCancelBtnW5()
        {
            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
          return false;  
        }

        public static bool ClickOnDoneBtn()
        {
            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
            {
                var doneBtn = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.DoneBtn);
                Mouse.Click(doneBtn);
                return true;
            }
            return false;
        }


        #endregion

        #region Controls

        private class WWithHoldingsConstants
        {
            public const string Allowances = "txtAllowances (Federal W-4 Line 5)";
            public const string AdditionalAmount = "txtAdditional Amount to WH (Federal W-4 Line 6)";
            public const string LastName = "cmbLast Name different than SSN Card? (Federal W-4 Line 4)";
            public const string MaritalStatus = "cmbMarital Status (Federal W-4 Line 3)";
            public const string Exempt = "cmbExempt (Federal W-4 Line 7)";
            public const string W5FormRadio = "chkW5Form";
            public const string WotcChk = "chkWOTCForm";
            public const string BackBtn = "_btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
        }

        private class Ww5PopupConstants
        {
            public const string Eligible = "cmbEligible (Federal W-5 Line 1)";
            public const string Status = "cmbFiling status (Federal W-5 Line 2)";
            public const string SpouseW5 = "cmbMarried and spouse has filed Form W-5 with employer (Federal W-5 Line 3)";
            public const string DoneBtn = "btnSave";
            public const string CancelBtn = "btnCancel";
        }

        #endregion
    }
}