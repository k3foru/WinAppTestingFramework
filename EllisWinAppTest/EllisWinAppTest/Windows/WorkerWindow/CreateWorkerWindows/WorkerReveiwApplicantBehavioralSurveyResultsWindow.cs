using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerReveiwApplicantBehavioralSurveyResultsWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerSurveyWindowProperties()
        {
            var surveyWindow =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"});//Name = "Survey-" + Globals.WorkerName });
            return surveyWindow;
        }

        #endregion

        #region Review Bahvioral Survey results Methods

        public static bool ClickOnContinueBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(surveyWindow, ReviewConstants.ContinueBtn);
                Playback.Wait(3000);
                Mouse.Click(continueBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnBackBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(surveyWindow, ReviewConstants.BackBtn);
                Mouse.Click(backBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(surveyWindow, ReviewConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Tescor Methods

        public static bool ClickonContinueBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                MouseActions.ClickButton(surveyWindow, ReviewConstants.ContinueBtntescor);
                return true;
            }
            return false;
        }

        public static bool ClickonCancelBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                MouseActions.ClickButton(surveyWindow, ReviewConstants.CancelBtntescor);
                return true;
            }
            return false;
        }

        public static bool ClickonBackBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                MouseActions.ClickButton(surveyWindow, ReviewConstants.BackBtntescor);
                return true;
            }
            return false;
        }

        public static bool ClickonPrintBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                MouseActions.ClickButton(surveyWindow, ReviewConstants.PrintLetterBtn);
                return true;
            }
            return false;
               
        }

        public static void ClosePrintWindow()
        {
            SendKeys.SendWait("{ESC}");
        }

        public static bool ClickonCloseBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                MouseActions.ClickButton(surveyWindow, ReviewConstants.CloseBtn);
                return true;
            }
            return false;
        }

        public static bool EnterDatainTescor()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var tescorSsn = Actions.GetWindowChild(surveyWindow, ReviewConstants.TescorSsn);
                tescorSsn.SetFocus();
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{TAB}");

                var ssn = Actions.GetWindowChild(surveyWindow, ReviewConstants.Ssn);
                ssn.SetFocus();
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{TAB}");

                var chkBox = Actions.GetWindowChild(surveyWindow, ReviewConstants.DateFilterChkBx);
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");

                return true;
            }
            return false;
           
        }

        #endregion

        #region Controls

        private class ReviewConstants
        {
            public const string ContinueBtn = "btnContinue";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "_btnBack";
            public const string TescorSsn = "cmbTescorSSN";
            public const string Ssn = "cmbSSN";
            public const string SearchDate = "dtpSearchDate";
            public const string TryagainBtn = "_btnTryAgain";
            public const string DateFilterChkBx = "cbUseDateFilter";
            public const string BackBtntescor = "btnBack";
            public const string ContinueBtntescor = "btnContinue";
            public const string CancelBtntescor = "btnCancel";
            public const string PrintLetterBtn = "btnPrintLetter";
            public const string CloseBtn = "btnClose";

        }

        #endregion
    }
}
