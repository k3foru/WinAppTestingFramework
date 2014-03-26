using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerSurveyWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        private static UITestControl GetNotesPopUpProperties()
        {
            var notesPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Notes Popup" });
            return notesPopUp;
        }

        #endregion

        #region Survey Tab Methods

        public static bool VerifyWorkerProfileWindowDisplayed()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool EnterDatainSsn(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var ssn = Actions.GetWindowChild(workerProfileWindow, WorkerSurveyTabConstants.SSN);
                Actions.SetText(ssn, data.ItemArray[14].ToString());
                return true;
            }

            return false;
        }

        public static bool ClickOnSearchBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, "btnSearch");
                return true;
            }
            return false;
        }

        public static bool ClickOnUpdateBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, "btnUpdate");
                return true;
            }
            return false;
        }

        public static void EnterDataInSurveyGrid()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var select = TableActions.SelectCellFromTable(workerProfileWindow, WorkerSurveyTabConstants.SurveyGrid,
                "Band 0 row 1", "Select");
            select.SetFocus();
            Mouse.Click(select);
            SendKeys.SendWait("{TAB}");
            Playback.Wait(2000);
            SendKeys.SendWait("{TAB}");
            Playback.Wait(2000);
            SendKeys.SendWait("{TAB}");
            Playback.Wait(2000);
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            Playback.Wait(2000);
        }
        #endregion

        #region Notes Pop Up Methods

        public static bool ClickOnSaveandCLoseBtn()
        {
            var notesPopUp = GetNotesPopUpProperties();
            if (notesPopUp.Exists)
            {
                MouseActions.ClickButton(notesPopUp, "btnSaveClose");
                return true;
            }
            return false;
        }

        public static bool EnterNotes(DataRow data)
        {
            var notesPopUp = GetNotesPopUpProperties();
            if (notesPopUp.Exists)
            {
                var notes = Actions.GetWindowChild(notesPopUp, "txtNotes");
                Actions.SetText(notes, data.ItemArray[15].ToString());
                return true;
            }
            return false;
        }

        public static bool VerifyNotesPopUpDisplayed()
        {
            var notesPopUp = GetNotesPopUpProperties();
            if (notesPopUp.Enabled)
            {
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class WorkerSurveyTabConstants
        {
            public const string SSN = "mskSSN";
            public const string SurveyGrid = "grdSurvey";
        }

        #endregion
    }
}