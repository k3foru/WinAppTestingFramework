using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerAlreadyExistWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerAlreadyExistWindowProperties()
        {
            var existsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Already Exist" });
            return existsWindow;
        }

        private static UITestControl GetTelephoneOverideWindowProperties()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            var tOveride = existsWindow.Container.SearchFor<WinWindow>(new { Name = "Telephone Number Override" });
            return tOveride;
        }

        #endregion

        #region Worker Already Exists Methods

        public static bool ClickOnContinueBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.ContinueBtn);
                Mouse.Click(continueBtn);
                return true;
            }

            return false;
           
        }

        public static bool ClickOnOverideBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var overrideBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.OverrideBtn);
                Mouse.Click(overrideBtn);
                return true;
            }
            return false;

        }

        public static bool ClickOnUpdateProfileBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var updateProfileBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.UpdateProfileBtn);
                Mouse.Click(updateProfileBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnBackBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.BackBtn);
                Mouse.Click(backBtn);
                return true;
            }
            return false;

        }

        public static bool SelectWorkerFromGrid()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var row = TableActions.SelectRowFromTable(existsWindow, "grdWorkers", "WorkerInfoDomain row 1");
                Mouse.DoubleClick(row);
                return true;
            }
            return false;

        }

        public static bool ClickonContinueBtnTelephone()
        {
            var tOveride = GetTelephoneOverideWindowProperties();
            if (tOveride.Exists)
            {
                MouseActions.ClickButton(tOveride, "btnContinue");
                return true;
            }
            
            return false;
        }

        #endregion

        #region Controls

        private class WorkerexistsPopUpConstants
        {
            public const string WorkersGrid = "grdWorkers";
            public const string ContinueBtn = "btnContinue";
            public const string OverrideBtn = "btnOverride";
            public const string UpdateProfileBtn = "btnUpdateProfile";
            public const string BackBtn = "btnBack";
        }

        #endregion
    }
}