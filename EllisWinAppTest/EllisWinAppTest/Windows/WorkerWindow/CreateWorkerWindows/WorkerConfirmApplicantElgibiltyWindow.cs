using System;
using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerConfirmApplicantElgibiltyWindow :AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerConfirmationWindowProperties()
        {
            var cWorkerWindow =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"}); //ClassName = "Confirmation-" + Globals.WorkerName });
            return cWorkerWindow;
        }

        #endregion

        #region Confirmation Methods

        public static bool ClickOnYesBtn()
        {
            var cWorkerWindow = GetWorkerConfirmationWindowProperties();
            if (cWorkerWindow.Exists)
            {
                var yesBtn = Actions.GetWindowChild(cWorkerWindow, ConfirmationConstants.YesBtn);
                Mouse.Click(yesBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnNoBtn()
        {
            var cWorkerWindow = GetWorkerConfirmationWindowProperties();
            if (cWorkerWindow.Exists)
            {
                var noBtn = Actions.GetWindowChild(cWorkerWindow, ConfirmationConstants.NoBtn);
                Mouse.Click(noBtn);
                return true;
            }
            return false;
            
        }

        #endregion

        #region Controls

        private class ConfirmationConstants
        {
            public const string YesBtn = "btnYes";
            public const string NoBtn = "btnNo";
        }

        #endregion
    }
}
