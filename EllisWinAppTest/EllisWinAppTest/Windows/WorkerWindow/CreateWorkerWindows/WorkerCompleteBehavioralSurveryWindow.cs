using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerCompleteBehavioralSurveryWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerSurveyWindowProperties()
        {
            var surveyWindow =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"});
            return surveyWindow;
        }

        #endregion

        #region Complete Bahvioral Survey Methods

        public static bool ClickOnGetResultsBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var getResultsBtn = Actions.GetWindowChild(surveyWindow, SurveyConstants.GetResultsBtn);
                Mouse.Click(getResultsBtn);
                return true;
            }
            return false;
            
        }

        public static bool ClickOnBackBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(surveyWindow, SurveyConstants.BackBtn);
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
                var cancelBtn = Actions.GetWindowChild(surveyWindow, SurveyConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
           return false; 
        }

        public static bool ClickOnRejectBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
            {
                var rejectBtn = Actions.GetWindowChild(surveyWindow, SurveyConstants.RejectBtn);
                Mouse.Click(rejectBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class SurveyConstants
        {
            public const string RejectBtn = "btnReject";
            public const string GetResultsBtn = "btnGetResults";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "_btnBack";
        }

        #endregion
    }
}