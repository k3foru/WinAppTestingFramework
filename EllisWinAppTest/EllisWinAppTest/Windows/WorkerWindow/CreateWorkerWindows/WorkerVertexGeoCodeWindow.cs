using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerVertexGeoCodeWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerVertexGeoCodeWindowProperties()
        {
            var geoCodeWindow = App.Container.SearchFor<WinWindow>(new { Name = "Geo Code" });
            var vertexGeoCodeWindow = geoCodeWindow.Container.SearchFor<WinWindow>(new { Name = "New Applicant" });
            return vertexGeoCodeWindow;
        }

        #endregion

        #region Vertex Geo Code Methods

        public static bool ClickOnOkBtn()
        {
            var vertexGeoCodeWindow = GetWorkerVertexGeoCodeWindowProperties();
            if (vertexGeoCodeWindow.Exists)
            {
                var okBtn = Actions.GetWindowChild(vertexGeoCodeWindow, VertexGeoCodeConstants.OkBtn);
                Mouse.Click(okBtn);
                return true;
            }
            return false;
        }

        public static bool VerifyWorkerVertexGeoCodeWindowDisplayed()
        {
            var vertexGeoCodeWindow = GetWorkerVertexGeoCodeWindowProperties();
            if (vertexGeoCodeWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        #endregion

        #region controls

        private class VertexGeoCodeConstants
        {
            public const string OkBtn = "_OKButton";
        }

        #endregion
    }
}