using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerGeoCodeWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerGeoCodeWindowProperties()
        {
            var geoCodeWindow = App.Container.SearchFor<WinWindow>(new {Name = "Geo Code"});
            return geoCodeWindow;
        }

        #endregion

        #region Geo Code Methods

        public static bool ClickOnOkBtn()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            if (geoCodeWindow.Exists)
            {
                var okBtn = Actions.GetWindowChild(geoCodeWindow, GeoCodeConstants.OkBtn);
                Mouse.Click(okBtn);
                return true;
            }
            return false;
        }

        public static bool VerifyGeoCodeWindowDisplayed()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            if (geoCodeWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            if (geoCodeWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(geoCodeWindow, GeoCodeConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class GeoCodeConstants
        {
            public const string OkBtn = "btnOK";
            public const string CancelBtn = "btnCancel";
            public const string GeoCodeGrid = "grdGeoCode";
            public const string GeoCodeRow = "VertexGeoCodeDomain row 1";
        }

      
        #endregion
    }
}