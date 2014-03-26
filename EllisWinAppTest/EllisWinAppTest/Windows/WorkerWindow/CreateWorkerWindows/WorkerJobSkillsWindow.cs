using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerJobSkillsWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerSkillsWindowProperties()
        {
            var jWorkerWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });//{ Name = "JobSkills-"+ Globals.WorkerName });
            return jWorkerWindow;
        }

        private static UITestControl GetAddWorkerSkillsWindowProperties()
        {
            var addWorkerSkillsWindow = App.Container.SearchFor<WinWindow>(new { Name = "Add Worker Skills" });
            return addWorkerSkillsWindow;
        }

        #endregion

        #region Job Skills Methods

        public static bool ClickonAddOrUpdateBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.AddOrUpdateBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonBackBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.BackBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonCancelBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.CancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonRejectBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.RejectBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonContinueBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.ContinueBtn);
                return true;
            }
            return false;
        }

        public static void EnterLicenseData(DataRow data)
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            var cell = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "Add Row", "License Type");
            cell.SetFocus();
            Mouse.DoubleClick(cell);
            SendKeys.SendWait(data.ItemArray[68].ToString());
            Actions.SendTab();
            Actions.SendText("{BACKSPACE}");
            SendKeys.SendWait(data.ItemArray[69].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[70].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[71].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[72].ToString());
            Playback.Wait(2000);

            var dutyChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.DutyChkBox);
            Actions.SetCheckBox((WinCheckBox)dutyChkBox, data.ItemArray[73].ToString());

            var vehicleChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.VehicleChkBox);
            Actions.SetCheckBox((WinCheckBox)vehicleChkBox, data.ItemArray[74].ToString());

        }


        #endregion

        #region Add Worker Skills Methods

        public static bool ClickonSearchBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.SearchBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonAddSelectedBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.AddSelectedBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonSaveBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.SaveBtn);
                return true;
            }
            return false;
        }

        public static bool EnterDataInAddSkills(DataRow data)
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                var postionFocus = Actions.GetWindowChild(addWorkerSkillsWindow,
                WorkerAddSkillsWindowConstants.PositionFocus);
                var cmbBox = (WinComboBox)postionFocus;
                DropDownActions.SelectDropdownByText(cmbBox, data.ItemArray[66].ToString());

                var checkBox = Actions.GetWindowChild(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.CheckBox);
                Actions.SetCheckBox((WinCheckBox)checkBox, data.ItemArray[67].ToString());

                return true;
            }
            return false;

        }

        #endregion

        #region Controls

        private class WJobSkillsWindowConstants
        {
            public const string AddOrUpdateBtn = "btnAddSkills";
            public const string BackBtn = "btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string DutyChkBox = "chkLightDutyWorkOnly";
            public const string VehicleChkBox = "chkVehicleAvailablity";
            public const string LicenseGrid = "grdLicenses";
            public const string CertificationsGrid = "grdCertifications";
            public const string BackgroundCheckGrid = "grdBackgroundCheck";
            public const string SkillsGrid = "ultraGrid1";

        }

        private class WorkerAddSkillsWindowConstants
        {
            public const string PositionFocus = "cmbPositionFocus";
            public const string Title = "txtSearchText";
            public const string SkillsExperienceGrid = "grdSkillsExperience";
            public const string SearchBtn = "btnSearch";
            public const string CheckBox = "chkSelect";
            public const string AddSelectedBtn = "btnAdd";
            public const string SaveBtn = "btnAddSkillsExperience";
        }

        #endregion

    }
}
