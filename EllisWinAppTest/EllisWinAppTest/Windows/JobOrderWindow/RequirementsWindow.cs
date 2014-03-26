using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class RequirementsWindow : AppContext
    {
        private static UITestControl RequirementsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Create New JobOrder"});
            //var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            return joborderWindow;
        }

        private static UITestControl AddSkillsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Add Skills"});
            return joborderWindow;
        }

        private static UITestControl AddSkillsInternalWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Add Skills"});
            var intwindow = joborderWindow.Container.SearchFor<WinWindow>(new {Name = "Results - Select to Add"});
            return intwindow;
        }

        private static UITestControlCollection GetAddSkillsDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = AddSkillsInternalWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetAddSkillsEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = AddSkillsInternalWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        public static void ClickOnButton(string btnName)
        {
            Factory.ClickOnButton(RequirementsWindowProperties(), btnName);
        }

        public static void EnterDatainRequirementsWindow(DataRow dataRow)
        {
            ClickOnButton("Add/Update Skills");
            var windowInst = AddSkillsWindowProperties();
            if (!string.IsNullOrEmpty(dataRow.ItemArray[79].ToString()))
            {
                var ddInstance = Actions.GetWindowChild(windowInst, "cmbPositionFocus");
                DropDownActions.SelectDropdownByText(ddInstance, dataRow.ItemArray[79].ToString());
            }

            if (!string.IsNullOrEmpty(dataRow.ItemArray[79].ToString()))
            {
                var txtBoxInstance = Actions.GetWindowChild(windowInst, "txtSearchText");
                Actions.SetText(txtBoxInstance, dataRow.ItemArray[80].ToString());
                //DropDownActions.SelectDropdownByText(ddInstance, dataRow.ItemArray[80].ToString());
            }

            ClickOnButton("Search");

            //Playback.Wait(3000);
            //SendKeys.SendWait("{TAB}");
            //SendKeys.SendWait(" ");

            var chkBoxControl = Actions.GetWindowChild(windowInst, "chkSelect");
            //var chkbox = chkBoxControl.Container.SearchFor<WinCheckBox>(new {Name = "chkSelect"});
            Actions.SetCheckBox((WinCheckBox) chkBoxControl, "True");

            ClickOnButton("Add Selected");
            //ClickOnButton("Save");

            var btnControl = Actions.GetWindowChild(windowInst, "btnAddSkillsExperience");
            Mouse.Click(btnControl);
        }
    }
}