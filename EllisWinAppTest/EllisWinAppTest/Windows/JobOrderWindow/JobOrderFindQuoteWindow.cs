using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;


namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class JobOrderFindQuoteWindow : AppContext
    {
        private static UITestControl GetCreateJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Create New JobOrder"});
            return joborderWindow;
        }

        private static UITestControlCollection GetJobOrderEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetCreateJobOrderWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection getJobOrderDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetCreateJobOrderWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        public static void EnterJobOrderFindQuoteData(DataRow data)
        {
            var getJobOrderWindow = GetCreateJobOrderWindowProperties();

            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            {
                var zipEdit = Actions.GetWindowChild(getJobOrderWindow, "txtJobSiteZip");
                zipEdit.SetFocus();
                            SendKeys.SendWait("{HOME}");
                Actions.SetText(zipEdit, data.ItemArray[9].ToString());

            }
            
            ClickOnButton("GO");
            //Enter data in dropdown fields
            //var getBasicJobInformationControlCollection = getJobOrderDropdownControlCollection();

                if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                    DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbState",
                    data.ItemArray[10].ToString());
                Playback.Wait(2000);
                if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
                    DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbCity",
                    data.ItemArray[11].ToString());
                Playback.Wait(2000);
                if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                    DropDownActions.SelectDropdownByText(getJobOrderWindow, "_cboPostalCodes",
                    data.ItemArray[9].ToString());
                Playback.Wait(2000);
                if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                    DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbCounty",
                    data.ItemArray[12].ToString());
                Playback.Wait(2000);

            ClickOnButton("GO");
            
        }

        public static void ClickOnButton(string btnName)
        {
            Factory.ClickOnButton(GetCreateJobOrderWindowProperties(), btnName);
        }


    }
}
