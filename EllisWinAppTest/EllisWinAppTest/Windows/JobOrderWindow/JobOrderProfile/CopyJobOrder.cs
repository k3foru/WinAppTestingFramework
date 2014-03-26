using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile
{
    internal class CopyJobOrder : AppContext
    {
        public static bool CopyJobOrderDetails()
        {
            var jobOrderWindow = OpenJobOrder.JobOrderProfileWindowProperties();

            if (jobOrderWindow.Exists)
            {
                Factory.ClickOnButton(jobOrderWindow, "Copy Job Order");
                var copyWindow = GetCopyOptionsWindowProperties();

                var chkBox = Actions.GetWindowChild(copyWindow, "_chkCopyJobOrderDetails");
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");

                Factory.ClickOnButton(copyWindow, "OK");
                return true;
            }
            return false;

        }

        public static bool CopyJobOrderAdditionalCharges()
        {
           var jobOrderWindow = OpenJobOrder.JobOrderProfileWindowProperties();

           if (jobOrderWindow.Exists)
            {
                Factory.ClickOnButton(jobOrderWindow, "Copy Job Order");
                var copyWindow = GetCopyOptionsWindowProperties();

                var chkBox = Actions.GetWindowChild(copyWindow, "_chkCopyAdditionalCharges");
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");

                Factory.ClickOnButton(copyWindow, "OK");
                return true;
            }
            return false;
        }

        public static UITestControl GetCopyOptionsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            return winGroup;
        }
    }
}