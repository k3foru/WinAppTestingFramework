using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.EllisWindow
{
    public class LandingPage : AppContext
    {
        public static bool VerifyJobOrderSchedulesDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var jo = workSpace.Container.SearchFor<WinButton>(new {Name = "Job Orders"});
            return jo.Enabled;
        }

        public static bool VerifyDispatchDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var dis = workSpace.Container.SearchFor<WinButton>(new {Name = "Dispatch"});
            return dis.Enabled;
        }

        public static bool VerifyWorkersDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var work = workSpace.Container.SearchFor<WinButton>(new {Name = "Workers"});
            return work.Enabled;
        }

        public static bool VerifyCustomersDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var cust = workSpace.Container.SearchFor<WinButton>(new {Name = "Customers"});
            return cust.Enabled;
        }

        public static bool VerifyArDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var ar = workSpace.Container.SearchFor<WinButton>(new {Name = "AR"});
            return ar.Enabled;
        }

        public static void SelectFromToolbar(string name)
        {
            var toolbar = EllisWindow.Container.SearchFor<WinToolBar>(new {Name = "ToolBar"});
            var joMenuItem = toolbar.Items[GetToolbarItem(name)];
            joMenuItem.WaitForControlEnabled();
            Mouse.Click(joMenuItem);
            Playback.Wait(2000);
        }

        public static void ClickOnCalendarButton(string type)
        {
            var btnWindow = Actions.GetWindowChild(EllisWindow, LandingPageControls.CalanderButton);
            var button = btnWindow.Container.SearchFor<WinButton>(new {Name = type});
            Mouse.Click(button);
        }

        public static void EnterDate(string datetype, string date)
        {
            var fromWindow = Actions.GetWindowChild(EllisWindow, datetype);
            var comboBox = (WinComboBox) fromWindow;
            MouseActions.Click(comboBox);
            for (int i = 0; i < 10; i++)
            {
                SendKeys.SendWait("{BACKSPACE}");
                Playback.Wait(200);
            }
            //SendKeys.SendWait("{HOME}");
            Playback.Wait(1500);
            SendKeys.SendWait(date);
        }

        public static void ClickDateTextbox(string datetype)
        {
            var fromWindow = Actions.GetWindowChild(EllisWindow, datetype);
            var comboBox = (WinComboBox) fromWindow;
            MouseActions.Click(comboBox);
        }

        public static void ClickOnCalendarClient()
        {
            var clientWindow = Actions.GetWindowChild(EllisWindow, LandingPageControls.CalendarClient);
            var client = (WinClient) clientWindow;
            Mouse.Click(client);
        }

        private static int GetToolbarItem(string name)
        {
            int num = 0;
            if (name == "JobOrder")
                num = 0;
            if (name == "Dispatch")
                num = 1;
            if (name == "Workers")
                num = 2;
            if (name == "Customers")
                num = 3;
            if (name == "AR")
                num = 4;

            return num;
        }

        public static void SelectCustomerInvoicesFromNavigationExplorer()
        {
            var child = Actions.GetWindowChild(EllisWindow, LandingPageControls.NavigationExplorer);
            var btn = child.Container.SearchFor<WinButton>(new {Name = "Customer Invoices"});
            Mouse.Click(btn);
        }

        public class LandingPageControls
        {
            public const string CalanderButton = "btnToggle";
            public const string AdvancedFromDate = "advancedFromDate";
            public const string AdvancedToDate = "advancedToDate";
            public const string CalendarClient = "CalendarView";

            public const string Advanced = "Advanced...";
            public const string Simple = "Simple...";

            public const string NavigationExplorer = "_NavigationExplorerBar";
        }
    }
}