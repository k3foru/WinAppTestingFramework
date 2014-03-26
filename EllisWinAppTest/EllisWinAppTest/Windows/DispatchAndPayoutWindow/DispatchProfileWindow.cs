using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.DispatchAndPayoutWindow
{
    class DispatchProfileWindow : AppContext
    {
        public static UITestControl DispatchProfileWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return winInst;
        }

        public static UITestControl AssignWorerWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Assign Worker(s)" });
            return winInst;
        }

        public static UITestControl AssignAckWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Dispatch and Payout" });
            return winInst;
        }

        public static UITestControl ValidationMessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Message" });
            return winInst;
        }

        public static UITestControl ErrorMessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "ERROR" });
            return winInst;
        }

        public static bool HandleSelectWorkerValidationWindow()
        {
            var winInst = ValidationMessageWindowProperties();
            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_messageLabel");
                Console.WriteLine("Message for not selecting worker (atleast one) displayed as below...");
                Console.WriteLine(control.GetProperty("Value"));
                control = Actions.GetWindowChild(winInst, "_OKButton");
                Mouse.Click(control);
                return true;
            }
            return false;
        }

        public static bool HandleNoWorkerNoWeekSelectedValidationWindow()
        {
            var winInst = ErrorMessageWindowProperties();
            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_messageLabel");
                Console.WriteLine("Message for not selecting worker (atleast one) and not selecting day displayed as below...");
                Console.WriteLine(control.GetProperty("Value"));
                control = Actions.GetWindowChild(winInst, "_OKButton");
                Mouse.Click(control);
                return true;
            }
            return false;
        }
    }
}
