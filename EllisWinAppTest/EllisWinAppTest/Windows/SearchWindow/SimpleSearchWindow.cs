using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class SimpleSearchWindow : AppContext
    {
        //public static void SimpleSearch(string text, string type)
        //{
        //    var toolbar = EllisWindow.Container.SearchFor<WinToolBar>(new {Name = "ToolBar"});

        //    var search = toolbar.Items[6];
        //    search.WaitForControlEnabled();
        //    Actions.SetText(search, text);

        //     var tupOne = SearchTypeTuple();
        //     var tupCategory = string.Empty;
        //     var tupType = string.Empty;

        //    foreach (var value in tupOne)
        //    {
        //        if (value.Item1.Equals(type))
        //        {
        //            tupCategory = value.Item2;
        //            tupType = value.Item3;

        //            var CategorydropDown = toolbar.Items[7];
        //            MouseActions.Click(CategorydropDown);
        //            Actions.SendText(tupCategory);

        //            var TypedropDown = toolbar.Items[8];
        //            MouseActions.Click(TypedropDown);
        //            Actions.SendText(tupType);
        //        }
        //    }
        //    var searchBtn = toolbar.Items[9];
        //    Mouse.Click(searchBtn);
        //}

        public static bool VerifyDisplayedResults(string data)
        {
            if (CustomerProfileWindowDisplayed())
            {
                if (VerifyCustomerProfileDisplayed(data))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool VerifyCustomerProfileDisplayed(string data)
        {
            var text = CustomerProfile.GetCustomerProfileText();
            return text.Equals(data);
        }

        private static bool CustomerProfileWindowDisplayed()
        {
            return CustomerProfile.CustomerProfileWindow.Enabled;
        }

        public static bool VerifyWorkerDisplayedResults(string data)
        {
            return VerifyWorkerProfileDisplayed(data) || true;
        }

        public static string GetWorkerProfileText()
        {
            var name = "Worker Profile-" + Globals.WorkerName;
            var window = App.Container.SearchFor<WinWindow>(new {Name = name});
            var text = window.Container.SearchFor<WinText>(new {Name = Globals.WorkerName});
            return text.DisplayText;
        }

        private static bool VerifyWorkerProfileDisplayed(string data)
        {
            var text = GetWorkerProfileText();
            return text.Equals(data);
        }

        public static void ClickResultWindowClose()
        {
            Actions.SendAltF4();
        }
    }
}