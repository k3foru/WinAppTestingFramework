//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class ListActions : AppContext
    {
        public static void SelectFromList(WinList control, string value)
        {
            var currentValue = GetSelectedValue(control);
            var allValues = GetListOfItems(control);

            foreach (var item in allValues.Where(item => item != currentValue))
            {
                control.SetFocus();
                control.SelectedItemsAsString = value;
                break;
            }
        }

        public static string GetSelectedValue(WinList control)
        {
            return control.SelectedItemsAsString;
        }

        public static List<string> GetListOfItems(WinList control)
        {
            var items =
                control.GetProperty("Items") as
                UITestControlCollection;

            return (from WinListItem item in items select item.DisplayText).ToList();
        }
    }
}
