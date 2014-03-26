using System;
using Ellis.WinApp.Testing.Framework;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    internal class SearchWindow : AppContext
    {
        public static void SelectSearchElements(string text, string type, string searchType)
        {
            var toolbar = EllisWindow.Container.SearchFor<WinToolBar>(new {Name = "Toolbar"});

            if (searchType.Equals(SearchTypeConstants.Simple))
            {
                var search = toolbar.Items[6];
                search.WaitForControlEnabled();
                Actions.SetText(search, text);
            }

            var tupOne = SearchTypeTuple();
            var tupCategory = String.Empty;
            var tupType = String.Empty;

            foreach (var value in tupOne)
            {
                if (value.Item1.Equals(type))
                {
                    tupCategory = value.Item2;
                    tupType = value.Item3;

                    var categorydropDown = toolbar.Items[7];
                    MouseActions.Click(categorydropDown);
                    Actions.SendText(tupCategory);

                    var typedropDown = toolbar.Items[8];
                    MouseActions.Click(typedropDown);
                    Actions.SendText(tupType);
                }
            }
            var num = 0;
            if (searchType.Equals(SearchTypeConstants.Simple))
                num = 9;
            else if (searchType.Equals(SearchTypeConstants.Advanced))
                num = 10;

            var searchBtn = toolbar.Items[num];
            Mouse.Click(searchBtn);
        }

        public static WinWindow GetCustomerSearchResultsWindowProperties()
        {
            var window = App.Container.SearchFor<WinWindow>(new {Name = "Search Results"});
            GlobalWindows.CustomerSearchResultsWindow = window;
            return window;
        }

        public static void SelectFirstCustomerNameFromResults()
        {
            var searchResults = GetCustomerSearchResultsWindowProperties();
            var tableName = Actions.GetWindowChild(searchResults, CSearchControls.SearchResultGrid);
            var table = (WinTable) tableName;

            var row = table.Container.SearchFor<WinRow>(new {Name = "CustomerAdvancedSummaryDomain row 1"});
            var cell = row.Container.SearchFor<WinCell>(new {Name = "Customer Name"});
            Globals.CustomerName = cell.Value;
            MouseActions.DoubleClick(cell);
        }


        public static Tuple<string, string, string>[] SearchTypeTuple()
        {
            Tuple<string, string, string>[] SearchType =
            {
                Tuple.Create("Customer", "C", "C"),
                Tuple.Create("Worker", "W", "N"),
                //Tuple.Create("WorkerNo", "W","N"),
                //Tuple.Create("SSN", "W","N"),
                Tuple.Create("Lockout", "N", "L"),
                Tuple.Create("BillingLineItem", "A", "B"),
                Tuple.Create("Invoices", "A", ""),
                Tuple.Create("CreditCard", "A", "C"),
                Tuple.Create("ITransactions", "A", "I"),
                Tuple.Create("IRelationShips", "A", "II"),
                Tuple.Create("JobOrder", "Q", "J"),
                Tuple.Create("Quote", "Q", "Q"),
                Tuple.Create("Dispatch", "Q", "D"),
                Tuple.Create("WorkTicket", "Q", "W"),
                Tuple.Create("CheckRegister", "Q", "C")
            };

            return SearchType;
        }

        private class CSearchControls
        {
            public const string SearchResultGrid = "_grdSearchResult";
        }

        internal class SearchTypeConstants
        {
            public const string Simple = "Simple";
            public const string Advanced = "Advanced";
        }
    }
}