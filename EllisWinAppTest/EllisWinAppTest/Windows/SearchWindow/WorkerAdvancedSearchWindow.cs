using System.Data;
using System.Linq;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class WorkerAdvancedSearchWindow : AppContext
    {


        #region Window Properties

        private static UITestControl GetWorkerSearchWindowProperties()
        {
            var workerSearchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            return workerSearchWindow;
        }

        private static UITestControl GetWorkerSearchResultsWindowProperties()
        {
            var workerSearchResultsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return workerSearchResultsWindow;
        }

        #endregion

        #region Worker Advanced Search Methods

        public static void SendTabs(int noOfTabs)
        {
            for (int tab = 0; tab < noOfTabs; tab++)
            {
                SendKeys.SendWait("{TAB}");
                Playback.Wait(1000);
            }
        }

        public static bool EnterWorkerAdvancedSearchData(DataRow data)
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                var firstname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.FirstName);
                Actions.SetText(firstname, data.ItemArray[3].ToString());

                var lastname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.LastName);
                Actions.SetText(lastname, data.ItemArray[4].ToString());

                var address = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Address);
                Actions.SetText(address, data.ItemArray[5].ToString());

                var city = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.City);
                Actions.SetText(city, data.ItemArray[6].ToString());

                var state = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.State);
                //state.SetFocus();
                //SendKeys.SendWait(data.ItemArray[7].ToString());
                DropDownActions.SelectDropdownByText(state, data.ItemArray[7].ToString());

                var zip = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.ZipCode);
                Actions.SetText(zip, data.ItemArray[8].ToString());

                var ssn = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.SSN);
                Actions.SetText(ssn, data.ItemArray[9].ToString());

                var workerNo = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.WorkerNumber);
                Actions.SetText(workerNo, data.ItemArray[10].ToString());

                return true;
            }

            return false;

        }

        public static bool EnterWorkerNameAsSearchData(string name)
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                var firstname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.FirstName);
                Actions.SetText(firstname, name);
                return true;
            }
            return false;

        }

        public static bool ClickCancelBtn()
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchWindow, WorkerSearchConstants.CancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickSearchBtn()
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchWindow, WorkerSearchConstants.SearchBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Worker Search Results Methods

        public static bool ClickOnRefineSearchBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.RefineSearchBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnPrintBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.PrintBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnExportBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.ExportBtn);
                return true;
            }
            return false;
        }

        public static bool CloseSearchResultsWindow()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)workerSearchResultsWindow);
                return true;
            }
            return false;
        }

        public static bool SelectWorkerfromSearchResults()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                TableActions.OpenRecordFromTable(workerSearchResultsWindow, WorkerSearchResultsConstatnts.SearchGrid,
                    "Worker Number", "000946008");
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class WorkerSearchConstants
        {
            public const string LastName = "txtLastName";
            public const string ZipCode = "txtWorkerZipCode";
            public const string City = "txtWorkerCity";
            public const string Address = "txtWorkerAddress";
            public const string FirstName = "txtFirstName";
            public const string WorkerNumber = "txtWorkerNumber";
            public const string SSN = "txtSSN";
            public const string State = "cmbWorkerState";
            public const string PositionFocus = "lstPositionFocus";
            public const string Title = "lstTitle";
            public const string Skills = "lstSkillsExperience";
            public const string SearchBtn = "_buttonSearch";
            public const string CancelBtn = "_cancelButton";
        }

        private class WorkerSearchResultsConstatnts
        {
            public const string SearchGrid = "_grdSearchResult";
            public const string RefineSearchBtn = "_btnRefineSearch";
            public const string PrintBtn = "_btnPrint";
            public const string ExportBtn = "_btnExport";

        }

        #endregion
    }
}