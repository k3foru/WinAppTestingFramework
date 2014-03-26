using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerGarnishmentsWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        #endregion

        #region TransactionHistory Tab Methods

        public static void SelectTransactionHistoryTab()
        {
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{RIGHT}");
        }

        public static bool ClickonGoBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var goBtn = Actions.GetWindowChild(workerProfileWindow, WorkerTransactionHistoryTabConstants.GoBtn);
                Mouse.Click(goBtn);
                return true;
            }
            return false;
        }

        public static bool EnterDataInTransactionHistoryTab(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            if (workerProfileWindow.Exists)
            {
                var tType = Actions.GetWindowChild(workerProfileWindow,
                    WorkerTransactionHistoryTabConstants.TransactionType);
                DropDownActions.SelectDropdownByText(tType, data.ItemArray[4].ToString());
                //tType.SetFocus();
                //SendKeys.SendWait(data.ItemArray[15].ToString());

                var dateFrom = Actions.GetWindowChild(workerProfileWindow, WorkerTransactionHistoryTabConstants.DateFrom);
                dateFrom.SetFocus();
                //DropDownActions.SelectDropdownByText(dateFrom, data.ItemArray[16].ToString());
                Actions.SendText(" ");
                Actions.SendText("{HOME}");
                SendKeys.SendWait(data.ItemArray[5].ToString());

                var dateTo = Actions.GetWindowChild(workerProfileWindow, WorkerTransactionHistoryTabConstants.DateTo);
                dateTo.SetFocus();
                //DropDownActions.SelectDropdownByText(dateTo, data.ItemArray[17].ToString());
                Actions.SendText(" ");
                Actions.SendText("{HOME}");
                SendKeys.SendWait(data.ItemArray[6].ToString());

                var oNumber = Actions.GetWindowChild(workerProfileWindow,
                    WorkerTransactionHistoryTabConstants.OrderNumber);
                DropDownActions.SelectDropdownByText(oNumber, data.ItemArray[7].ToString());
                //oNumber.SetFocus();
                //SendKeys.SendWait(data.ItemArray[18].ToString());
                return true;
            }
            return false;
        }

        #endregion

        #region Existing Orders Tab Methods

        public static bool ClickOnPrintBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var printbtn = Actions.GetWindowChild(workerProfileWindow, WorkerExistingOrderTabConstants.PrintBtn);
            if (printbtn.Enabled)
            {
                Mouse.Click(printbtn);
                return true;
            }

            return false;

        }

        public static void SelectDataInComboBox(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var cmbBox = Actions.GetWindowChild(workerProfileWindow, WorkerExistingOrderTabConstants.ExistingsOrders);
            DropDownActions.SelectDropdownByText(cmbBox, data.ItemArray[3].ToString());

        }

        #endregion

        #region Controls

        private class WorkerExistingOrderTabConstants
        {
            public const string ExistingsOrders = "cmbExistingOrders";
            public const string ExistingOrdersgrd = "grdExistingOrders";
            public const string PrintBtn = "btnPrint1";
        }

        private class WorkerTransactionHistoryTabConstants
        {
            public const string TransactionType = "cmbTransactionType";
            public const string OrderNumber = "cmbOrderNumber";
            public const string DateFrom = "ultradteDateFrom";
            public const string DateTo = "ultradteDateTo";
            public const string TransactionHistory = "grdTransactionHistory";
            public const string GoBtn = "btnGo";
        }

        #endregion
    }
}