using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.AccountsReceivableWindow
{
    public class ARWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetCustomerCollectionsWindowProperties()
        {
            //var ccollectionsWindow =
            //   App.Container.SearchFor<WinWindow>(new { Name = "Customer Collections - Call Back" });
            //return ccollectionsWindow;
            return Actions.GetWindowProperties(App, "Customer Collections - Call Back");
        }

        private static UITestControl GetAddNoteWindowProperties()
        {
            //var addNoteWindow =
            //   App.Container.SearchFor<WinWindow>(new { Name = "Add Note" });
            //return addNoteWindow;
            return Actions.GetWindowProperties(App, "Add Note");
        }

        public static UITestControl GetPaymentLockboxWindowProperties()
        {
            //var paymentLockboxWindow =
            //   App.Container.SearchFor<WinWindow>(new { Name = "Payment Lockbox Batch" });
            //return paymentLockboxWindow;
            return Actions.GetWindowProperties(App, "Payment Lockbox Batch");
        }

        public static UITestControl GetPaymentProfileWindowProperties()
        {
            //var paymentLockboxWindow =
            //   App.Container.SearchFor<WinWindow>(new { Name = "Payment Profile - Item #1" });
            //return paymentLockboxWindow;
            return Actions.GetWindowProperties(App, "Payment Profile - Item #1");
        }

        private static UITestControl GetPaymentProfileInvoiceWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Payment Profile - Invoice Search");
        }

        private static UITestControl GetAlertWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Alert - Inconsistent Assignment");
        }

        #endregion

        #region Add Notes Methods

        public static bool AddNoteWindowDisplayed()
        {
            var prop = GetAddNoteWindowProperties();
            return prop.Exists;
        }

        private static void SelectAnInvoiceNumberFromGrid(UITestControl prop)
        {
            var row = TableActions.SelectRowFromTable(prop, ARControls.UnpaidInvoiceGrid,
                "CollectionInvoiceSummaryDomain row 1");
            //var tableName = Actions.GetWindowChild(prop, ARControls.UnpaidInvoiceGrid);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "CollectionInvoiceSummaryDomain row 1" });
            var cell = row.Container.SearchFor<WinCell>(new {Value = "False"});
            Mouse.Click(cell);
        }

        public static void AddNewNoteToUnpaidInvoice()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            SelectAnInvoiceNumberFromGrid(prop);

            //var notes = Actions.GetWindowChild(prop, ARControls.Notes);
            //Actions.SetText(notes, "This is an unpaid invoice. This note is just for testing purpose.");

            //var btn = Actions.GetWindowChild(prop, ARControls.SaveAndClose);
            //MouseActions.Click(btn);

            Actions.SetText(prop, ARControls.Notes, "This is an unpaid invoice. This note is just for testing purpose.");
            MouseActions.ClickButton(prop, ARControls.SaveAndClose);
        }

        public static void AddNewNoteToCancelCallback()
        {
            var prop = GetAddNoteWindowProperties();
            SelectAnInvoiceNumberFromGrid(prop);

            //var notes = Actions.GetWindowChild(prop, ARControls.Notes);
            //Actions.SetText(notes, "This is a note to cancel call back. This note is just for testing purpose.");

            //var btn = Actions.GetWindowChild(prop, ARControls.SaveNotes);
            //MouseActions.Click(btn);

            Actions.SetText(prop, ARControls.Notes,
                "This is a note to cancel call back. This note is just for testing purpose.");
            MouseActions.ClickButton(prop, ARControls.SaveNotes);
        }

        public static bool VerifyNewNoteAdded()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            var cell = TableActions.SelectCellFromTable(prop, ARControls.OverdueCustomers,
                "OverdueCustomerSummaryDomain row 2",
                "Customer Name");
            //var tableName = Actions.GetWindowChild(prop, ARControls.OverdueCustomers);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "OverdueCustomerSummaryDomain row 2" });
            //var cell = row.Container.SearchFor<WinCell>(new { Name = "Customer Name" });
            Globals.CustomerName = cell.Value;
            Mouse.DoubleClick(cell);

            var returnAns =
                CustomerProfileWindow.VerifyNewNoteDisplayedInGrid(
                    "This is an unpaid invoice. This note is just for testing purpose.");
            return returnAns;
        }

        public static void ClickAddNoteButton()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            //var btn = Actions.GetWindowChild(prop, ARControls.AddNoteButton);
            //Mouse.Click(btn);
            MouseActions.ClickButton(prop, ARControls.AddNoteButton);
        }

        #endregion

        #region Verification Methods

        public static bool VerifyInvoices(string data)
        {
            var upInvoice = Actions.GetWindowChild(EllisWindow, ARControls.UnpaidInvoices);
            var dd = (WinComboBox) upInvoice;
            return dd.SelectedItem.Equals(data);
        }

        public static bool VerifyMyOrg(string data)
        {
            var myOrg = Actions.GetWindowChild(EllisWindow, ARControls.CollectingOrg);
            var dd = (WinComboBox) myOrg;

            return dd.SelectedItem.Equals(data);
        }

        public static bool VerifyOverDueDisplayed()
        {
            var row = TableActions.SelectRowFromTable(EllisWindow, ARControls.OverdueCustomers,
                "OverdueCustomerSummaryDomain row 1");
            //var tableName = Actions.GetWindowChild(EllisWindow, ARControls.OverdueCustomers);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "OverdueCustomerSummaryDomain row 1" });
            var cell = row.Container.SearchFor<WinCell>(new {Instance = "6"});

            return cell.Enabled;
        }

        public static void SelectCustomerCollectionFromLandingPage()
        {
            var cell = TableActions.SelectCellFromTable(EllisWindow, ARControls.OverdueCustomers,
                "OverdueCustomerSummaryDomain row 1",
                "Customer Name");
            Globals.CustomerName = cell.Value;
            Mouse.DoubleClick(cell);

            //var tableName = Actions.GetWindowChild(EllisWindow, ARControls.OverdueCustomers);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "OverdueCustomerSummaryDomain row 1" });
            //var cell = row.Container.SearchFor<WinCell>(new { Name = "Customer Name" });
            //Globals.CustomerName = cell.Value;
            //Mouse.Click(cell);
        }

        public static bool VerifyCustomerProfileWindowDisplayedWhenCustomerNumberClicked()
        {
            var custom = Actions.GetWindowChild(EllisWindow, ARControls.CustNumber);
            Mouse.Click(custom);

            var prop = CustomerProfileWindow.GetCustomerProfileWindowProperties();
            return prop.Exists;
        }

        public static bool VerifyCustomerCollectionsWindowDisplayed()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            return prop.Exists;
        }

        public static bool VerifyProfileWindowDisplayedWhenCusNameLinkClicked()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            //var custName = Actions.GetWindowChild(prop, ARControls.CustName);
            //Mouse.Click(custName);
            MouseActions.ClickButton(prop, ARControls.CustName);

            return CustomerProfileWindow.GetCustomerProfileWindowProperties().Enabled;
        }

        public static bool VerifyProfileWindowDisplayedWhenCusNumberLinkClicked()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            //var custName = Actions.GetWindowChild(prop, ARControls.CustNumber);
            //Mouse.Click(custName);
            MouseActions.ClickButton(prop, ARControls.CustNumber);

            return CustomerProfileWindow.GetCustomerProfileWindowProperties().Enabled;
        }

        public static bool VerifyBalanceDueAmountDisplayed()
        {
            var prop = GetCustomerCollectionsWindowProperties();
            UITestControl due = null;
            try
            {
                due = Actions.GetWindowChild(prop, ARControls.BalanceDue);
            }
            catch (Exception)
            {
                //Suppress any expcetion here
            }

            return due != null;
        }

        #endregion

        #region Lockbox Methods

        public static void SelectFirstCustomerInvoiceFromTable()
        {
            var cell = TableActions.SelectCellFromTable(EllisWindow, ARControls.LockboxTable,
                "LockboxBatchSummaryDomain row 1", "Batch #");
            Globals.CustomerName = cell.Value;
            MouseActions.Click(cell);

            //var tableName = Actions.GetWindowChild(EllisWindow, ARControls.LockboxTable);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "LockboxBatchSummaryDomain row 1" });
            //var cell = row.Container.SearchFor<WinCell>(new { Name = "Batch #" });
            //MouseActions.Click(cell);
        }

        public static void SelectFirstRemittenceFromTable()
        {
            var prop = GetPaymentLockboxWindowProperties();
            var cell = TableActions.SelectCellFromTable(prop, ARControls.PaymentDetailsTable, "RemittanceTransfer row 1",
                "#");
            MouseActions.Click(cell);

            //var tableName = Actions.GetWindowChild(prop, ARControls.PaymentDetailsTable);
            //var table = (WinTable)tableName;
            //var row = table.Container.SearchFor<WinRow>(new { Name = "RemittanceTransfer row 1" });
            //var cell = row.Container.SearchFor<WinCell>(new { Name = "#" });
            //MouseActions.Click(cell);
        }

        public static void ClickOpeninNewWindowButton()
        {
            var prop = GetPaymentProfileWindowProperties();
            MouseActions.ClickButton(prop, ARControls.OpenInNewWindow);
        }

        public static bool VerifyPaymentInvoiceWindowDisplayed()
        {
            var prop = GetPaymentProfileInvoiceWindowProperties();
            return prop.Enabled;
        }

        public static bool VerifyRemainingAmountDisplayedOnWindow()
        {
            var prop = GetPaymentProfileInvoiceWindowProperties();
            var label = Actions.GetWindowChild(prop, ARControls.RemainingAmount);
            return label.Exists;
        }

        public static void ClosePaymentInvoiceWindow()
        {
            var prop = GetPaymentProfileInvoiceWindowProperties();
            MouseActions.ClickButton(prop, ARControls.CancelButton);
        }

        public static void ClosePaymentProfileWindow()
        {
            var prop = GetPaymentProfileWindowProperties();
            MouseActions.ClickButton(prop, ARControls.CancelButton);
        }

        public static void ClickCompleteButton()
        {
            var prop = GetPaymentLockboxWindowProperties();
            MouseActions.ClickButton(prop, ARControls.CompleteButton);
        }

        public static void HandleAlertWindow()
        {
            try
            {
                var prop = GetAlertWindowProperties();
                MouseActions.ClickButton(prop, ARControls.OkButton);
            }
            catch (Exception)
            {
                //suppress any exception here
            }
        }

        public static bool VerifyCustomerInvGridDisplayed()
        {
            var tableName = Actions.GetWindowChild(EllisWindow, ARControls.InvoiceGrid);
            var table = (WinTable) tableName;

            return table.Exists;
        }

        #endregion

        private class ARControls
        {
            public const string UnpaidInvoices = "_cmbunpaidInvoices";
            public const string CollectingOrg = "_cmbCollectingOrgFilter";
            public const string GoBtn = "_btnGo";
            public const string ClearBtn = "_btnClear";
            public const string OverdueCustomers = "_gridOverdueCustomers";
            public const string AddNoteButton = "_btnAddNote";
            public const string OkButton = "btnOk";

            public const string CustNumber = "_lblNumber";
            public const string CustName = "_lblName";

            #region Customer Invoices Controls

            public const string InvoiceGrid = "_grdInvoices";

            #endregion

            #region Customer Collections Controls

            public const string CustomerContacts = "_cbCustomerContacts";
            public const string UnpaidInvoiceGrid = "_grdUnpaidInvoices";
            public const string Notes = "_txtNotes";
            public const string SaveAndClose = "_btnRecord";
            public const string Cancel = "_btnCancel";

            public const string SaveNotes = "_btnSave";

            public const string BalanceDue = "_lblBalanceDueAmount";

            #endregion

            #region ARM user Lockbox controls

            public const string LockboxTable = "lockboxBatches";
            public const string PaymentDetailsTable = "_paymentDetailsGrid";
            public const string OpenInNewWindow = "openInNewWindowButton";
            public const string RemainingAmount = "remainingAmountLabel";
            public const string CancelButton = "btnCancel";
            public const string CompleteButton = "btnComplete";

            #endregion
        }
    }
}