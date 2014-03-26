using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    internal class NotificationAdvanceSearchWindow : AppContext
    {
        public static void EnterNotificationAdvancedSearchData(DataRow data)
        {
            var editControlcollection = GetEditControlCollection();
            var dropDownCollection = GetDropDownControlCollection();
            var searchWindow = GetNotificationSearchWindowProperties();

            foreach (var control in editControlcollection.Where(control => control.FriendlyName != null))
            {
                if (control.FriendlyName.Equals(NotificationSearchConstants.RequestNumber) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[3].ToString()))
                    Actions.SetText(control, data.ItemArray[3].ToString());

                if (control.FriendlyName.Equals(NotificationSearchConstants.Requester) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[4].ToString()))
                    Actions.SetText(control, data.ItemArray[4].ToString());

                if (control.FriendlyName.Equals(NotificationSearchConstants.Customer) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[5].ToString()))
                    Actions.SetText(control, data.ItemArray[5].ToString());

                if (control.FriendlyName.Equals(NotificationSearchConstants.DecsionMaker) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[6].ToString()))
                    Actions.SetText(control, data.ItemArray[6].ToString());

                //if (control.FriendlyName.Equals(NotificationSearchConstants.Branch) &&
                //   !control.GetProperty("Value").Equals(data.ItemArray[8].ToString()))

                //    Actions.SetText(control, data.ItemArray[8].ToString());


                //Actions.SetText(control, data.ItemArray[8].ToString());

                //if (control.FriendlyName.Equals(NotificationSearchConstants.SSN) &&
                //   !control.GetProperty("Value").Equals(data.ItemArray[9].ToString()))
                //    Actions.SetText(control, data.ItemArray[9].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, NotificationSearchConstants.Branch,
                     data.ItemArray[7].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, NotificationSearchConstants.LockoutStatus,
                     data.ItemArray[7].ToString());

            //if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, NotificationSearchConstants.Branch,
            //        data.ItemArray[7].ToString());

            //if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, NotificationSearchConstants.LockoutStatus,
            //        data.ItemArray[7].ToString());
        }

        private static UITestControl GetNotificationSearchWindowProperties()
        {
            var notificationSearchWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Search"});
            return notificationSearchWindow;
        }

        private static UITestControlCollection GetEditControlCollection()
        {
            Playback.Wait(3000);
            var group = GetControlsGroup();
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetDropDownControlCollection()
        {
            Playback.Wait(3000);
            var group = GetControlsGroup();
            var dDownControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var dDownControlcollection = Actions.GetControlCollection(dDownControl);
            return dDownControlcollection;
        }

        private static WinGroup GetControlsGroup()
        {
            var searchWindow = GetNotificationSearchWindowProperties();
            var group = searchWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var secondGroup = group.Container.SearchFor<WinGroup>(new {Name = ""});

            return secondGroup;
        }

        public static void ClickCancelBtn()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Search"});
            var cancelBtn = searchWindow.Container.SearchFor<WinButton>(new {Name = "Cancel"});
            MouseActions.Click(cancelBtn);
        }

        public static void ClickSearchBtn()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Search"});
            var searchBtn = searchWindow.Container.SearchFor<WinButton>(new {Name = "Search"});
            MouseActions.Click(searchBtn);
        }

        public class NotificationSearchConstants
        {
            public const string RequestNumber = "txtLockoutRequestNumber";
            public const string DecsionMaker = "txtDMFirstName";
            public const string Customer = "txtCFirstName";
            public const string Requester = "txtRFirstName";
            public const string LockoutStatus = "cbLockoutType";
            public const string Branch = "cbBranch";
            //public const string SSN = "txtSSN";
            //public const string State = "cmbWorkerState";
            //public const string CustomerFEDID = "_txtCustomerFEDID";
            //public const string InvoiceNumber = "_txtEllisInvoiceNo";
            //public const string BillingAddressZip = "_txtBillingAddressZip";
            //public const string CustomerAddress = "_txtCustomerAddress";
            //public const string BillingAddressCity = "_txtBillingAddressCity";

            //public const string Section = "_cmbSection";
            //public const string SubSection = "_cmbSubSection";
            //public const string BillingState = "_cboBillingState";
            //public const string AccountManagedBy = "_cboAccountManagedBy";
            //public const string Branch = "_cmbBranch";
            //public const string FromDate = "_calendarFrom";
            //public const string ToDate = "_calendarTo";
        }
    }
}