using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    internal class ARAdvancedSearchWindow : AppContext
    {
        public static void EnterBillingSearchData(DataRow data)
        {
            var editControlcollection = GetEditControlCollection();
            var dropDownCollection = GetDropDownControlCollection();

            foreach (var control in editControlcollection.Where(control => control.FriendlyName != null))
            {
                if (control.FriendlyName.Equals(BillingSearchConstants.CustomerName) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[3].ToString()))
                    Actions.SetText(control, data.ItemArray[3].ToString());

                if (control.FriendlyName.Equals(BillingSearchConstants.CustomerNumber) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[4].ToString()))
                    Actions.SetText(control, data.ItemArray[4].ToString());

                //if (control.FriendlyName.Equals(BillingSearchConstants.BillingPhone) &&
                //   !control.GetProperty("Value").Equals(data.ItemArray[7].ToString()))
                //    Actions.SetText(control, data.ItemArray[7].ToString());               
            }

            //if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, BillingSearchConstants.DispatchStart, data.ItemArray[6].ToString());

            //if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, BillingSearchConstants.DispatchEnd, data.ItemArray[7].ToString());

            //if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, BillingSearchConstants.District, data.ItemArray[8].ToString());

            //if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            //   DropDownActions.SelectFromDropDown(dropDownCollection, BillingSearchConstants.Branch, data.ItemArray[9].ToString());
        }

        public static void EnterCreditCardSearchData(DataRow data)
        {
            var editControlcollection = GetEditControlCollection();
            var dropDownCollection = GetDropDownControlCollection();
            var searchWindow = GetSearchWindowProperties();

            foreach (var control in editControlcollection.Where(control => control.FriendlyName != null))
            {
                if (control.FriendlyName.Equals(CreditCardSearchConstants.CreditCard) &&
                    !control.GetProperty("Value").Equals(data.ItemArray[10].ToString()))
                    Actions.SetText(control, data.ItemArray[10].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CreditCardSearchConstants.Transaction,
                    data.ItemArray[11].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CreditCardSearchConstants.ProcessedFrom,
                    data.ItemArray[12].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CreditCardSearchConstants.ProcessedTo,
                    data.ItemArray[13].ToString());
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
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static WinGroup GetControlsGroup()
        {
            var searchWindow = GetSearchWindowProperties();
            var group = searchWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var secondGroup = group.Container.SearchFor<WinGroup>(new {Name = ""});

            return secondGroup;
        }

        private static UITestControl GetSearchWindowProperties()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Search"});
            return searchWindow;
        }

        public static void ClickCancelBtn()
        {
            var searchWindow = GetSearchWindowProperties();
            var cancelBtn = searchWindow.Container.SearchFor<WinButton>(new {Name = "Cancel"});
            MouseActions.Click(cancelBtn);
        }

        public static void ClickSearchBtn()
        {
            var searchWindow = GetSearchWindowProperties();
            var searchBtn = searchWindow.Container.SearchFor<WinButton>(new {Name = "Search"});
            MouseActions.Click(searchBtn);
        }

        public static bool VerifyDefaultDistrictSelected(string data)
        {
            var window = GetSearchWindowProperties();

            var cmb = Actions.GetWindowChild(window, "_cmbDistrict");
            var cmbBox = (WinComboBox) cmb;
            return cmbBox.SelectedItem.Equals(data);
        }

        public static bool VerifyInvoicingOrganizationIsNull()
        {
            var window = GetSearchWindowProperties();

            var cmb = Actions.GetWindowChild(window, "cmbOrganization");
            var cmbBox = (WinComboBox) cmb;
            return string.IsNullOrEmpty(cmbBox.SelectedItem);
        }

        public class BillingSearchConstants
        {
            public const string CustomerName = "_txtCustomerName";
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string DispatchStart = "_dtpStartDate";
            public const string DispatchEnd = "_dtpEndDate";
            public const string District = "_ddlDistrict";
            public const string Branch = "_ddlBranch";
        }

        public class CreditCardSearchConstants
        {
            public const string CreditCard = "_txtCreditCardNo";
            public const string Transaction = "_cbTransactionSatus";
            public const string ProcessedFrom = "_calendarFrom";
            public const string ProcessedTo = "_calendarTo";
        }
    }
}