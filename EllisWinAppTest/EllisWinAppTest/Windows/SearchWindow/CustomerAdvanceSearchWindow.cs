using System;
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
    internal class CustomerAdvanceSearchWindow : AppContext
    {
        public static void SendTabs(int noOfTabs)
        {
            for (var tab = 0; tab < noOfTabs; tab++)
            {
                Actions.SendTab();
                Playback.Wait(1000);
            }
        }

        public static void EnterAdvancedSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            Actions.SetText(searchWindow, CustomerSearchConstants.CustomerName, data.ItemArray[3].ToString());
            Actions.SetText(searchWindow, CustomerSearchConstants.CustomerNumber, data.ItemArray[4].ToString());

            //SelectRadioBtn(data.ItemArray[6].ToString());
            //Actions.SetText(searchWindow, CustomerSearchConstants.CustomerFEDID, data.ItemArray[5].ToString());

            Actions.SetText(searchWindow, CustomerSearchConstants.BillingPhone, data.ItemArray[7].ToString());
            Actions.SetText(searchWindow, CustomerSearchConstants.CustomerAddress, data.ItemArray[8].ToString());
            Actions.SetText(searchWindow, CustomerSearchConstants.BillingAddressCity, data.ItemArray[9].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CustomerSearchConstants.BillingState, data.ItemArray[10].ToString());

            Actions.SetText(searchWindow, CustomerSearchConstants.BillingAddressZip, data.ItemArray[11].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CustomerSearchConstants.AccountManagedBy, data.ItemArray[12].ToString()); 

            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CustomerSearchConstants.Branch,
                     data.ItemArray[13].ToString());

            Actions.SetText(searchWindow, CustomerSearchConstants.ParentCustomerNumber, data.ItemArray[14].ToString());
            Actions.SetText(searchWindow, CustomerSearchConstants.InvoiceNumber, data.ItemArray[15].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CustomerSearchConstants.FromDate,
                     data.ItemArray[17].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
                DropDownActions.SelectDropdownByText(searchWindow, CustomerSearchConstants.ToDate,
                     data.ItemArray[18].ToString());

            Actions.SetText(searchWindow, CustomerSearchConstants.TicketNumber, data.ItemArray[20].ToString());
            Actions.SetText(searchWindow, CustomerSearchConstants.JobOrder, data.ItemArray[21].ToString());
        }

        //public static void EnterAdvancedSearchData(DataRow data)
        //{
        //    var editControlcollection = GetEditControlCollection();
        //    var dropDownCollection = GetDropDownControlCollection();

        //    foreach (var control in editControlcollection.Where(control => control.FriendlyName != null))
        //    {
        //        if (control.FriendlyName.Equals(CustomerSearchConstants.CustomerName) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[3].ToString()))
        //            Actions.SetText(control, data.ItemArray[3].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.CustomerNumber) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[4].ToString()))
        //            Actions.SetText(control, data.ItemArray[4].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.BillingPhone) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[7].ToString()))
        //            Actions.SetText(control, data.ItemArray[7].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.CustomerAddress) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[8].ToString()))
        //            Actions.SetText(control, data.ItemArray[8].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.BillingAddressCity) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[8].ToString()))
        //            Actions.SetText(control, data.ItemArray[8].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.ParentCustomerNumber) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[15].ToString()))
        //            Actions.SetText(control, data.ItemArray[15].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.InvoiceNumber) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[16].ToString()))
        //            Actions.SetText(control, data.ItemArray[16].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.TicketNumber) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[20].ToString()))
        //            Actions.SetText(control, data.ItemArray[20].ToString());

        //        if (control.FriendlyName.Equals(CustomerSearchConstants.JobOrder) &&
        //            !control.GetProperty("Value").Equals(data.ItemArray[21].ToString()))
        //            Actions.SetText(control, data.ItemArray[21].ToString());
        //    }

        //    //foreach (
        //    //    WinComboBox control in
        //    //        dropDownCollection.Where(
        //    //            control => control.FriendlyName != null || control.FriendlyName != "")
        //    //            .Where(
        //    //                control =>
        //    //                    control.FriendlyName != null &&
        //    //                    control.FriendlyName.Equals(CustomerSearchConstants.AccountManagedBy)))
        //    //   DropDownActions.SelectFromDropDown(control, data.ItemArray[13].ToString());

        //    if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
        //        DropDownActions.SelectFromDropDown(dropDownCollection, CustomerSearchConstants.AccountManagedBy,
        //             data.ItemArray[12].ToString());

        //    if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
        //        DropDownActions.SelectFromDropDown(dropDownCollection, CustomerSearchConstants.Branch,
        //             data.ItemArray[13].ToString());

        //    if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
        //        DropDownActions.SelectFromDropDown(dropDownCollection, CustomerSearchConstants.FromDate,
        //             data.ItemArray[17].ToString());

        //    if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
        //        DropDownActions.SelectFromDropDown(dropDownCollection, CustomerSearchConstants.ToDate,
        //             data.ItemArray[18].ToString());
        //}

        public static void EnterCustomerNameAsSearchData(string name)
        {
            var control = GetSearchWindowProperties();
            Actions.SetText(control, CustomerSearchConstants.CustomerName, name);
            ClickOnSearchButton();
        }

        public static void ClickOnSearchButton()
        {
            var searchWindow = GetSearchWindowProperties();
            var searchBtn = searchWindow.Container.SearchFor<WinButton>(new { Name = "Search" });
            MouseActions.Click(searchBtn);

            Playback.Wait(5000);
        }

        public static bool VerifySearchResultsWindowDisplayed()
        {
            try
            {
                var searchResultsWindow = GetSearchResultsWindowProperties();
                if (searchResultsWindow.Enabled)
                    return true;
            }
            catch (Exception)
            {
                //suppress any exception here
            }
            return false;
        }

        public static bool VerifySearchValidationWindowDisplayed()
        {
            try
            {
                var searchValidationWindow = GetSearchValidationErrorWindowProperties();
                if (searchValidationWindow.Enabled)
                    return true;
            }
            catch (Exception)
            {
                //suppress any exception here
            }
            return false;
        }

        //private static void EnterDataInTextBox(UITestControl control, string data)
        //{
        //    if (control.FriendlyName.Equals(CustomerSearchConstants.CustomerName) &&
        //        !control.GetProperty("Value").Equals(data))
        //        Actions.SetText(control, data);
        //}

        //private static void SetCustomerFedId(string data)
        //{
        //    var control = GetControlsGroup();
        //    var fedText = control.Container.SearchFor<WinEdit>(new { Name = "Text area" });
        //    Actions.SetText(fedText, data);
        //}

        private static void SelectRadioBtn(string data)
        {
            var group = GetControlsGroup();

            var editControl = group.Container.SearchFor<WinWindow>(new { Name = "FED-ID" });

            var control = GetControlsGroup();

            if (data == "FED-ID")
            {
                var radio = control.Container.SearchFor<WinRadioButton>(new { Name = "FED-ID" });
                MouseActions.Click(radio);
            }
            else if (data == "SSN")
            {
                var radio = control.Container.SearchFor<WinRadioButton>(new { Name = "SSN" });
                Mouse.Click(radio);
            }
            else if (data == "ITIN")
            {
                var radio = control.Container.SearchFor<WinRadioButton>(new { Name = "ITIN" });
                MouseActions.Click(radio);
            }
        }

        //private static UITestControlCollection GetEditControlCollection()
        //{
        //    Playback.Wait(3000);
        //    var group = GetControlsGroup();
        //    var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
        //    var editControlcollection = Actions.GetControlCollection(editControl);
        //    return editControlcollection;
        //}

        //private static UITestControlCollection GetDropDownControlCollection()
        //{
        //    Playback.Wait(3000);
        //    var group = GetControlsGroup();
        //    var editControl = group.Container.SearchFor<WinComboBox>(new { Name = "" });
        //    var editControlcollection = Actions.GetControlCollection(editControl);
        //    return editControlcollection;
        //}

        private static WinGroup GetControlsGroup()
        {
            var searchWindow = GetSearchWindowProperties();
            var group = searchWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var secondGroup = group.Container.SearchFor<WinGroup>(new { Name = "" });

            return secondGroup;
        }

        private static UITestControl GetSearchWindowProperties()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            return searchWindow;
        }

        private static UITestControl GetSearchResultsWindowProperties()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return searchWindow;
        }

        private static UITestControl GetSearchValidationErrorWindowProperties()
        {
            var searchValidationWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Validation Error" });
            return searchValidationWindow;
        }

        public static void CloseAdvanceSearchWindow()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            var closeBtn = searchWindow.Container.SearchFor<WinButton>(new { Name = "Close" });
            MouseActions.Click(closeBtn);
        }

        public static void CloseSearchResultsWindow()
        {
            //Playback.Wait(10000);
            ////var searchWindow =
            ////    App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            ////searchWindow.SetFocus();
            ////var closeBtn = searchWindow.Container.SearchFor<WinButton>(new { Name = "Close" });
            ////MouseActions.Click(closeBtn);
            ////Helpers.Factory.SendAltF4();
            //SendKeys.SendWait("{RIGHT}");
            //Playback.Wait(1000);
            //SendKeys.SendWait("{ENTER}");
            TitlebarActions.ClickClose((WinWindow) GetSearchResultsWindowProperties());
        }

        public static void ClickValidationOk()
        {
            //var searchValidationWindow = GetSearchValidationErrorWindowProperties();
            //var okWindow = searchValidationWindow.Container.SearchFor<WinWindow>(new { Name = "OK" });
            //var btn = okWindow.Container.SearchFor<WinButton>(new { Name = "OK" });
            //MouseActions.Click(btn);

            TitlebarActions.ClickClose((WinWindow) GetSearchValidationErrorWindowProperties());
            SendKeys.SendWait("%O");
        }

        public class CustomerSearchConstants
        {
            public const string JobOrder = "_txtJobOrder";
            public const string UltraTextEditor = "UltraTextEditor1";
            public const string CreditTerms = "_cboCreditTerms";
            public const string BillingPhone = "_txtBillingPhone";
            public const string ParentCustomerNumber = "_txtParentCustomerNumber";
            public const string TicketNumber = "_txtTicketNumber";
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string CustomerName = "_txtCustomerName";
            public const string CustomerFEDID = "_txtCustomerFEDID";
            public const string InvoiceNumber = "_txtEllisInvoiceNo";
            public const string BillingAddressZip = "_txtBillingAddressZip";
            public const string CustomerAddress = "_txtCustomerAddress";
            public const string BillingAddressCity = "_txtBillingAddressCity";
            public const string FedId = "_optFedIDType";

            public const string Section = "_cmbSection";
            public const string SubSection = "_cmbSubSection";
            public const string BillingState = "_cboBillingState";
            public const string AccountManagedBy = "_cboAccountManagedBy";
            public const string Branch = "_cmbBranch";
            public const string FromDate = "_calendarFrom";
            public const string ToDate = "_calendarTo";
        }
    }
}