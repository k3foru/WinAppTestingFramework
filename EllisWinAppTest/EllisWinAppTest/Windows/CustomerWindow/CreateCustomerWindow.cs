using System.Data;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.CustomerWindow
{
    public class CreateCustomerWindow : AppContext
    {
        public static void ClickOnCreateCustomer()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new {Name = EllisHomeConstants.File});
            var customer = file.Container.SearchFor<WinMenuItem>(new {Name = EllisHomeConstants.Customer});
            var ccustomer = customer.Container.SearchFor<WinMenuItem>(new {Name = EllisHomeConstants.CCustomer});

            MouseActions.Click(file);
            MouseActions.Click(customer);
            MouseActions.Click(ccustomer);
        }

        private static UITestControl GetCreateCustomerWindowProperties()
        {
            var ccustomerWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Create Customer"});
            return ccustomerWindow;
        }

        private static UITestControl GetExistingCustomerFoundWindowProperties()
        {
            var ccustomerWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Existing Customer Found" });
            return ccustomerWindow;
        }

        public static void EnterCustomerData(DataRow data, string fedid, string cod)
        {
            Thread.Sleep(2000);
            var customerWindow = GetCreateCustomerWindowProperties();

            //CheckAsAbove(data.ItemArray[4].ToString());

            Globals.CustomerName = Generator.GenerateNewName(data.ItemArray[3].ToString());
            //Globals.CustomerLegalName = Factory.GenerateNewName(data.ItemArray[5].ToString());

            Actions.SetText(customerWindow, CCustomerConstants.CustomerName, Globals.CustomerName);
            Actions.SetText(customerWindow, CCustomerConstants.CustomerLegalName, Globals.CustomerName);

            var credit = Actions.GetWindowChild(customerWindow, CCustomerConstants.CreditTerms);

            switch (cod)
            {
                case null:
                case "Credit":
                    DropDownActions.SelectDropdownByText(credit, "Credit");
                    break;
                case "COD":
                    DropDownActions.SelectDropdownByText(credit, "COD");
                    break;
            }

            DropDownActions.SelectDropdownByText(credit, "COD");

            if (fedid != null)
            {
                var control = string.Empty;
                var applicantWindow = GetCreateCustomerWindowProperties();
                switch (fedid)
                {
                    case "Fed":
                        control = "Fed ID";
                        break;
                    case "SSN":
                        control = "SSN";
                        break;
                    case "ITIN":
                        control = "ITIN";
                        break;
                }

                var rbtControl = applicantWindow.Container.SearchFor<WinRadioButton>(new {Name = control});
                var children = rbtControl.GetChildren();
                foreach (var child in children)
                {
                    if (child.Name.Equals(control))
                    {
                        child.SetProperty("Selected", true);

                        Playback.Wait(2000);
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{HOME}");
                        Actions.SendText(data.ItemArray[8].ToString());
                    }
                }
            }

            Actions.SetText(customerWindow, CCustomerConstants.ContactName, data.ItemArray[9].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.PhoneNumber, data.ItemArray[10].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.CellNumber, data.ItemArray[11].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.FaxNumber, data.ItemArray[12].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.Email, data.ItemArray[13].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.AddressLine1, data.ItemArray[14].ToString());
            Actions.SetText(customerWindow, CCustomerConstants.AddressLine2, data.ItemArray[15].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[16].ToString()))
            {
                var country = Actions.GetWindowChild(customerWindow, CCustomerConstants.Country);
                DropDownActions.SelectDropdownByText(country, data.ItemArray[16].ToString());
            }

            Actions.SetText(customerWindow, CCustomerConstants.ZipExt, data.ItemArray[19].ToString());

            if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
            {
                var state = Actions.GetWindowChild(customerWindow, CCustomerConstants.State);
                DropDownActions.SelectDropdownByText(state, data.ItemArray[17].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
            {
                var zip = Actions.GetWindowChild(customerWindow, CCustomerConstants.Zip);
                DropDownActions.SelectDropdownByText(zip, data.ItemArray[18].ToString());
            }

            Playback.Wait(2000);

            if (!string.IsNullOrEmpty(data.ItemArray[20].ToString()))
            {
                var city = Actions.GetWindowChild(customerWindow, CCustomerConstants.City);
                DropDownActions.SelectDropdownByText(city, data.ItemArray[20].ToString());
            }
        }

        public static void HandleExistingFEDCustomer()
        {
            var window = GetExistingCustomerFoundWindowProperties();

            if (window.Exists)
            {
                MouseActions.ClickButton(window, CCustomerConstants.CreateCustomer);
            }
        }

        private static void CheckAsAbove(string check)
        {
            var applicantWindow = GetCreateCustomerWindowProperties();
            Actions.SetCheckBox(applicantWindow, CCustomerConstants.SameAsAbove, check);
        }

        public static void ClickSave()
        {
            var customerWindow = GetCreateCustomerWindowProperties();
            var btn = customerWindow.Container.SearchFor<WinButton>(new {Name = "Save"});
            MouseActions.Click(btn);
        }

        public static void ClickCancel()
        {
            var customerWindow = GetCreateCustomerWindowProperties();
            var btn = customerWindow.Container.SearchFor<WinButton>(new {Name = "Cancel"});
            MouseActions.Click(btn);
        }

        public class CCustomerConstants
        {
            public const string CustomerName = "_txtDoingBusinessAsName";
            public const string CustomerLegalName = "_txtCustomerBusinessName";
            public const string CreditTerms = "_cboCreditTerms";
            public const string TaxId = "_txtFedID";
            public const string ContactName = "_txtWorkContactName";
            public const string PhoneNumber = "_PrimaryPhoneNumber";
            public const string CellNumber = "_CellPhoneNumber";
            public const string FaxNumber = "_FaxNumber";
            public const string Email = "_txtEmail";
            public const string AddressLine1 = "_txtAddressLine1";
            public const string AddressLine2 = "_txtAddressLine2";
            public const string Country = "_cboCountry";
            public const string State = "_cboState";
            public const string Zip = "_cboPostalCodes";
            public const string ZipExt = "_txtPostalCodePlusFour";
            public const string City = "_cboCity";

            public const string SameAsAbove = "_chkSameAsAbove";
            public const string CreateCustomer = "btnCreateCustomer";
        }
    }
}