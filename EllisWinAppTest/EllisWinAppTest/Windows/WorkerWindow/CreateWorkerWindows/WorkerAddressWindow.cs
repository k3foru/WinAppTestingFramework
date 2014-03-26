using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerAddressWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerAddressWindowProperties()
        {
            var wAddresWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return wAddresWindow;
        }

        private static UITestControl GetRejectPopUpProperties()
        {
            var rejectPopUp =
                App.Container.SearchFor<WinWindow>(new {ClassName = "WindowsForms10.Window.8.app.0.265601d"});
            return rejectPopUp;
        }

        #endregion

        #region Address Methods

        public static bool EnterAddressData(DataRow data)
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();

            if (wAddresWindow.Exists)
            {
                var suffix = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Suffix);
                //suffix.SetFocus();
                //SendKeys.SendWait(data.ItemArray[11].ToString());
                DropDownActions.SelectDropdownByText(suffix, data.ItemArray[11].ToString());

                var gender = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Gender);
                //gender.SetFocus();
                //SendKeys.SendWait(data.ItemArray[12].ToString());
                DropDownActions.SelectDropdownByText(gender, data.ItemArray[12].ToString());

                var race = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Race);
                //race.SetFocus();
                //SendKeys.SendWait(data.ItemArray[13].ToString());
                DropDownActions.SelectDropdownByText(race, data.ItemArray[13].ToString());

                var jobCategory = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.JobCategory);
                //jobCategory.SetFocus();
                //SendKeys.SendWait(data.ItemArray[14].ToString());
                DropDownActions.SelectDropdownByText(jobCategory, data.ItemArray[14].ToString());

                var rAddress = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RStreetAddress);
                Actions.SetText(rAddress, data.ItemArray[15].ToString());

                var rSuiteNo = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RApartment);
                Actions.SetText(rSuiteNo, data.ItemArray[16].ToString());

                var rState = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RState);
                //rState.SetFocus();
                //SendKeys.SendWait(data.ItemArray[17].ToString());
                DropDownActions.SelectDropdownByText(rState, data.ItemArray[17].ToString());

                var rZip = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RZip);
                DropDownActions.SelectDropdownByText(rZip, data.ItemArray[18].ToString());

                var rCity = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RCity);
                DropDownActions.SelectDropdownByText(rCity, data.ItemArray[19].ToString());

                var mAddress = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MStreetAddress);
                Actions.SetText(mAddress, data.ItemArray[20].ToString());

                var mSuiteNo = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MApartment);
                Actions.SetText(mSuiteNo, data.ItemArray[21].ToString());

                var mState = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MState);
                DropDownActions.SelectDropdownByText(mState, data.ItemArray[22].ToString());

                var mZip = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MZip);
                DropDownActions.SelectDropdownByText(mZip, data.ItemArray[23].ToString());

                var mCity = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MCity);
                DropDownActions.SelectDropdownByText(mCity, data.ItemArray[24].ToString());

                return true;
            }

            return false;

        }


        public static bool ClickOnContinueBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.ContinueBtn);
                Mouse.Click(continueBtn);
                return true;
            }
            
            return false;
        }

        public static bool ClickOnBackBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
            {
                var backBtn = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.BackBtn);
                Mouse.Click(backBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnRejectBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
            {
                var rejectBtn = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RejectBtn);
                Mouse.Click(rejectBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCancelBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Reject PopUp Methods

        public static bool ClickOnDoneBtnReject()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
            {
                MouseActions.ClickButton(rejectPopUp, "btnDone");
                return true;
            
            }
            return false;
           
        }

        public static bool ClickOnBackBtnReject()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
            {
                MouseActions.ClickButton(rejectPopUp, "btnBack");
                return true;
            }
            return false;
            
        }

        public static bool EnterDataInRejectPopUp()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
            {
                var rejectCmb = Actions.GetWindowChild(rejectPopUp, "ultraRejectCombo");
                DropDownActions.SelectDropdownByText(rejectCmb, "Others");
                return true;
            }
            return false;

        }

        #endregion

        #region Controls

        private class AWorkerConstants
        {
            public const string Suffix = "cmbSuffix";
            public const string Gender = "cmbGender";
            public const string Race = "cmbRace";
            public const string JobCategory = "cmbJobCategory";
            public const string RStreetAddress = "txtResidenceStreet";
            public const string RApartment = "txtResidenceApt";
            public const string RState = "cmbResidenceState";
            public const string RZip = "cmbResidenceZipcode";
            public const string RCity = "cmbResidenceCity";
            public const string MStreetAddress = "txtMailingStreet";
            public const string MApartment = "txtMailingApt";
            public const string MState = "cmbMailingState";
            public const string MZip = "cmbMailingZipCode";
            public const string MCity = "cmbMailingCity";
            public const string BackBtn = "_btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string AddressChkBox = "chkAddress";
            public const string DeliveryRadioBtn = "OptDelivery";
        }

        #endregion
    }
}