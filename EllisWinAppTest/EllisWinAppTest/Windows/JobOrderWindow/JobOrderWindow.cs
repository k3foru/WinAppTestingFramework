using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;
using System;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    public class JobOrderWindow : AppContext
    {
        public static void ClickOnCreateJobOrder()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.File });
            var joborder = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.JobOrder });
            var cjoborder = joborder.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.CJobOrder });

            MouseActions.Click(file);
            MouseActions.Click(joborder);
            MouseActions.Click(cjoborder);
        }

        public static UITestControl GetNewJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return joborderWindow;
        }

        private static UITestControl GetCreateJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            return joborderWindow;
        }

        public static bool VerifyCreateJobOrderWindowDisplayed()
        {
            var joWindow = GetCreateJobOrderWindowProperties();
            return joWindow.Exists;
        }

        private static UITestControlCollection GetJobOrderEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetCreateJobOrderWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }


        public static void EnterJobOrderData(DataRow data)
        {
            var getJobOrderControlCollection = GetJobOrderEditControlCollection();

            foreach (var control in getJobOrderControlCollection)
            {
                if (control.FriendlyName != null)
                {
                    switch (control.FriendlyName.ToString())
                    {
                        case JobOrderConstants.CustNumber:
                            Actions.SetText(control, data.ItemArray[3].ToString());
                            break;

                        case JobOrderConstants.CustName:
                            Actions.SetText(control, data.ItemArray[4].ToString());
                            break;

                        case JobOrderConstants.PhoneNumber:
                            control.SetFocus();
                            Actions.SetText(control, "");
                            Actions.SendText("{HOME}");
                            Actions.SendText(data.ItemArray[5].ToString());
                            break;

                        case JobOrderConstants.FebTaxId:
                            control.SetFocus();
                            Actions.SendText("");
                            Actions.SendText("{HOME}");
                            Actions.SendText(data.ItemArray[6].ToString());
                            //Actions.SetText(control, data.ItemArray[6].ToString());
                            break;

                        default:
                            Console.WriteLine(control.FriendlyName.ToString());
                            break;
                    }
                }
            }
        }

        public static void EnterJobOrderFindQuoteData(DataRow data)
        {
            var getJobOrderControlCollection = GetJobOrderEditControlCollection();

            foreach (var control in getJobOrderControlCollection)
            {
                if (control.FriendlyName != null)
                {
                    switch (control.FriendlyName.ToString())
                    {
                        case "Text area":
                            Actions.SetText(control, data.ItemArray[9].ToString());
                            break;


                        //case "mskFederalTaxId":
                        //    Actions.SendText("");
                        //    Actions.SendText("{HOME}");
                        //    //Actions.SendText(control, data.ItemArray[6].ToString());
                        //    Actions.SetText(control, data.ItemArray[6].ToString());
                        //    break;

                        default:
                            Console.WriteLine(control.FriendlyName.ToString());
                            break;
                    }
                }
            }
        }

        public static void ClickOnButton(string btnName)
        {
            Factory.ClickOnButton(GetCreateJobOrderWindowProperties(), btnName);
        }

        public static void ClickOnContinueBtn()
        {
            Playback.Wait(3000);
            var jobSearchWindow = GetCreateJobOrderWindowProperties();
            var continueBtn = Actions.GetWindowChild(jobSearchWindow, "btnCreateJobOrder");
            Mouse.Click(continueBtn);
        }

        //public static void ClickOnGoBtn()
        //{
        //    Playback.Wait(3000);
        //    var jobSearchWindow = GetCreateJobOrderWindowProperties();
        //    var butGroup = jobSearchWindow.Container.SearchFor<WinGroup>(new { Name = "" });
        //    var searchBtn = butGroup.Container.SearchFor<WinButton>(new { Name = "GO" });
        //    MouseActions.Click(searchBtn);
        //}


        


        public static void CloseJobOrderProfileWindow()
        {
            var newJobOrder = JobOrderWindow.GetNewJobOrderWindowProperties();
            var joNum = Actions.GetWindowChild(newJobOrder, "lblJobOrderNumber");
            Console.WriteLine("Closing Profile Window with #"+joNum.GetProperty("Name"));
            var closeBtn = Actions.GetWindowChild(newJobOrder, "btnClose");
            Mouse.Click(closeBtn);
        }

        public static bool CreateNewJobOrder(DataRow dataRow)
        {
            JobOrderWindow.ClickOnCreateJobOrder();
            JobOrderWindow.EnterJobOrderData(dataRow);
            JobOrderWindow.ClickOnButton("Search");

            Playback.Wait(3000);
            JobOrderWindow.ClickOnContinueBtn();

            // Find Quote Tab/Window
            Playback.Wait(3000);
            JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
            JobOrderFindQuoteWindow.ClickOnButton("GO");
            Playback.Wait(2000);
            JobOrderWindow.ClickOnContinueBtn();

            // Enter Basic Job Order Details
            BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
            BasicJobInformationWindow.ClickOnContinueBtn();

            // Enter Schedule And Additional Charges Details
            ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
            ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

            // Enter Order Notes in Schedule And Additional Charges window
            ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

            // Focus back to Schedule And Additional Charges window
            ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

            //Enter data in Requirements window
            RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
            RequirementsWindow.ClickOnButton("Continue >");
            Playback.Wait(3000);

            //Enter data in Pre-Qualifying Requirements Window

            PreQualifyingQuestionsWindow.ClickonSaveButton();
            PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
            PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
            Playback.Wait(3000);
            var status = PreQualifyingQuestionsWindow.HandleAlertWindow();

            ////Close New JobOrder Profile Window
            //Playback.Wait(3000);
            //var newJobOrder = JobOrderWindow.GetNewJobOrderWindowProperties();
            //var joNum = Actions.GetWindowChild(newJobOrder, "lblJobOrderNumber");
            //Globals.JobOrderNo = joNum.GetProperty("Name").ToString();
            //MouseActions.ClickButton(newJobOrder,"btnClose");


            return status;
        }
    }

    public class JobOrderConstants
    {
        public const string CustNumber = "mskCustomerNumber";
        public const string CustName = "txtCustomerName";
        public const string PhoneNumber = "mskPhoneNumber";
        public const string FebTaxId = "mskFederalTaxId";
        // public const string CustName = "txtCustomerName";
        //public const string CustName = "txtCustomerName";
    }
}