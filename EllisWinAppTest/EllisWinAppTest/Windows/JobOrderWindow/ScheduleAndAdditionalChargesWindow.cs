using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;
using System;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class ScheduleAndAdditionalChargesWindow : AppContext
    {
        private static UITestControl ScheduleAndAdditionalChargesWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            return joborderWindow;
        }

        private static UITestControl AddOrderNotesWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Order Notes" });
            return joborderWindow;
        }

        public static UITestControlCollection GetButtonColloction(UITestControl windowInstence, string butName)
        {
            var group = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
            var btnControl = group.Container.SearchFor<WinButton>(new { Name = butName });
            var btnControlcollection = Actions.GetControlCollection(btnControl);

            return btnControlcollection;
        }

        private static UITestControlCollection GetScheduleAndAdditionalChargesWindowEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ScheduleAndAdditionalChargesWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetScheduleAndAdditionalChargesWindowDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ScheduleAndAdditionalChargesWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinComboBox>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetScheduleAndAdditionalChargesWindowRadioButtonControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ScheduleAndAdditionalChargesWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinRadioButton>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }
        
        public static void ClickOnContinueBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Continue >");

            foreach (var control in butColloction)
            {
                Mouse.Click(control);
            }
        }

        public static void ClickOnBackBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "< Back");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnSaveBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Save");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnCancelJobOrderBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Cancel Job Order");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnAddNotesBtn()
        {
            //Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var btn = Actions.GetWindowChild(windowInstence, "_btnAddNotes");
            Mouse.Click(btn);

        }

        public static void ClickOnOkBtn()
        {
            //Playback.Wait(3000);
            var windowInstence = AddOrderNotesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "OK");
            foreach (var control in butColloction)
            {
                control.SetFocus();
                SendKeys.SendWait("{ENTER}");
            }
        }

        public static void EnterDataInScheduleAndAdditionalChargesWindow(DataRow data)
        {
            //var ScheduleAndAdditionalChargesWindowControlColloction = GetScheduleAndAdditionalChargesWindowEditControlCollection();
            //Factory.GetAllControlNames(ScheduleAndAdditionalChargesWindowProperties());
            JobOrderSchedule(data);

        }

        public static void JobOrderSchedule(DataRow data)
        {
            var winInst = ScheduleAndAdditionalChargesWindowProperties();
            var tableName = Actions.GetWindowChild(winInst, "_grdJobOrderSchedule");
            var table = (WinTable)tableName;
            //var tablecont = table.Rows.GetNamesOfControls();

            var row = table.Container.SearchFor<WinRow>(new { Name = "Template Add Row" });

            var cell = row.Container.SearchFor<WinCell>(new { Name = "Start Time" });
            Mouse.DoubleClick(cell);
            SendKeys.SendWait(data.ItemArray[58].ToString());

            
            var cell2 = row.Container.SearchFor<WinCell>(new { Name = "Assigned To Branch" });
            Mouse.Click(cell2);
            SendKeys.SendWait(data.ItemArray[59].ToString());

            //var cell3 = row.Container.SearchFor<WinCell>(new { Name = "MAP" });
            //Mouse.Click(cell3);
            

            var cell4 = row.Container.SearchFor<WinCell>(new { Name = "Repeat Dispatch Allowed?" });
            cell4.SetFocus();
            if (cell4.GetProperty("Value").ToString() != data.ItemArray[61].ToString())
            {
                Mouse.Click(cell4);
            }
                

            
            switch (DateTime.Now.DayOfWeek)
            {
                    case DayOfWeek.Saturday:
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[63].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[71].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                    case DayOfWeek.Sunday:
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[71].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                    case DayOfWeek.Monday:
                    
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[71].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                    case DayOfWeek.Tuesday:
                    
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[71].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                    case DayOfWeek.Wednesday:
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[71].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                    case DayOfWeek.Thursday:
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[73].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                    case DayOfWeek.Friday:
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

            }


            //cell = row.Container.SearchFor<WinCell>(new { Name = "Assigned to Branch" });
            //Mouse.DoubleClick(cell);

                
        }

        public static void EnterDataInJobOrderNotesWindow(DataRow data)
        {
            var scheduleAndAdditionalChargesWindowControlColloction = AddOrderNotesWindowProperties();
           // Factory.GetAllControlNames(ScheduleAndAdditionalChargesWindowControlColloction);

            var textOrderNotes = Actions.GetWindowChild(scheduleAndAdditionalChargesWindowControlColloction,
                "_txtOrderNotes");
            Actions.SetText(textOrderNotes, data.ItemArray[77].ToString());

            var dropDownOrderNotes = Actions.GetWindowChild(scheduleAndAdditionalChargesWindowControlColloction,
                "_cmbRequestedBy");
            Playback.Wait(1000);
            DropDownActions.SelectDropdownByText(dropDownOrderNotes, data.ItemArray[78].ToString());
            
            ClickOnOkBtn();
        }
    }
}
