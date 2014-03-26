using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System;
using Point = System.Drawing.Point;

namespace EllisWinAppTest.Helpers
{
    public class Factory : AppContext
    {
        public static void GetWinProperties()
        {
            EllisWindow = App.SearchFor<WinWindow>(
                new { Name = "Ellis" });
        }

        //method to click on button in any window
        // Provide window instance as 1st parameter and button name as 2nd parameter
        //public static void ClickOnButton(UITestControl windowInstence, string butName)
        //{
        //    var group = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
        //    var btnControl = group.Container.SearchFor<WinButton>(new { Name = butName });
        //    var btnControlcollection = Actions.GetControlCollection(btnControl);

        //    foreach (var control in btnControlcollection)
        //    {
        //        //MouseActions.Click(control);
        //        control.SetFocus();
        //        SendKeys.SendWait("{ENTER}");
        //    }
        //}

        public static void ClickOnButton(UITestControl windowInstence, string butName)
        {
            var winGroup = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
            var btnControl = winGroup.Container.SearchFor<WinButton>(new { Name = butName });
            btnControl.SetFocus();
            MouseActions.Click(btnControl);
        }

        public static ApplicationUnderTest GetAUT()
        {
            var processes = Process.GetProcessesByName("Ellis");
            App = processes.Length > 0
                ? ApplicationUnderTest.FromProcess(processes[0])
                : ApplicationUnderTest.Launch(TestData.Path, TestData.AltPath);

            return App;
        }

        public static Point GetMouseCoOrdinates(UITestControl control)
        {
            var clickPoints = new Point(control.BoundingRectangle.Width / 2 + control.BoundingRectangle.X,
                control.BoundingRectangle.Height / 2 + control.BoundingRectangle.Y);
            return clickPoints;
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }



        //public static void SelectDropdownByText(UITestControl dropdownControl, string text)
        //{
        //    var control = (WinComboBox)dropdownControl;
        //    control.SetFocus();
        //    string item;

        //    item = control.SelectedItem;
        //    //if (string.IsNullOrEmpty(item))
        //    //{
        //    //    if(false)
        //    //        item = control.GetProperty("Value").ToString();
        //    //    else
        //    //    {
        //    //        item = control.GetProperty("Value").ToString();
        //    //    }
        //    //}

        //    var i = 0;

        //    if (string.IsNullOrEmpty(item))
        //    {
        //        control.SetFocus();
        //        SendKeys.SendWait(text);
        //        Playback.Wait(500);
        //        item = control.SelectedItem;
        //    }

        //    if (!string.IsNullOrEmpty(item))
        //    {
        //        if (!item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
        //        {
        //            control.SetFocus();
        //            SendKeys.SendWait(text);
        //            item = control.SelectedItem;

        //            if (!item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
        //            {
        //                do
        //                {
        //                    var first = text[0];
        //                    control.SetFocus();
        //                    SendKeys.SendWait(first.ToString());
        //                    SendKeys.SendWait("{TAB}");
        //                    Playback.Wait(150);
        //                    i++;
        //                } while (
        //                    !control.SelectedItem.TrimEnd()
        //                        .Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase) &&
        //                    i <= 20);
        //            }
        //        }
        //    }
        //}

        //public static void SelectDropdownByText(UITestControl dropdownControl, string controlName, string text)
        //{
        //    var dd = Actions.GetWindowChild(dropdownControl, controlName);
        //    var control = (WinComboBox)dd;
        //    control.SetFocus();
        //    string item;

        //    item = control.SelectedItem;
        //    var i = 0;

        //    if (string.IsNullOrEmpty(item))
        //    {
        //        control.SetFocus();
        //        SendKeys.SendWait(text);
        //        Playback.Wait(500);
        //        item = control.SelectedItem;
        //    }

        //    if (!string.IsNullOrEmpty(item))
        //    {
        //        if (!item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
        //        {
        //            control.SetFocus();
        //            SendKeys.SendWait(text);
        //            item = control.SelectedItem;

        //            if (!item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
        //            {
        //                do
        //                {
        //                    var first = text[0];
        //                    control.SetFocus();
        //                    SendKeys.SendWait(first.ToString());
        //                    SendKeys.SendWait("{TAB}");
        //                    Playback.Wait(150);
        //                    i++;
        //                } while (
        //                    !control.SelectedItem.TrimEnd()
        //                        .Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase) &&
        //                    i <= 20);
        //            }
        //        }
        //    }
        //}

        //public static void SelectDropdownByValue(UITestControl dropdownControl, string text)
        //{
        //    var control = (WinComboBox)dropdownControl;
        //    //MouseActions.Click(control);
        //    control.SetFocus();
        //    var item = control.GetProperty("Value");
        //    var i = 0;

        //    if (string.IsNullOrEmpty(item.ToString()))
        //    {
        //        SendKeys.SendWait(text);
        //        Playback.Wait(500);
        //    }

        //    if (!string.IsNullOrEmpty(item.ToString()) &&
        //        control.SelectedItem.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        SendKeys.SendWait(text);
        //        do
        //        {
        //            var first = text[0];
        //            control.SetFocus();
        //            SendKeys.SendWait(first.ToString());
        //            SendKeys.SendWait("{TAB}");
        //            Playback.Wait(150);
        //            i++;
        //        } while (
        //            !control.SelectedItem.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase) &&
        //            i <= 20);
        //    }
        //}

        //////public static void SelectFromDropDown(UITestControlCollection dropDownCollection, string dropdownName, string data)
        //////{
        //////    foreach (
        //////        WinComboBox control in
        //////            dropDownCollection.Where(
        //////                control => control.ControlName != null || control.ControlName != "")
        //////                .Where(
        //////                    control =>
        //////                        control.ControlName != null &&
        //////                        control.ControlName.Equals(dropdownName)))
        //////    {
        //////        Factory.SetDropDown(control, data);
        //////        break;
        //////    }
        //////}

        ////public static void SelectFromDropDown(UITestControlCollection dropDownCollection, string dropdownName,
        ////    string data)
        ////{
        ////    foreach (var uiTestControl in dropDownCollection)
        ////    {
        ////        var control = (WinComboBox) uiTestControl;
        ////        if (control.FriendlyName != null &&
        ////            control.ControlName.Equals(dropdownName) &&
        ////            (!control.SelectedItem.Equals(data, StringComparison.InvariantCultureIgnoreCase)))
        ////        {
        ////            SetDropDown(control, data);
        ////            break;
        ////        }
        ////    }
        ////}

        ////private static void SetDropDown(WinComboBox dropdownControl, string text)
        ////{
        ////    if (string.IsNullOrEmpty(dropdownControl.SelectedItem))
        ////    {
        ////        MouseActions.Click(dropdownControl);
        ////        Actions.SendText(text);
        ////    }
        ////    else if (!dropdownControl.SelectedItem.Equals(text))
        ////    {
        ////        MouseActions.Click(dropdownControl);
        ////        Actions.SendText(text);
        ////        //foreach (var t in dropdownControl.Items.Where(t => t.Name == text))
        ////        //{
        ////        //    dropdownControl.SelectedItem = text;
        ////        //    break;
        ////        //}
        ////    }
        ////}

        public static void GetAllControlNames(UITestControl control)
        {
            var group = control.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);

            var dDownControl = group.Container.SearchFor<WinComboBox>(new { Name = "" });
            var ddownControlCollection = Actions.GetControlCollection(dDownControl);

            var btnControl = group.Container.SearchFor<WinButton>(new { Name = "" });
            var btnControlCollection = Actions.GetControlCollection(btnControl);

            //var chkboxCotrol = group.Container.SearchFor<WinCheckBox>(new { Name = "Same as above" });
            //var cheboxControlCollection = Actions.GetControlCollection(chkboxCotrol);

            var tabCotrol = group.Container.SearchFor<WinTabPage>(new { Name = "Skills" });
            var tabControlCollection = tabCotrol.FindMatchingControls();

            var chkboxCotrol = group.Container.SearchFor<WinTable>(new { Name = "" });
            var cheboxControlCollection = chkboxCotrol.FindMatchingControls();

            TextWriter tsw = new StreamWriter(@"C:\ControlCollection.txt");

            //Writing text to the file.
            tsw.WriteLine("-----------------------Text Box Controls\n\n\n");
            foreach (var edit in editControlcollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Drop Down Controls\n\n\n");
            foreach (var edit in ddownControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }


            tsw.WriteLine("Button Controls\n\n\n");
            foreach (var edit in btnControlCollection)
            {
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Table Controls\n\n\n");
            foreach (var edit in cheboxControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Tab Controls\n\n\n");
            foreach (var edit in tabControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }
            //Close the file.
            tsw.Close();
        }


        //public static string GenerateNewName(string name)
        //{
        //    var n = new Random().Next(99999);
        //    var newName = name + n;
        //    return newName;
        //}

        //public static string GenerateSSNNumber()
        //{
        //    var rnd = new Random();
        //    var num = rnd.Next(100000000, 999999999).ToString();
        //    return num;
        //}

        //public static string GenerateRandomNumber(int length)
        //{
        //    var rnd = new Random();
        //    var num = rnd.Next(0, length).ToString();
        //    return num;
        //}

        //public static String GetTextBetween(String source, String leftWord, String rightWord)
        //{
        //    return
        //        Regex.Match(source, String.Format(@"{0}\s(?<words>[\w\s\/1234567890]+)\s{1}", leftWord, rightWord),
        //                    RegexOptions.IgnoreCase).Groups["words"].Value;
        //}

        //public static void SelectRadioButton(WinRadioButton radioButtonControl, string radioButtonGroupName)
        //{
        //    var btnCont = Actions.GetControlCollection(radioButtonControl);

        //    foreach (var control in btnCont)
        //    {
        //        if (control.GetProperty("ControlName").ToString() == radioButtonGroupName)
        //            Mouse.Click(control);
        //    }
        //}

        //public static void SelectRadioButton(WinRadioButton radioButtonControl)
        //{
        //    var btnCont = Actions.GetControlCollection(radioButtonControl);

        //    foreach (var control in btnCont)
        //    {
        //        Mouse.Click(control);
        //    }
        //}

        //public static bool OpenRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        //{
        //    var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
        //    var table = (WinTable)tableName;
        //    foreach (var rowC in table.Rows)
        //    {
        //        rowC.SetFocus();
        //        var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
        //        var callValue = rowHeader.GetProperty("Value").ToString();

        //        if (callValue == columnValue)
        //        {
        //            Mouse.Click(rowHeader);
        //            Mouse.DoubleClick(rowHeader);
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        //public static bool SelectRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        //{
        //    var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
        //    var table = (WinTable)tableName;

        //    foreach (var rowC in table.Rows)
        //    {
        //        rowC.SetFocus();
        //        var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
        //        var callValue = rowHeader.GetProperty("Value").ToString();

        //        if (callValue == columnValue)
        //        {
        //            Mouse.Click(rowHeader);
        //            var rowC1 = table.Container.SearchFor<WinRow>(new {Name = rowC.GetProperty("Name").ToString()});
        //            rowC1.SetFocus();
        //            var select = rowC1.Container.SearchFor<WinCell>(new { Name = "Select" });
        //            select.SetFocus();
        //            if (select.GetProperty("Value").ToString()!= "True")
        //            {
        //                Mouse.Click(select);
        //            }
        //            return true;
        //        }
        //    }
        //    return false;
        //}
    }



    public static class Globals
    {
        public static string CustomerName { get; set; }
        public static string CustomerLegalName { get; set; }
        public static string CustomerContact { get; set; }
        public static string WorkerName { get; set; }
        public static string SSN { get; set; }
        public static string QuotedBy { get; set; }
        public static string JobOrderNo { get; set; }

        public static string Temp { get; set; }
    }

    public static class GlobalWindows
    {
        public static WinWindow CustomerProfileWindow { get; set; }
        public static WinWindow CustomerContactInfoWindow { get; set; }
        public static WinWindow CustomerSearchResultsWindow { get; set; }
        public static WinWindow CustomerWorkerComp { get; set; }
        public static WinWindow CustomerAddNote { get; set; }
        public static WinWindow COIPreviewWindow { get; set; }
        public static WinWindow DocumentOnFileWindow { get; set; }
        public static WinWindow DocumentOnFileFindFileWindow { get; set; }
        public static WinWindow QuoteProfileWindow { get; set; }
        public static WinWindow JobOrderProfileWindow { get; set; }

        public static WinWindow DialogWindow { get; set; }
    }

}