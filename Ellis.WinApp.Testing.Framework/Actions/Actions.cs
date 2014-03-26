//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class Actions : AppContext
    {
        public static void SetText(UITestControl control, string text)
        {
            control.WaitForControlReady();
            control.SetFocus();
            MouseClickOnCoordinates(control);
            if (!control.GetProperty("Value").Equals(null))
            {
                SendKeys.SendWait("^A");
                SendKeys.SendWait("{DELETE}");
            }
            SendKeys.SendWait("{HOME}");
            Playback.Wait(1000);
            SendKeys.SendWait(text);
        }

        public static void SendEnter()
        {
            SendKeys.SendWait("{Enter}");
        }

        public static void SendTab()
        {
            SendKeys.SendWait("{TAB}");
        }

        public static void SendAltF4()
        {
            SendText("%{F4}");
        }

        public static void SendText(string text)
        {
            SendKeys.SendWait(text);
        }

        //public static void MouseClickOnCoordinates(UITestControl control)
        //{
        //  control.WaitForControlReady();
        //  control.SetFocus();
        //  Point screenCoordinate = new Point();
        //  // ISSUE: explicit reference operation
        //  // ISSUE: variable of a reference type
        //  Point& local = @screenCoordinate;
        //  Rectangle boundingRectangle = control.BoundingRectangle;
        //  int num1 = boundingRectangle.Width / 2;
        //  boundingRectangle = control.BoundingRectangle;
        //  int x1 = boundingRectangle.X;
        //  int x2 = num1 + x1;
        //  boundingRectangle = control.BoundingRectangle;
        //  int num2 = boundingRectangle.Height / 2;
        //  boundingRectangle = control.BoundingRectangle;
        //  int y1 = boundingRectangle.Y;
        //  int y2 = num2 + y1;
        //  // ISSUE: explicit reference operation
        //  ^local = new Point(x2, y2);
        //  Mouse.Click(screenCoordinate);
        //}

        public static void MouseClickOnCoordinates(UITestControl control)
        {
            control.WaitForControlReady();
            control.SetFocus();
            var clickPoints = new Point(control.BoundingRectangle.Width / 2 + control.BoundingRectangle.X,
                control.BoundingRectangle.Height / 2 + control.BoundingRectangle.Y);
            Mouse.Click(clickPoints);
        }

        public static UITestControlCollection GetControlCollection(WinEdit editControl)
        {
            return editControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinText editControl)
        {
            return editControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinRadioButton radioControl)
        {
            return radioControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinComboBox dropDoBoxControl)
        {
            return dropDoBoxControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinCheckBox checkControl)
        {
            return checkControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinToolTip tooltipControl)
        {
            return tooltipControl.FindMatchingControls();
        }

        public static UITestControlCollection GetControlCollection(WinButton btnControl)
        {
            return btnControl.FindMatchingControls();
        }

        public static void SetCheckBox(WinCheckBox checkbox, string check)
        {
            checkbox.WaitForControlReady();
            checkbox.SetFocus();
            checkbox.Checked = Convert.ToBoolean(check);
        }

        public static UITestControl GetWindowChild(UITestControl control, string controlName)
        {
            return CodedUIExtension.SearchFor<WinWindow>(control.Container, new
            {
                ControlName = controlName
            }).GetChildren()[3];
        }

        public static bool SelectRadioButton(UITestControl cContactWindow, string control)
        {
            using (var enumerator = Enumerable.Where(CodedUIExtension.SearchFor<WinRadioButton>(cContactWindow.Container, new
            {
                Name = control
            }).GetChildren(), child => child.Name.Equals(control)).GetEnumerator())
            {
                if (!enumerator.MoveNext()) return false;
                enumerator.Current.WaitForControlReady();
                enumerator.Current.SetFocus();
                enumerator.Current.SetProperty("Selected", true);
                Playback.Wait(2000);
                return true;
            }
        }

        public static void SelectRadioButton(WinRadioButton radioButtonControl)
        {
            var btnCont = GetControlCollection(radioButtonControl);

            foreach (var control in btnCont)
            {
                control.WaitForControlReady();
                control.SetFocus();
                Mouse.Click(control);
            }
        }

        public static void SetText(UITestControl windowProperties, string control, string data)
        {
            SetText(GetWindowChild(windowProperties, control), data);
        }

        public static void SetCheckBox(UITestControl windowProperties, string control, string check)
        {
            SetCheckBox((WinCheckBox)GetWindowChild(windowProperties, control), check);
        }

        public static void SelectFromListBox(UITestControl windowProperties, string control, string data)
        {
            ListActions.SelectFromList((WinList)GetWindowChild(windowProperties, control), data);
        }

        public static UITestControl GetWindowProperties(UITestControl container, string windowName)
        {
            return CodedUIExtension.SearchFor<WinWindow>(container.Container, new
            {
                Name = windowName
            });
        }

        public static string TrimDate(string trimMe)
        {
            var length = trimMe.IndexOf(' ');
            if (length == 0)
                return string.Empty;
            return length > 0 ? trimMe.Substring(0, length) : trimMe;
        }

        public static string GetTextBetween(string source, string leftWord, string rightWord)
        {
            return Regex.Match(source, string.Format("{0}\\s(?<words>[\\w\\s\\/1234567890]+)\\s{1}", leftWord, rightWord), RegexOptions.IgnoreCase).Groups["words"].Value;
        }
    }
}
