//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class DropDownActions : AppContext
    {
        public static int GetDropDownItemsCount(WinComboBox control)
        {
            var count = control.Items.Count;
            return count;
        }

        public static List<string> GetAllDropDownItems(WinComboBox dropDown)
        {
            var items =
                dropDown.GetProperty("Items") as
                    UITestControlCollection;

            return items.Select(item => item.Name).ToList();
        }

        public static string GetCurrentItem(WinComboBox dropDown)
        {
            var currentValue = dropDown.SelectedItem;
            return currentValue;
        }

        public static string GetCurrentItem(UITestControl window, string controlName)
        {
            var control = Actions.GetWindowChild(window, controlName);
            return GetCurrentItem((WinComboBox) control);
        }

        //public static void SelectDropdownByText(UITestControl dropdownControl, string text)
        //{
        //    int i = 0;
        //    var control = (WinComboBox)dropdownControl;
        //    do
        //    {
        //        MouseActions.Click(control); 
        //        SendKeys.SendWait("^A");
        //        SendKeys.SendWait("{DELETE}"); 
        //        if (string.IsNullOrEmpty(control.SelectedItem))
        //        {
        //            SendKeys.SendWait(text);
        //            Playback.Wait(500);
        //        }

        //        if (!control.SelectedItem.Equals(text))
        //        {
        //            var first = text[0];
        //            SendKeys.SendWait(first.ToString());
        //            i++;
        //        }
        //    } while (!control.SelectedItem.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase) && i!=15);
        //}

        public static void SelectDropdownByText(UITestControl dropdownControl, string text)
        {
            var control = (WinComboBox)dropdownControl;
            control.SetFocus();
            string item;

            item = control.SelectedItem;
            //if (string.IsNullOrEmpty(item))
            //{
            //    if(false)
            //        item = control.GetProperty("Value").ToString();
            //    else
            //    {
            //        item = control.GetProperty("Value").ToString();
            //    }
            //}

            var i = 0;

            if (string.IsNullOrEmpty(item))
            {
                control.SetFocus();
                SendKeys.SendWait(text);
                Playback.Wait(500);
                item = control.SelectedItem;
            }

            if (!string.IsNullOrEmpty(item) &&
                !item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
            {
                control.SetFocus();
                SendKeys.SendWait(text);
                item = control.SelectedItem;

                if (!item.TrimEnd().Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase))
                {
                    do
                    {
                        var first = text[0];
                        control.SetFocus();
                        SendKeys.SendWait(first.ToString());
                        SendKeys.SendWait("{TAB}");
                        Playback.Wait(150);
                        i++;
                    } while (
                        !control.SelectedItem.TrimEnd()
                            .Equals(text.TrimEnd(), StringComparison.InvariantCultureIgnoreCase) &&
                        i <= 20);
                }
            }
        }

        public static void SelectDropdownByText(UITestControl window, string controlName, string text)
        {
            var control = Actions.GetWindowChild(window, controlName);
            control.SetFocus();
            SelectDropdownByText(control, text);
        }

        //public static void SelectFromDropDown(UITestControlCollection dropDownCollection, string dropdownName,
        //    string data)
        //{
        //    foreach (var control in dropDownCollection.Cast<WinComboBox>()
        //        .Where(control => control.FriendlyName != null && control.ControlName.Equals(dropdownName)
        //            && (!control.SelectedItem.Equals(data, StringComparison.InvariantCultureIgnoreCase))))
        //    {
        //        SetDropDown(control, data);
        //        break;
        //    }
        //}

        //private static void SetDropDown(WinComboBox dropdownControl, string text)
        //{
        //    if (string.IsNullOrEmpty(dropdownControl.SelectedItem))
        //    {
        //        MouseActions.Click(dropdownControl);
        //        Actions.SendText(text);
        //    }
        //    else if (!dropdownControl.SelectedItem.Equals(text))
        //    {
        //        MouseActions.Click(dropdownControl);
        //        Actions.SendText(text);
        //        //foreach (var t in dropdownControl.Items.Where(t => t.Name == text))
        //        //{
        //        //    dropdownControl.SelectedItem = text;
        //        //    break;
        //        //}
        //    }
        //}
    }
}
