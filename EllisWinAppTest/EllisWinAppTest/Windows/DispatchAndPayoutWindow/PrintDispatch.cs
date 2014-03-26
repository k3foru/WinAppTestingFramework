using System;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.DispatchAndPayoutWindow
{
    class PrintDispatch : AppContext
    {
        public static UITestControl PrintOptionsWindowProperties()
        {
            var printWindow = App.Container.SearchFor<WinWindow>(new { Name = "Print Dispatch Sheet" });
            return printWindow;

        }

        public static string EllisExceptionWindowProperties()
        {
            var exceptionWindow = App.Container.SearchFor<WinWindow>(new { Name = "Ellis Exception Information" });
            if (exceptionWindow.Exists)
            {
                var exceptionTextControl = exceptionWindow.Container.SearchFor<WinText>(new { ControlName = "_messageLabel" });
                var exception = exceptionTextControl.GetProperty("Name").ToString();
                MouseActions.ClickButton(exceptionWindow, "ButtonAccept");
                return exception;
            }

            else
            {
                return null;
            }
        }

        public static bool PrintDispatchWithMap()
        {
            MouseActions.ClickButton(EllisWindow, "btnPrintDispatchSheet");

            var printWindowControl = PrintOptionsWindowProperties();
            //Actions.SelectRadioButton(printWindowControl, "Need Map");
            Playback.Wait(5000);
            MouseActions.ClickButton(printWindowControl, "btnPrint");

            var exp = EllisExceptionWindowProperties();
            if ( exp == null)
            {
                var print = App.Container.SearchFor<WinWindow>(new {Name = "Print"});
                MouseActions.ClickButton(print, "Print");
                Playback.Wait(2000);

                return true;
            }
            else
            {
                MouseActions.ClickButton(printWindowControl, "btnCancel");

                return false;
            }
        }

        public static bool PrintDispatchWithOutMap()
        {
            MouseActions.ClickButton(EllisWindow, "btnPrintDispatchSheet");

            var printWindowControl = PrintOptionsWindowProperties();
            //Actions.SelectRadioButton(printWindowControl, "Do Not Need Map");
            SendKeys.SendWait(" ");
            Playback.Wait(5000);
            MouseActions.ClickButton(printWindowControl, "btnPrint");

            var exp = EllisExceptionWindowProperties();
            if (exp == null)
            {
                var dskInst = UITestControl.Desktop;
                var print = dskInst.Container.SearchFor<WinWindow>(new { Name = "Print" });
                var prnBtn = print.Container.SearchFor<WinButton>(new {Name = "Print"});
                prnBtn.SetFocus();
                Mouse.Click(prnBtn);
                //MouseActions.ClickButton(print, "Print");
                Playback.Wait(2000);
                //MouseActions.ClickButton(print, "Print");
                prnBtn.SetFocus();
                Mouse.Click(prnBtn);

                return true;
            }
            else
            {
                Console.WriteLine(exp);
                MouseActions.ClickButton(printWindowControl, "btnCancel");

                return false;
            }
        }
    }
}
