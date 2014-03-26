using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.DispatchAndPayoutWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.JobOrderTests
{
    [CodedUITest]
    public class DispatchAndPayoutTests : AppContext
    {
        [TestMethod]
        public void OpenPendingAssignments()
        {
            EllisHome.Initialize();
            LandingPage.SelectFromToolbar("Dispatch");
            Playback.Wait(3000);
            //OpenByStatus.LoopbyStatus();

            //var status = Windows.JobOrderWindow.JobOrderProfile.OpenJobOrder.SelectJobOrderFromTable();
        }

        [TestMethod]
        public void PrintDispatchWithMap()
        {
            EllisHome.Initialize();
            OpenByStatus.OpenDispatchAndPayoutWindow("All");
            PrintDispatch.PrintDispatchWithMap();

        }

        [TestMethod]
        public void PrintDispatchWithOutMap()
        {
            EllisHome.Initialize();
            OpenByStatus.OpenDispatchAndPayoutWindow("All");
            PrintDispatch.PrintDispatchWithOutMap();
        }

        [TestMethod]
        public void OpenDispatchPayoutByStatus()
        {
            EllisHome.Initialize();
            OpenByStatus.OpenDispatchAndPayoutWindow("All");


        }

        [TestMethod]
        public void VerifyDispatchAndPayoutProfile()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows)
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var calRange = Actions.GetWindowChild(EllisWindow, "btnToggle");
                if (calRange.GetProperty("Name").Equals("Advanced..."))
                    Mouse.Click(calRange);

                var calRangeFrom = Actions.GetWindowChild(EllisWindow, "advancedFromDate");
                calRangeFrom.SetFocus();
                SendKeys.SendWait("{END}");
                SendKeys.SendWait("+{HOME}");
                SendKeys.SendWait("{DEL}");
                SendKeys.SendWait("03102014");
                SendKeys.SendWait("{TAB}");

                var calRangeTo = Actions.GetWindowChild(EllisWindow, "advancedToDate");
                calRangeTo.SetFocus();
                SendKeys.SendWait("{END}");
                SendKeys.SendWait("+{HOME}");
                SendKeys.SendWait("{DEL}");
                SendKeys.SendWait("03202014");
                SendKeys.SendWait("{TAB}");

                TableActions.OpenRecordFromTable(EllisWindow, "grdDispatchJobOrder", "Date", "03/12/2014");
                var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                if (dispatchProfile.Exists)
                {
                    var controlInst = Actions.GetWindowChild(dispatchProfile, "txtQuickAddWorker");
                    Actions.SetText(controlInst, "test");

                    MouseActions.ClickButton(dispatchProfile, "btnAdd");

                    controlInst = Actions.GetWindowChild(dispatchProfile, "txtQuickAddWorker");
                    Actions.SetText(controlInst, "math");

                    MouseActions.ClickButton(dispatchProfile, "btnAdd");


                    // Select Worker from Grid: grdOrderDetails
                    var workerFound = TableActions.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", "test,testone");
                    if (workerFound)
                    {
                        var setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                        Actions.SetCheckBox((WinCheckBox)setWeek, "True");

                        MouseActions.ClickButton(dispatchProfile, "btnAssignWorker");
                        MouseActions.ClickButton(DispatchProfileWindow.AssignWorerWindowProperties(), "btnOK");
                        MouseActions.ClickButton(DispatchProfileWindow.AssignAckWindowProperties(), "_OKButton");
                    }

                    var profileClose = Actions.GetWindowChild(dispatchProfile, "btnCancel");
                    Mouse.Click(profileClose);

                }
            }
        }

        [TestMethod]
        public void VerifyDispatchAndPayoutLandingPage()
        {
            EllisHome.Initialize();
            OpenByStatus.OpenDispatchAndPayoutWindow("All");
            PrintDispatch.PrintDispatchWithMap();

        }
    }
}
