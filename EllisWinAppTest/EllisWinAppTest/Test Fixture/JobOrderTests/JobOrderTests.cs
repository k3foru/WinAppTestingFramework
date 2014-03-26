using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace EllisWinAppTest.JobOrderTests
{
    [CodedUITest]
    public class JoborderTests : AppContext
    {
        [TestMethod]
        public void CreateJobOrder()
        {
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);

            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString().Equals("CreateJobOrder"))
                {
                    //Console.WriteLine(dataRow.ItemArray[1]);
                    var jobOrderCreated = JobOrderWindow.CreateNewJobOrder(dataRow);
                    Assert.IsTrue(jobOrderCreated, "Job order not saved successfully");
                    JobOrderWindow.CloseJobOrderProfileWindow();
                }
            }
        }

        [TestMethod]
        public void OpenExsistingJobOrder()
        {
            EllisHome.Initialize();
            var status = OpenJobOrder.SelectJobOrderFromTable();

            if (status)
                OpenJobOrder.CloseJobOrderProfile();

            Assert.IsTrue(status, "Profile not found");
        }

        [TestMethod]
        public void OpenSpecificJobOrder()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.JobOrderVerify);
            LandingPage.SelectFromToolbar("Job Orders");
            var status = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", "000990957");

            if (status)
                OpenJobOrder.CloseJobOrderProfile();

            Assert.IsTrue(status, "Profile not found");
        }

        [TestMethod]
        public void CopyJobOrderDetails()
        {
            var data = EllisHome.Initialize(ExcelFileNames.JobOrder);
            //open Job Order
            //OpenJobOrder.SelectJobOrderFromTable();

            foreach (var dataRow in data)
            {
                if (dataRow.ItemArray[1].ToString() == "CopyJobOrder" && dataRow.ItemArray[2].ToString() != "")
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());

                    //Copy Job Order Details from opened job order
                    var status = CopyJobOrder.CopyJobOrderDetails();

                    if (status)
                        OpenJobOrder.CloseJobOrderProfile();
                    JobOrderWindow.ClickOnButton("Cancel");
                    Assert.IsTrue(status, "Job Order not copied successfully");
                }
            }

        }

        [TestMethod]
        public void CopyJobOrderAdditionalCharges()
        {
            var data = EllisHome.Initialize(ExcelFileNames.JobOrder);
            //open Job Order
            //OpenJobOrder.SelectJobOrderFromTable();

            foreach (var dataRow in data)
            {
                if (dataRow.ItemArray[1].ToString() == "CopyJobOrder" && dataRow.ItemArray[2].ToString() != "")
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());

                    //Copy Job Order Details from opened job order
                    var status = CopyJobOrder.CopyJobOrderAdditionalCharges();

                    if (status)
                        OpenJobOrder.CloseJobOrderProfile();
                    JobOrderWindow.ClickOnButton("Cancel");
                    Assert.IsTrue(status, "Job Order not copied successfully");
                }
            }

            Playback.Wait(2000);
            Assert.IsTrue(JobOrderWindow.VerifyCreateJobOrderWindowDisplayed(), "Job Order not copied successfully");
        }

        [TestMethod]
        public void CopyAndCreateJobOrder()
        {

            //Create job order from a copied details
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "CopyJobOrder" && dataRow.ItemArray[2].ToString() != "")
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #",
                        dataRow.ItemArray[2].ToString());

                    //Copy Job Order Details from opened job order
                    var status = CopyJobOrder.CopyJobOrderDetails();

                    if (status)
                    {
                        //Console.WriteLine(dataRow.ItemArray[1]);
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

                        status = PreQualifyingQuestionsWindow.HandleAlertWindow();
                        if (status)
                        {
                            Assert.IsFalse(status, "Job Order alredy exist for this customer");
                        }
                        else
                        {
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
                            status = PreQualifyingQuestionsWindow.HandleAlertWindow();


                            Assert.IsTrue(status, "Job order not saved successfully");
                        }

                    }


                }
            }
        }

        [TestMethod]
        public void CancelExistingJobOrder()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.JobOrder);

            foreach (var data in dataRows)
            {
                if (data.ItemArray[1].ToString() == "Cancel Job Order" && data.ItemArray[2].ToString() != "")
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", data.ItemArray[2].ToString());
                    var joprofile = OpenJobOrder.JobOrderProfileWindowProperties();
                    if (joprofile.Exists)
                    {
                        MouseActions.ClickButton(joprofile, "btnCancelJobOrder");
                        CancelJobOrder.CancelNewJobOrder();
                        CancelJobOrder.EnterJobOrderNotes(data.ItemArray[77].ToString(), data.ItemArray[78].ToString());
                        var cancelStatus = CancelJobOrder.HandleAlertWindow();
                        Assert.IsTrue(cancelStatus, "Job Order not canceled");

                        //Closing the newly created job order window
                        JobOrderWindow.CloseJobOrderProfileWindow();
                    }
                    Assert.IsTrue(joprofile.Exists, "No Job Order Profile found with given data");
                }
            }
        }

        [TestMethod]
        public void CancelNewJobOrder()
        {
            var runStatus = string.Empty;
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows)
            {
                //Data in "CancelJobOrderNotes" field is mandetory in TestData
                if (dataRow.ItemArray[77].ToString() != "" && dataRow.ItemArray[78].ToString() != "" && dataRow.ItemArray[1].ToString() == "CreateJobOrder")
                {
                    var jobOrderCreated = JobOrderWindow.CreateNewJobOrder(dataRow);
                    Assert.IsTrue(jobOrderCreated, "Job order not saved successfully");

                    //Cancel newly created job order

                    LandingPage.SelectFromToolbar("Job Orders");
                    TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", Globals.JobOrderNo);
                    var joprofile = OpenJobOrder.JobOrderProfileWindowProperties();
                    if (joprofile.Exists)
                    {
                        MouseActions.ClickButton(joprofile, "btnCancelJobOrder");
                        CancelJobOrder.CancelNewJobOrder();
                        CancelJobOrder.EnterJobOrderNotes(dataRow.ItemArray[77].ToString(), dataRow.ItemArray[78].ToString());
                        var cancelStatus = CancelJobOrder.HandleAlertWindow();
                        Assert.IsTrue(cancelStatus, "Job Order not canceled");

                        //Closing the newly created job order window
                        JobOrderWindow.CloseJobOrderProfileWindow();
                    }


                }
            }

        }

        [TestMethod]
        public void OpenExpiredJobOrder()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.JobOrderEdit);

            foreach (var dataRow in dataRows)
            {
                if (dataRow.ItemArray[1].ToString() == "ExpiredJobOrder" && dataRow.ItemArray[2].ToString() != "")
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    var expJoProfile = false;
                    var status = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());

                    if (status)
                    {
                        //---------------------------------------------------------------------------------------------
                        //JoExpiration   
                        //---------------------------------------------------------------------------------------------
                        var lblControl = Actions.GetWindowChild(OpenJobOrder.JobOrderProfileWindowProperties(), OpenJobOrder.JobOrderSummaryConstents.JoExpiration);
                        Console.WriteLine("Job Order Expiry Date: " + Convert.ToDateTime(lblControl.GetProperty("Name").ToString()));
                        if (Convert.ToDateTime(lblControl.GetProperty("Name").ToString()) < DateTime.Now)
                        {
                            Console.WriteLine("Job Order Expired On: " + Convert.ToDateTime(lblControl.GetProperty("Name").ToString()));
                            expJoProfile = true;
                        }


                        OpenJobOrder.CloseJobOrderProfile();
                    }
                    Assert.IsTrue(expJoProfile, "Expired Job ORder Profile not found");
                    Assert.IsTrue(status, "Profile not found");
                }

            }

        }

        [TestMethod]
        public void EditExpiredJobOrder()
        {
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderEdit);
            var status = OpenJobOrder.SelectExpiredJobOrderFromTable();
            if (status)
            {


                OpenJobOrder.CloseJobOrderProfile();
            }


            Assert.IsTrue(status, "Profile not found");
        }

        [TestMethod]
        public void VerifyJobOrderSummaryData()
        {
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderVerify);
            foreach (var data in datarows)
            {
                var profileStatus = OpenJobOrder.SelectJobOrderFromTable();
                var status = true;
                if (profileStatus)
                {
                    status = OpenJobOrder.VerifyJobOrder(data);
                    //Assert.IsTrue(status, "Profile data not matched");
                    OpenJobOrder.CloseJobOrderProfile();
                }
                Assert.IsTrue(status, "Profile data not matched");
                Assert.IsTrue(profileStatus, "Profile not found");


            }
        }

        [TestMethod]
        public void EditJobOrder()
        {
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderEdit);
            foreach (var data in datarows)
            {
                var profileStatus = OpenJobOrder.SelectJobOrderFromTable();
                if (profileStatus)
                {
                    OpenJobOrder.SelectTab("Basic Job Info");
                    Playback.Wait(2000);
                    OpenJobOrder.EditBasicJobInfoOfJobOrder(data);

                    OpenJobOrder.SelectTab("OrderDetails/Addl Charges");
                    Playback.Wait(2000);
                    OpenJobOrder.EditOrderDetailsAddlChargesOfJobOrder(data);

                    OpenJobOrder.SelectTab("Requirements");
                    Playback.Wait(2000);
                    OpenJobOrder.EditRequirementsOfJobOrder(data);

                    OpenJobOrder.SelectTab("Pre-Qualifying Questions");
                    Playback.Wait(2000);
                    OpenJobOrder.EditPreQualifyingQuestionsOfJobOrder(data);

                    OpenJobOrder.SelectTab("Safety");
                    Playback.Wait(2000);
                    OpenJobOrder.EditSafetyOfJobOrder(data);

                    OpenJobOrder.SelectTab("Progress Billing");
                    Playback.Wait(2000);
                    OpenJobOrder.EditProgressBillingOfJobOrder(data);

                    OpenJobOrder.CloseJobOrderProfile();

                }

                Assert.IsTrue(profileStatus, "Profile not found");

            }
        }

        [TestMethod]
        public void CopyFromCreateJobOrder()
        {
        }
    }
}