using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.EllisHomeTests
{
    [CodedUITest]
    public class LandingPageTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        [TestMethod]
        public void VerifyJobOrderSchedulesDisplayed()
        {
            Initialize();

            LandingPage.SelectFromToolbar("JobOrder");
            Assert.IsTrue(LandingPage.VerifyJobOrderSchedulesDisplayed(), "Job Order not displayed");
            Cleanup();
        }

        [TestMethod]
        public void DispatchMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Dispatch");
            Assert.IsTrue(LandingPage.VerifyDispatchDisplayed(), "Dispatch Menu Items not displayed");
            Cleanup();
        }

        [TestMethod]
        public void WorkersMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Assert.IsTrue(LandingPage.VerifyWorkersDisplayed(), "Active Workers not displayed");
            Cleanup();
        }

        [TestMethod]
        public void CustomersMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Customers");
            Assert.IsTrue(LandingPage.VerifyCustomersDisplayed(), "Customer Quotes not displayed");
            Cleanup();
        }

        [TestMethod]
        public void ARMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("AR");
            Assert.IsTrue(LandingPage.VerifyArDisplayed(), "Customer Collections not displayed");
            Cleanup();
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}