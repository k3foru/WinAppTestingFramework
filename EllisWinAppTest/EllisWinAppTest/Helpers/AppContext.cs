using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.Helpers
{
    public class AppContext
    {
        protected static ApplicationUnderTest App;

        public TestContext TestContext { get; set; }

        protected static WinWindow EllisWindow;
    }
}