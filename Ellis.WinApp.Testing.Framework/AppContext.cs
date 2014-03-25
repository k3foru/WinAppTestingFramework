//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Ellis.WinApp.Testing.Framework
{
    public class AppContext
    {
        protected static ApplicationUnderTest App;

        public TestContext TestContext { get; set; }

        protected static WinWindow EllisWindow;
    }
}