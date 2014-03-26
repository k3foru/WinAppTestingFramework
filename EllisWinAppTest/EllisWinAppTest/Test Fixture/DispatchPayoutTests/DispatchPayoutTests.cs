using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;

namespace EllisWinAppTest.DispatchPayoutTests
{
    [CodedUITest]
    public class DispatcPayoutTest : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
        }


    }
}
