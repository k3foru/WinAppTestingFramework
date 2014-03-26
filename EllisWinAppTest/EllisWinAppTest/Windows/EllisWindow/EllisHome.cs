using System.Collections.Generic;
using System.Data;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.EllisWindow
{
    public class EllisHome : AppContext
    {
        public static IEnumerable<DataRow> Initialize(string excelName)
        {
            WindowsActions.KillEllisProcesses();
            App = LaunchEllisAsCSRUser();
            var datarows = ExcelReader.ImportSpreadsheet(excelName);
            App.SetFocus();
            return datarows;
        }

        public static void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = LaunchEllisAsCSRUser();
            App.SetFocus();
        }

        public static ApplicationUnderTest LaunchEllis()
        {
            App = ApplicationUnderTest.Launch(TestData.Path, TestData.AltPath);
            Factory.GetWinProperties();

            return App;
        }


        public static ApplicationUnderTest LaunchEllisAsCSRUser()
        {
            //LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.CSRUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsARMUser()
        {
            //LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.ARMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsDMUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.DMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsARRUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.ARRUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsAVPUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.AVPUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsNABSUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.NABSUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsNAPSUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.NAPSUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsTAUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.TAUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsRAUser()
        {
            LaunchActions.LaunchAppAsDifferentUser(TestData.Path, TestData.RAUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsDiffUserFromDesktop()
        {
            LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            Playback.Wait(6000);
            App = Factory.GetAUT();
            Factory.GetWinProperties();
            return App;
        }

        //public static ApplicationUnderTest LaunchEllisWithCredentials()
        //{
        //    //var pass = TestData.AppPassword;
        //    //SecureString Result = new SecureString();
        //    //foreach (char c in pass.ToCharArray())
        //    //    Result.AppendChar(c);

        //    var s = new NetworkCredential("", TestData.AppPassword).SecurePassword;

        //    //var path = @"C:\Program Files (x86)\True Blue\Ellis\Ellis.exe";
        //    //ProcessStartInfo myProc = new ProcessStartInfo(path);

        //    //try
        //    //{                
        //    //    myProc.Domain = "CORP";
        //    //    myProc.UserName = TestData.AppUsername;
        //    //    myProc.Password = s;
        //    //    myProc.UseShellExecute = false;
        //    //    myProc.WorkingDirectory = Path.GetDirectoryName(path);

        //    //    //Process.Start(myProc);
        //    //}
        //    //catch (Exception myException)
        //    //{
        //    //    // error handling
        //    //}


        //    //App = Actions.LaunchAppAsDifferentUser(TestData.Path, TestData.AppUsername, TestData.AppPassword, TestData.AppDomain);
        //    App = ApplicationUnderTest.Launch(TestData.Path, TestData.AltPath, null, TestData.AppUsername, s, "CORP");
        //    Factory.GetWinProperties();

        //    return App;
        //}


        public static void ClickOnFileExit()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new {Name = EllisHomeConstants.File});
            var exit = file.Container.SearchFor<WinMenuItem>(new {Name = EllisHomeConstants.Exit});

            Mouse.Click(file);
            Playback.Wait(2000);
            Mouse.Click(exit);
        }
    }

    internal class EllisHomeConstants
    {
        public const string File = "File";
        public const string Exit = "Exit";
        public const string Worker = "Workers";
        public const string CApplicant = "Create Applicant";
        public const string JobOrder = "Job Orders";
        public const string CJobOrder = "Create Job Order";

        public const string Customer = "Customers";
        public const string CCustomer = "Create Customer";
    }
}