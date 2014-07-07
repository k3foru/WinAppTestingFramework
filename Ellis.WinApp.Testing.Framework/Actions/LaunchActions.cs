//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using Microsoft.VisualStudio.TestTools.UITesting;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class LaunchActions : AppContext
    {
        public static ApplicationUnderTest LaunchAppAsDifferentUser(string appPath, string username, string password,
            string domain)
        {
            var s = new NetworkCredential("", password).SecurePassword;
            var path = @appPath;

            // Open App.Config of executable
            var configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = @path + ".config";
            var config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);

            // Add an Application Setting.
            config.AppSettings.Settings.Remove("IsBranch");
            config.AppSettings.Settings.Add("IsBranch", "true");

            // Save the configuration file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");
            config.Save(ConfigurationSaveMode.Modified, true);

            var myProc = new ProcessStartInfo(path);

            try
            {
                myProc.Domain = domain;
                myProc.UserName = username;
                myProc.Password = s;
                myProc.UseShellExecute = false;
                myProc.WorkingDirectory = Path.GetDirectoryName(path);
            }
            catch (Exception)
            {
                // error handling
            }

            App = ApplicationUnderTest.Launch(myProc);
            return App;
        }

        public static void LaunchAppAsDiffferentUserFromDesktop()
        {
            DiffUser.EllisDifferentUser();
        }
    }
}
