//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.Diagnostics;
using System.Linq;

namespace Ellis.WinApp.Testing.Framework.Actions
{

    public class WindowsActions : AppContext
    {
        public static void KillEllisProcesses()
        {
            try
            {
                var runingProcess = Process.GetProcesses();

                foreach (var t in runingProcess.Where(t => t.ProcessName == "Ellis"))
                    t.Kill();
            }
            catch
            {
                //Suppress any exception here
            }
        }

        public static void KillProcesses(string processName)
        {
            try
            {
                var runingProcess = Process.GetProcesses();

                foreach (var t in runingProcess.Where(t => t.ProcessName == processName))
                    t.Kill();
            }
            catch
            {
                //Suppress any exception here
            }
        }
    }
}
