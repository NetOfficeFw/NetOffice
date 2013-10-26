using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// a quick and dirty tool without enterprise architecture(stable)
    /// </summary>
    static class Program
    {
        static Mutex _mutex = new Mutex(true, "{BB5FC79D-93EE-4e9b-A641-208F55689B1E}");

        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (_mutex.WaitOne(TimeSpan.Zero, true))
            { 
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FormMain());
                _mutex.ReleaseMutex();
            }
        }
    }
}
