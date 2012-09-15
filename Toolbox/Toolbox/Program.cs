using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Security.Principal;
using System.Diagnostics;

namespace NetOffice.DeveloperToolbox
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
                ProceedCommandLineArguments(args);
                PerformSelfElevation();
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(currentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                MainForm mainForm = new MainForm(args);
                Application.Run(mainForm);
         }

        static void currentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ErrorForm errorForm = new ErrorForm(null, ErrorCategory.Penalty, 0);
            errorForm.Show();
        }

        internal static bool IsAdmin
        {
            get 
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        internal static bool SelfElevation { get; set; }

        private static void PerformSelfElevation()
        {
            if (!IsAdmin && SelfElevation)
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Application.ExecutablePath;
                proc.Verb = "runas";

                try
                {
                    Process.Start(proc);
                }
                catch
                {
                    ; // The user refused the elevation. Do nothing and return directly ... (orininal MS)
                }
            }
        }

        private static void ProceedCommandLineArguments(string[] args)
        {
            foreach (string item in args)
            {
                if (item.Equals("-SelfElevation", StringComparison.InvariantCultureIgnoreCase))
                    SelfElevation = true;
            }
        }
    }
}
