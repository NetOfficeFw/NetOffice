using System;
using System.Collections.Generic;
using System.Windows.Forms;

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
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(currentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm(args));
         }

        static void currentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ErrorForm errorForm = new ErrorForm(null, ErrorCategory.Penalty, 0);
            errorForm.Show();
        }
    }
}
