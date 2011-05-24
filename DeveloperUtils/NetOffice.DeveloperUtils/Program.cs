using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length > 0)
            {
                SupportByLibrary.SupportByLibraryControl control = new NetOffice.DeveloperUtils.SupportByLibrary.SupportByLibraryControl(args);
            }
            else
            { 
                Application.Run(new Form1(args));
            }
        }


    }
}
