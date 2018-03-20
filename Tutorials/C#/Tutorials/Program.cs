using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TutorialsCS4
{
    static class Program
    {
        internal static string DocumentationBase = "https://netoffice.io/documentation/";

        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMain());
        }
    }
}