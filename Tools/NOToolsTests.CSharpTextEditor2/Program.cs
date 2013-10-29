using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace NOToolsTests.CSharpTextEditor2
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormHost());
        }
    }
}
