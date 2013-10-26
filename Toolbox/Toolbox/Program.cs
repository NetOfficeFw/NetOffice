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
                if (PerformSelfElevation())
                    return;
                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                MainForm mainForm = new MainForm(args);
                Application.Run(mainForm);
         }

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string assemblyName = args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll";
            string assemblyFullPath = string.Empty;
            switch (assemblyName)
            {
                case "ICSharpCode.SharpZipLib.dll":
                case "Mono.Cecil.dll":
                case "NetOffice.OutlookSecurity.dll":
                    assemblyFullPath = string.Format("{0}\\Bin\\{1}", Application.StartupPath, assemblyName);
                    return System.Reflection.Assembly.LoadFile(assemblyFullPath);
                case "AccessApi.dll":
                case "ADODBApi.dll":
                case "DAOApi.dll":
                case "ExcelApi.dll":
                case "MSComctlLibApi.dll":
                case "MSDATASRCApi.dll":
                case "MSHTMLApi.dll":
                case "MSProjectApi.dll":
                case "NetOffice.dll":
                case "OfficeApi.dll":
                case "OutlookApi.dll":
                case "OWC10Api.dll":
                case "PowerPointApi.dll":
                case "VBIDEApi.dll":
                case "VisioApi.dll":
                case "WordApi.dll":
                    assemblyFullPath = string.Format("{0}\\Project Wizard\\NetOffice Assemblies\\4.0\\{1}", Application.StartupPath, assemblyName);
                    return System.Reflection.Assembly.LoadFile(assemblyFullPath);
                default:
                    break;
            }
            
            return null;
        }

        /// <summary>
        /// Analyze commandline arguments
        /// </summary>
        /// <param name="args"></param>
        private static void ProceedCommandLineArguments(string[] args)
        {
            foreach (string item in args)
            {
                if (item.Equals("-SelfElevation", StringComparison.InvariantCultureIgnoreCase))
                    SelfElevation = true;
            }
        } 

        /// <summary>
        /// Returns the program has admin privilegs
        /// </summary>
        internal static bool IsAdmin
        {
            get 
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        /// <summary>
        /// Hold the info to perform self elevation at start if necessary
        /// </summary>
        internal static bool SelfElevation { get; set; }

        /// <summary>
        /// Perform self elevation if necessary and wanted
        /// </summary>
        /// <returns>new process started</returns>
        private static bool PerformSelfElevation()
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
                    return true;
                }
                catch
                {
                    ; // The user refused the elevation. Do nothing and return directly ... (orininal MS)
                }
            }
            return false;
        }
         
        /// <summary>
        /// display unhandled exceptions
        /// </summary>
        /// <param name="sender">source</param>
        /// <param name="e">args</param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ErrorForm errorForm = new ErrorForm(null, ErrorCategory.Penalty, 0);
            errorForm.Show();
        }
    }
}
