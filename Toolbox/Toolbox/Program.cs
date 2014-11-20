using System;
using System.IO;
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
        public static void Main(string[] args)
        {
            try
            {
                ProceedCommandLineArguments(args);
                if (PerformSelfElevation())
                    return;
                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Forms.MainForm mainForm = new Forms.MainForm(args);
                Application.Run(mainForm);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception, ErrorCategory.Penalty, 1033);
            }
         }

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                string assemblyName = args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll";
                string assemblyFullPath = string.Empty;
                switch (assemblyName)
                {
                    case "ICSharpCode.SharpZipLib.dll":
                    case "Mono.Cecil.dll":
                    case "NetOffice.OutlookSecurity.dll":
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
                    {
                        assemblyFullPath = Path.Combine(Program.SubFolder, assemblyName);
                        if (File.Exists(assemblyFullPath))
                            return System.Reflection.Assembly.LoadFile(assemblyFullPath);
                        else
                            throw new FileNotFoundException(String.Format("Failed to load {0}", assemblyName));
                    }
                    default:
                        break;
                }

                return null;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception,ErrorCategory.Penalty, 1033);
                return null;
            }
        }

        public static string SubFolder
        {
            get
            {
                string resultPath = String.Empty;

                #if DEBUG
                
                    resultPath = Path.Combine(GetRelativeDebugPath(), "Libs");
                
                #else
                    
                    resultPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Toolbox Binaries");
                
                #endif

                if (!Directory.Exists(resultPath))
                    throw new DirectoryNotFoundException(resultPath);

                return resultPath;
            }
        }

        internal static string GetRelativeDebugPath()
        {
            string result = String.Empty;
            string[] array = Application.StartupPath.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < array.Length-3; i++)
                result += array[i] + "\\";
            return result;
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
                bool result = principal.IsInRole(WindowsBuiltInRole.Administrator);
                identity.Dispose();
                return result;
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
                    ; // The user refused the elevation. Do nothing and return directly ... (original MS comment)
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
            Forms.ErrorForm.ShowError(null, e.ExceptionObject as Exception,ErrorCategory.Penalty, 1033);
        }
    }
}
