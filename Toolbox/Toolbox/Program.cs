using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Security.Principal;
using System.Diagnostics;

namespace NetOffice.DeveloperToolbox
{
    internal static class Program
    {
        /// <summary>
        /// cache field to check program has admin privileges only at once
        /// </summary>
        private static bool? _isAdmin;

        /// <summary>
        /// The main entry point for the component-based application. No need for a service architecture here so far. May this want be changed to CAB in the future.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                StartTime = DateTime.Now;
                ProceedCommandLineElevationArguments(args);
                if (PerformSelfElevation())
                    return;

                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Forms.MainForm mainForm = new Forms.MainForm(args);
                LoadedTime = DateTime.Now - StartTime;
                Console.WriteLine("Loaded in {0} seconds", LoadedTime.TotalSeconds);

                Application.Run(mainForm);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception, ErrorCategory.Penalty);
            }
        }

        /// <summary>
        /// Whats the time we're started
        /// </summary>
        internal static DateTime StartTime { get; private set; }

        /// <summary>
        /// How long we need to be loaded without show user interface
        /// </summary>
        internal static TimeSpan LoadedTime { get; private set; }

        /// <summary>
        /// The current used folder for dependent assemblies
        /// </summary>
        public static string DependencySubFolder
        {
            get
            {
                string resultPath = String.Empty;

                #if DEBUG

                    resultPath = Path.Combine(GetInternalRelativeDebugPath(), "Libs");
                
                #else
                    
                    resultPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Toolbox Binaries");
                
                #endif

                if (!Directory.Exists(resultPath))
                    throw new DirectoryNotFoundException(resultPath);

                return resultPath;
            }
        }

        /// <summary>
        /// Current NO public release version
        /// </summary>
        public static string CurrentNetOfficeVersion
        {
            get 
            {
                return "1.7.2.0";
            }
        }

        /// <summary>
        /// Find the local root folder in debug mode. The method use the Application.Startup method path and returns the folder 3x upward.
        /// </summary>
        /// <returns>The current related debug root folder</returns>
        private static string GetInternalRelativeDebugPath()
        {
            string result = String.Empty;
            string[] array = Application.StartupPath.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < array.Length - 3; i++)
                result += array[i] + "\\";
            return result;
        }

        /// <summary>
        /// Analyze commandline arguments for self elevation
        /// </summary>
        /// <param name="args"></param>
        private static void ProceedCommandLineElevationArguments(string[] args)
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
                if (null == _isAdmin)
                { 
                    WindowsIdentity identity = WindowsIdentity.GetCurrent();
                    WindowsPrincipal principal = new WindowsPrincipal(identity);                
                    bool result = principal.IsInRole(WindowsBuiltInRole.Administrator);
                    identity.Dispose();
                    _isAdmin = result;
                }
                return (bool)_isAdmin;
            }
        }

        /// <summary>
        /// Hold the info to perform self elevation at start if necessary
        /// </summary>
        internal static bool SelfElevation { get; set; }

        /// <summary>
        /// Perform self elevation if necessary and wanted
        /// </summary>
        /// <returns>new process started is sucsessfuly started</returns>
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
                    ; // The user refused the failed elevation. Do nothing and return directly ... (original MS comment)
                }
            }
            return false;
        }
         
        /// <summary>
        /// display unhandled exception(s)
        /// </summary>
        /// <param name="sender">source(ignored)</param>
        /// <param name="e">args</param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Forms.ErrorForm.ShowError(null, e.ExceptionObject as Exception, ErrorCategory.Penalty);
        }

        /// <summary>
        /// We handle missing dependencies at hand because we want this .exe assembly in a clean directory. This looks more nicely for the user
        /// </summary>
        /// <param name="sender">unkown sender</param>
        /// <param name="args">arguments with info what we are looking for</param>
        /// <returns>Resolved assembly or null</returns>
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
                            assemblyFullPath = Path.Combine(Program.DependencySubFolder, assemblyName);
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
                Forms.ErrorForm.ShowError(null, exception, ErrorCategory.Penalty);
                return null;
            }
        }

    }
}
