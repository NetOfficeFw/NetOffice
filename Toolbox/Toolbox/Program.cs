using System;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using System.Security.Principal;
using System.Diagnostics;
using System.Threading;
using NetOffice.DeveloperToolbox.Utils.Native;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Well known assembly loader class
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// cache field to check program has admin privileges only at once
        /// </summary>
        private static bool? _isAdmin = null;

        /// <summary>
        /// An error occured in the AssemblyResolve trigger. We dont show the error dialog again in Main(string[] args) in this case
        /// </summary>
        private static bool _isShutDown = false;

        /// <summary>
        /// Used as systemwide singleton to create a single-application-instance. Works different in Debug/Release build
        /// </summary>
        private static Mutex _systemSingleton = null;
        
        /// <summary>
        /// Set in PerformSingleInstanceValidation. Its mean we are the origin owner of the mutex and we have to free them
        /// </summary>
        private static bool _mutexOwner = false;

        /// <summary>
        /// Assemblies we know from the dependencies sub folder. We load them at hand in AppDomain AssemblyResolve trigger
        /// </summary>
        private static string[] _dependencies = new string[] { "ICSharpCode.SharpZipLib.dll", "Mono.Cecil.dll", "NetOffice.OutlookSecurity.dll",
                                                                "AccessApi.dll", "ADODBApi.dll", "DAOApi.dll",
                                                                "ExcelApi.dll", "MSComctlLibApi.dll", "MSDATASRCApi.dll",
                                                                "MSHTMLApi.dll", "MSProjectApi.dll", "NetOffice.dll",
                                                                "OfficeApi.dll", "OutlookApi.dll", "OWC10Api.dll",
                                                                "PowerPointApi.dll", "VBIDEApi.dll", "VisioApi.dll",
                                                                "WordApi.dll", "MSFormsApi.dll", "PublisherApi.dll" };


        /// <summary>
        /// The main entry point for the component-based application. No need for a service architecture here so far. May this want be changed to CAB in the future
        /// </summary>
        /// <param name="args">application arguments</param>
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                StartTime = DateTime.Now;
                CreateMutex();
                ProceedCommandLineElevationArguments(args);
                if (PerformSingleInstanceValidation() || PerformSelfElevation())
                    return;

                // Nice to know: Its more safe to trigger the AssemblyResolve event in Main(string[] args) only and move all other code to a Main2 method (call Main2 at last in Main)
                // because the runtime try to bind target/used assemblies(when jump into main) before the AssemblyResolve trigger is established.
                // But we dont use ouer custom-bind assemblies in this Main(string[] args) so everything is okay.
                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Forms.MainForm mainForm = new Forms.MainForm(args);
                LoadedTime = DateTime.Now - StartTime;
                Console.WriteLine("DeveloperToolbox loaded in {0} seconds", LoadedTime.TotalSeconds);

                Application.Run(mainForm);
            }
            catch (Exception exception)
            {
                if (!_isShutDown)
                    Forms.ErrorForm.ShowError(null, exception, ErrorCategory.Penalty);
            }
            finally
            {
                ReleaseMutex();
            }
        }

        /// <summary>
        /// Whats the time we're started
        /// </summary>
        internal static DateTime StartTime { get; private set; }

        /// <summary>
        /// How long we need to be loaded without showing user interface
        /// </summary>
        internal static TimeSpan LoadedTime { get; private set; }

        /// <summary>
        /// The current used folder for dependent assemblies. Its the application subfolder 'Toolbox Binaries' in Release-Build and a custom folder in Debug-Build
        /// </summary>
        public static string DependencySubFolder
        {
            get
            {
                string resultPath = String.Empty;

                #if DEBUG
                             
                    resultPath = Path.Combine(GetInternalRelativeDebugPath(), "Assemblies\\Any CPU");
                
                #else
                                        
                    resultPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Toolbox Binaries");
                
                #endif

                return resultPath;
            }
        }

        /// <summary>
        /// The current used folder for dependent assemblies when application is given in release package
        /// </summary>
        public static string DependencyReleaseSubFolder
        {
            get
            {
                string resultPath = Path.Combine(System.Windows.Forms.Application.StartupPath, @".NET 4\Assemblies\Any CPU");
                return resultPath;
            }
        }

        /// <summary>
        /// Current/Highest NetOffice public or preview release version
        /// </summary>
        public static string CurrentNetOfficeVersion
        {
            get 
            {
                return "1.7.4.4";
            }
        }

        /// <summary>
        /// Returns info the program has admin privilegs (Cache supported, not thread-safe)
        /// </summary>
        internal static bool IsAdmin
        {
            get 
            {
                if (null == _isAdmin)
                {
                    using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                    {
                        WindowsPrincipal principal = new WindowsPrincipal(identity);
                        bool result = principal.IsInRole(WindowsBuiltInRole.Administrator);
                        _isAdmin = result;
                    }
                }
                return (bool)_isAdmin;
            }
        }

        /// <summary>
        /// Returns info the assembly is currently in design mode. In other words its running in ouer IDE at design-time
        /// </summary>
        internal static bool IsDesign
        {
            get
            {
                return (System.ComponentModel.LicenseManager.UsageMode == System.ComponentModel.LicenseUsageMode.Designtime);
            }
        }

        /// <summary>
        /// Hold the info to perform self elevation at start if necessary
        /// </summary>
        internal static bool SelfElevation { get; set; }

        /// <summary>
        /// Perform self elevation if necessary and wanted
        /// </summary>
        /// <param name="forceElevation">force elevation even Program.SelfElevation is false</param>
        /// <returns>true if new process is sucsessfuly started, otherwise false</returns>
        internal static bool PerformSelfElevation(bool forceElevation = false)
        {
            if (!IsAdmin && (SelfElevation || forceElevation))
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Application.ExecutablePath;
                proc.Verb = "runas";

                try
                {
                    ReleaseMutex();
                    Process.Start(proc);
                    return true;
                }
                catch
                {
                    ; // The user refused the failed elevation. Do nothing and return directly ... (original MS guidance)
                }
            }
            return false;
        }

        /// <summary>
        /// Find the local root folder in debug mode. The method use the Application.Startup path and returns the folder 4x upward.
        /// </summary>
        /// <returns>The current related debug root folder</returns>
        private static string GetInternalRelativeDebugPath()
        {
            string result = String.Empty;
            string[] array = Application.StartupPath.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < array.Length - 4; i++)
                result += array[i] + "\\";
            return result;
        }

        /// <summary>
        /// Creates the systemwide singleton mutex for the single-application pattern
        /// </summary>
        private static void CreateMutex()
        {             
            #if DEBUG
                
                _systemSingleton = new Mutex(true, Guid.NewGuid().ToString());
                
            #else

                // The 0FF1CE idea in the GUID is from the MS-PowerPoint Product Code (i stoled from the pp devs)
                _systemSingleton = new Mutex(true, "D3413BEF-46D9-4F96-82FC-0000000FF1CE");
            
            #endif
        }

        /// <summary>
        /// Release the systemwide singleton mutex if we are the owner
        /// </summary>
        private static void ReleaseMutex()
        {
            if (null != _systemSingleton && _mutexOwner)
                _systemSingleton.ReleaseMutex();
            _systemSingleton = null;
        }

        /// <summary>
        /// Analyze commandline arguments for self elevation and set SelfElevation property
        /// </summary>
        /// <param name="args">arguments from command line</param>
        private static void ProceedCommandLineElevationArguments(string[] args)
        {
            if (null == args)
                return;           
            SelfElevation = args.Any(e => e.Equals("-SelfElevation", StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// We want to detect an instance of the application is already running.
        /// If its true we want post a (custom) message to the main window of these instance that means "bring you in front"
        /// </summary>
        /// <returns>true if a previous instance is running, otherwise false</returns>
        private static bool PerformSingleInstanceValidation()
        {            
            #if DEBUG
            
                // we want allow multiple instances in debug build
                // (its also easier to use because sometimes the mutex still lives on if debugging is aborted at hand) 
                _mutexOwner = true;
                return false;
            
            #else

                if (!_systemSingleton.WaitOne(TimeSpan.Zero))
                {
                    _mutexOwner = false;
                    // I dislike "on the fly-casts" but its okay for constant values(which it is) what i find
                    Win32.PostMessage((IntPtr)Win32.HWND_BROADCAST, Win32.WM_SHOWTOOLBOX, IntPtr.Zero, IntPtr.Zero);
                    return true;
                }
                else
                {
                    _mutexOwner = true;
                    return false;
                }
            
            #endif
        }

        /// <summary>
        /// Try to load an assembly with given file path
        /// </summary>
        /// <param name="assemblyFullPath">full qualified assembly path</param>
        /// <returns>loaded assembly instance</returns>
        private static Assembly LoadFile(string assemblyFullPath)
        {
            if (String.IsNullOrWhiteSpace(assemblyFullPath))
                throw new ArgumentNullException("assemblyFullPath");

            try
            {
                // we check its from well known dependencies folder and one of the registererd dependencies
                // because we dont want load any injected code from an attacker scenario
                // OPEN-TODO-1: Add file version/hash and signed assembly check to improve security

                string assemblyFolderPath = Path.GetDirectoryName(assemblyFullPath);
                string assemblyFileName = Path.GetFileName(assemblyFullPath);
                
                if (!DependencySubFolder.Equals(assemblyFolderPath, StringComparison.InvariantCultureIgnoreCase) &&
                    !DependencyReleaseSubFolder.Equals(assemblyFolderPath, StringComparison.InvariantCultureIgnoreCase))
                    throw new System.Security.SecurityException("Invalid assembly directory.");

                if (!_dependencies.Contains(assemblyFileName))
                    throw new System.Security.SecurityException("Invalid assembly file.");

                // UnsafeLoadFrom allows to load assemblies from may unsafe locations. A lot of issue reports before so i switch to this one
                return Assembly.UnsafeLoadFrom(assemblyFullPath);
            }
            catch (Exception exception)
            {
                throw new FileLoadException(String.Format("Failed to load {0}", assemblyFullPath), exception);
            }
        }

        /// <summary>
        /// display unhandled exception(s) with non-modal ErrorForm instance
        /// </summary>
        /// <param name="sender">source(ignored)</param>
        /// <param name="e">exception detailed informations</param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // if its in shutdown(because a heavy error occured) we dont want to show another error again
            if (_isShutDown)
                return;

            try
            {
                Forms.ErrorForm.ShowError(null, e.ExceptionObject as Exception, ErrorCategory.Penalty);
            }
            catch (Exception exception)
            {
                // no idea whats the problem right now(may no message loop) but log the error to further investigation
                Console.WriteLine("CurrentDomain_UnhandledException:{0}=>{1}", exception, e.ExceptionObject as Exception);   
            }
        }

        /// <summary>
        /// We handle missing dependencies at hand because we want this .exe assembly in a clean directory. This looks more nicely for the user
        /// </summary>
        /// <param name="sender">unkown sender(ignored)</param>
        /// <param name="args">arguments with info what we are looking for</param>
        /// <returns>resolved assembly or null</returns>
        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // if its in shutdown we dont need another assembly anymore
            if (_isShutDown)
                return null;

            try
            {
                // detect its a assembly reference or a file path and extract the assembly file name
                string assemblyName = null;
                if (args.Name.IndexOf(",", StringComparison.InvariantCultureIgnoreCase) > -1)
                    assemblyName = args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll";
                else
                    assemblyName = Path.GetFileName(args.Name);

                if (_dependencies.Contains(assemblyName))
                {
                    string assemblyFullPath = Path.Combine(Program.DependencySubFolder, assemblyName);
                    if (File.Exists(assemblyFullPath))
                        return LoadFile(assemblyFullPath);
                    else
                    {
                        assemblyFullPath = Path.Combine(Program.DependencyReleaseSubFolder, assemblyName);
                        if (File.Exists(assemblyFullPath))
                            return LoadFile(assemblyFullPath);
                        else
                            throw new FileNotFoundException(String.Format("Failed to load {0}", assemblyName));
                    }                        
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception, ErrorCategory.Penalty, "Unable to load a dependency.");
                _isShutDown = true;
                Application.Exit();
            }

            return null;
        }
    }
}