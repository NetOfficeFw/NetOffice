using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace NetOffice.Loader
{
    /// <summary>
    /// Encapsulate current appdomain with loader services and exception tolerant methods
    /// </summary>
    internal class CurrentAppDomain
    {
        #region Fields

        private Version _assemblyVersion;
        private static readonly string[] _assemblyNames = new string[] {
                                                                        "OfficeApi.dll", "ExcelApi.dll", "WordApi.dll",
                                                                        "OutlookApi.dll", "PowerPointApi.dll", "AccessApi.dll",
                                                                        "VisioApi.dll", "MSProjectApi.dll", "PublisherApi.dll",
                                                                        "VBIDEApi.dll", "MSFormsApi.dll"
                                                                       };
        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner core</param>
        internal CurrentAppDomain(Core owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            Owner = owner;
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Owner Core
        /// </summary>
        internal Core Owner { get; private set; }

        /// <summary>
        /// Core should load these assemblies while initialize if files exists in current codebase 
        /// </summary>
        internal string[] AssemblyNames
        {
            get
            {
                return _assemblyNames;
            }
        }

        /// <summary>
        /// Current Assembly Version
        /// </summary>
        private Version AssemblyVersion
        {
            get
            {
                if (null == _assemblyVersion)
                    _assemblyVersion = Owner.ThisAssembly.GetName().Version;
                return _assemblyVersion;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Returns all loaded assemblies in current appdomain
        /// </summary>
        /// <returns>loaded assemblies</returns>
        internal Assembly[] GetAssemblies()
        {
            try
            {
                return AppDomain.CurrentDomain.GetAssemblies();
            }
            catch (AppDomainUnloadedException)
            {
                return new Assembly[0];
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Try load an assembly
        /// </summary>
        /// <param name="fileName">full qualified file path</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly Load(string fileName)
        {
            try
            {
                if (ValidateVersion(fileName))
                {
                    if (Owner.Settings.LoadAssembliesUnsafe)
                        return Assembly.UnsafeLoadFrom(fileName);
                    else
                        return Assembly.LoadFrom(fileName);
                        // changed Load to LoadFrom, thanks to Frank Fajardo
                }
                else
                    return null;
            }
            catch(Exception exception)
            {
                Owner.Console.WriteLine("AssemblyLoad Exception<{1}> {0}", Path.GetFileName(fileName), exception.GetType().Name);
                return null;
            }
        }

        /// <summary>
        /// Try load an assembly
        /// </summary>
        /// <param name="name">assembly reference name</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly Load(AssemblyName name)
        {
            string localPath = String.Empty;
            try
            {
                bool versionMatch = false;
                localPath = false == String.IsNullOrEmpty(name.CodeBase) ? Resolver.UriResolver.ResolveLocalPath(name.CodeBase) : null;
                versionMatch = null == localPath ? ValidateVersion(name) : ValidateVersion(localPath);
                if (null == localPath)
                {
                    string thisLocalPath = Resolver.UriResolver.ResolveLocalPath(Owner.ThisAssembly.CodeBase);
                    string extension = Path.GetExtension(thisLocalPath);
                    string path = Path.GetDirectoryName(thisLocalPath);
                    localPath = Path.Combine(path, name.Name + extension);
                }

                if (!File.Exists(localPath))
                {
                    Owner.Console.WriteLine("AssemblyLoad {0} Unable To Find Assembly In Local Directory.", name.Name);
                    return null;
                }

                if (versionMatch)
                {
                    if (Owner.Settings.LoadAssembliesUnsafe)
                    {
                        return Assembly.UnsafeLoadFrom(localPath);
                    }
                    else
                    {
                        return Assembly.LoadFrom(localPath);
                    }
                }
                else
                    return null;
            }
            catch(Exception exception)
            {
                if (localPath != String.Empty && false == String.IsNullOrEmpty(name.CodeBase))
                    Owner.Console.WriteLine("AssemblyLoad(From Path) Exception<{1}> {0}", Path.GetFileName(localPath), exception.GetType().Name);
                else
                    Owner.Console.WriteLine("AssemblyLoad(From Name) Exception<{1}> {0}", name, exception.GetType().Name);
                return null;
            }
        }

        /// <summary>
        /// Try to validate argument file version match with current NetOffice version.
        /// The method does nothing if argument file not exists.
        /// </summary>
        /// <param name="fileName">target file to load</param>.resources
        /// <returns>true if file exists in current NetOffice version</returns>
        internal bool ValidateVersion(string fileName)
        {
            if (File.Exists(fileName))
            {
                FileVersionInfo info = FileVersionInfo.GetVersionInfo(fileName);
                if (info.FileVersion == AssemblyVersion.ToString())
                    return true;
                else
                    Owner.Console.WriteLine("Invalid Assembly Version {0}", info.FileVersion);
            }
            return false;
        }

        /// <summary>
        /// Try to validate argument file version match with current NetOffice version. 
        /// </summary>
        /// <param name="name">given assembly specification</param>
        /// <returns>true if file exists in current NetOffice version</returns>
        internal bool ValidateVersion(AssemblyName name)
        {
            if (name.Version == AssemblyVersion)
            { 
                return true;
            }
            else
            { 
                Owner.Console.WriteLine("Assembly Version {0} missmatch.", name.Version);
                return false;
            }
        }

        /// <summary>
        /// Try load known assembly names
        /// </summary>
        /// <param name="factory">core to use</param>
        internal void TryLoadAssemblies(Core factory)
        {
            foreach (string item in AssemblyNames)
            {
                string fileName = PathBuilder.BuildLocalPathFromAssemblyFileName(factory, item);               
                Assembly assembly = Load(fileName);
                Owner.Console.WriteLine("TryLoad {0} {1}", fileName, null != assembly ? "#ok" : "#fail");
            }
        }

        /// <summary>
        /// Returns embedded keytoken schema
        /// </summary>
        /// <param name="factory">factory type to use</param>
        /// <returns>keytoken line array</returns>
        internal static string[] KeyTokens(Core factory)
        {
            using (System.IO.Stream ressourceStream = factory.ThisAssembly.GetManifestResourceStream(factory.ThisType.Namespace + ".KeyTokens.txt"))
            {
                using (System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream))
                {
                    string text = textStreamReader.ReadToEnd();
                    return text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                }
            }
        }

        #endregion

        #region Trigger

        /// <summary>
        /// Occurs when the resolution of an assembly fails.
        /// </summary>
        /// <param name="sender">The source of the event</param>
        /// <param name="args">A System.ResolveEventArgs that contains the event data</param>
        /// <returns>The System.Reflection.Assembly that resolves the type, assembly, or resource or null if the assembly cannot be resolved</returns>
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                // dont care for resources
                if ((!String.IsNullOrEmpty(args.Name) && args.Name.ToLower().Trim().IndexOf(".resources") > -1))
                    return null;

                string thisLocalPath = Resolver.UriResolver.ResolveLocalPath(Owner.ThisAssembly.CodeBase);
                string extension = Path.GetExtension(thisLocalPath);
                string path = Path.GetDirectoryName(thisLocalPath);
                string fullFileName = Path.Combine(path, args.Name + extension);

                if (!System.IO.File.Exists(fullFileName))
                {
                    // given argument is possibly an assembly string like:
                    // System.Security, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
                    string[] chars = args.Name.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    if (chars != null && chars.Length > 0)
                        fullFileName = Path.Combine(path, chars[0] + extension);
                }

                if (System.IO.File.Exists(fullFileName))
                {
                    Owner.Console.WriteLine(string.Format("Try to resolve assembly {0}", args.Name));
                    Assembly assembly = Load(fullFileName);
                    return assembly;
                }
                else
                { 
                    Owner.Console.WriteLine(string.Format("Unable to resolve assembly {0}. The file doesnt exists in current codebase.", args.Name));
                    return null;
                }
            }
            catch (Exception exception)
            {
                Owner.Console.WriteException(exception);
                return null;
            }
        }

        #endregion
    }
}