using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Encapsulate current appdomain with exception tolerant methods
    /// </summary>
    internal class CurrentAppDomain
    {
        #region Fields

        private Version _assemblyVersion;

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
            catch
            {
                return new Assembly[0];
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
                localPath = false == String.IsNullOrEmpty(name.CodeBase) ? UriConvert.ToLocalPath(name.CodeBase) : null;
                versionMatch = null == localPath ? ValidateVersion(name) : ValidateVersion(localPath);
                if (null == localPath)
                {
                    string thisLocalPath = UriConvert.ToLocalPath(Owner.ThisAssembly.CodeBase);
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
        /// Try load an assembly
        /// </summary>
        /// <param name="fileName">full qualified file name</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly LoadFrom(string fileName)
        {
            try
            {
                if (ValidateVersion(fileName))
                {
                    if (Owner.Settings.LoadAssembliesUnsafe)
                        return Assembly.UnsafeLoadFrom(fileName);
                    else
                    {
                        //  todo: find a #pragma to make a possible exception silent
                        // even the developer want a hard debugger break in all(or match) CLR exception(s)
                        return Assembly.LoadFrom(fileName);
                    }
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
        /// Try to validate argument file version match with current NO version. The method does nothing if argument file not exists
        /// </summary>
        /// <param name="fileName">target file to load</param>.resources
        /// <returns>true if file exists in current NO version</returns>
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
        /// Try to validate argument file version match with current NO version. 
        /// </summary>
        /// <param name="name">given assembly specification</param>
        /// <returns>true if file exists in current NO version</returns>
        internal bool ValidateVersion(AssemblyName name)
        {
            if (name.Version == AssemblyVersion)
            { 
                return true;
            }
            else
            { 
                Owner.Console.WriteLine("Negative Assembly Version {0}", name.Version);
                return false;
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

                string thisLocalPath = UriConvert.ToLocalPath(Owner.ThisAssembly.CodeBase);
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
                    Console.WriteLine(string.Format("Try to resolve assembly {0}", args.Name));
                    Assembly assembly = LoadFrom(args.Name);
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
