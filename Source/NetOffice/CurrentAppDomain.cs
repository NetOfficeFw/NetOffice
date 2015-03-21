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
        /// <param name="fileName">full qualified file path</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly Load(string fileName)
        {
            if (ValidateVersion(fileName))
                return Assembly.Load(fileName);
            else
                return null;
        }

        /// <summary>
        /// Try load an assembly
        /// </summary>
        /// <param name="name">assembly reference name</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly Load(AssemblyName name)
        {
            try
            {
                return Assembly.Load(name);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Try load an assembly
        /// </summary>
        /// <param name="fileName">full qualified name</param>
        /// <returns>Assembly instance or null</returns>
        internal Assembly LoadFile(string fileName)
        {
            try
            {
                if (ValidateVersion(fileName))
                    return Assembly.LoadFile(fileName);
                else
                    return null;
            }
            catch
            {
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
                    return Assembly.LoadFrom(fileName);
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Try to validate argument file version match with current NO version. The method does nothing if argument file not exists
        /// </summary>
        /// <param name="fileName">target file to load</param>
        /// <returns>true if file exists in current NO version</returns>
        internal bool ValidateVersion(string fileName)
        {
            if (File.Exists(fileName))
            {
                FileVersionInfo info = FileVersionInfo.GetVersionInfo(fileName);
                if (info.FileVersion == AssemblyVersion.ToString())
                    return true;
            }
            return false;
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

                string directoryName = Owner.ThisAssembly.CodeBase.Substring(0, Owner.ThisAssembly.CodeBase.LastIndexOf("/"));
                directoryName = directoryName.Replace("/", "\\").Substring(8);
                string fileName = args.Name.Substring(0, args.Name.IndexOf(","));
                string fullFileName = System.IO.Path.Combine(directoryName, fileName + ".dll");
                if (System.IO.File.Exists(fullFileName))
                {
                    Console.WriteLine(string.Format("Try to resolve assembly {0}", args.Name));
                    Assembly assembly = Load(args.Name);
                    return assembly;
                }
                else
                {
                    Console.WriteLine(string.Format("Unable to resolve assembly {0}. The file doesnt exists in current codebase.", args.Name));
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
