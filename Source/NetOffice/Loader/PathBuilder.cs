using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NetOffice.Loader
{
    internal static class PathBuilder
    {
        /// <summary>
        /// Build path from dependent assambly
        /// </summary>
        /// <param name="assembly">assembly to resolve</param>
        /// <returns>resolved path</returns>
        public static string BuildLocalPathFromDependentAssembly(DependentAssembly assembly)
        {
            string fileName = assembly.ParentAssembly.CodeBase.Substring(0, assembly.ParentAssembly.CodeBase.LastIndexOf("/")) + "/" + assembly.Name;
            fileName = fileName.Replace("/", "\\").Substring(8);
            return fileName;
        }

        /// <summary>
        /// Build path from assembly file name
        /// </summary>
        /// <param name="factory">factory to use codebase directory from</param>
        /// <param name="assemblyName">assembly name like abc.dll</param>
        /// <returns>resolved path</returns>
        public static string BuildLocalPathFromAssemblyFileName(Core factory, string assemblyName)
        {
            string localAssemblyPath = Resolver.UriResolver.ResolveLocalPath(factory.ThisAssembly.CodeBase);
            string directoryName = System.IO.Path.GetDirectoryName(localAssemblyPath);
            string fullFileName = System.IO.Path.Combine(directoryName, assemblyName);
            return fullFileName;
        }
    }
}
