using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NetOffice.Loader
{
    internal static class PathBuilder
    {
        /// <summary>
        /// Build path from dependent assembly
        /// </summary>
        /// <param name="assembly">assembly to resolve</param>
        /// <returns>resolved path</returns>
        public static string BuildLocalPathFromDependentAssembly(DependentAssembly assembly)
        {
            string parentAssemblyDirectory = Path.GetDirectoryName(assembly.ParentAssembly.Location);
            string localFilename = Path.Combine(parentAssemblyDirectory, assembly.Name);
            return localFilename;
        }

        /// <summary>
        /// Build path from assembly file name
        /// </summary>
        /// <param name="factory">factory to use codebase directory from</param>
        /// <param name="assemblyName">assembly name like abc.dll</param>
        /// <returns>resolved path</returns>
        public static string BuildLocalPathFromAssemblyFileName(Core factory, string assemblyName)
        {
            string localAssemblyPath = Resolver.UriResolver.ResolveLocalPath(factory.ThisAssembly.Location);
            string directoryName = System.IO.Path.GetDirectoryName(localAssemblyPath);
            string fullFileName = System.IO.Path.Combine(directoryName, assemblyName);
            return fullFileName;
        }
    }
}
