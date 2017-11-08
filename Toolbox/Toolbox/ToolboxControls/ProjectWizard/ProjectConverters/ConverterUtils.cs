using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    /// <summary>
    /// Registry Root Kind
    /// </summary>
    internal enum RegistryHive
    {
        /// <summary>
        /// The well known local machine key
        /// </summary>
        HKEY_Local_Machine = 0,
        
        /// <summary>
        /// The well known current user key
        /// </summary>
        HKEY_Current_User
    }

    /// <summary>
    /// Environment to String helper
    /// </summary>
    internal class EnvironmentVersions
    {
        /// <summary>
        /// Converty given arguments to solution file(.sln) entry
        /// </summary>
        /// <param name="environment">given ide environment</param>
        /// <param name="language">target language</param>
        /// <returns>valid sln string entry</returns>
        public string this[IDE environment, ProgrammingLanguage language]
        {
            get
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        if (language == ProgrammingLanguage.CSharp)
                            return "# Visual C# Express 2010";
                        else
                            return "# Visual Basic Express 2010";
                    //case IDE.VS2012:
                    //    if (language == ProgrammingLanguage.CSharp)
                    //        return "# Visual C# Express 2012";
                    //    else
                    //        return "# Visual Basic Express 2012";
                    case IDE.VS20131517:
                        return "# Visual Studio Express 2013 for Windows Desktop\r\nVisualStudioVersion = 12.0.30723.0\r\nMinimumVisualStudioVersion = 10.0.40219.1";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    /// <summary>
    /// Solution Format to string helper
    /// </summary>
    internal class SolutionFormatVersions
    {
        /// <summary>
        /// Converts given argument to solution file(.sln) tools entry
        /// </summary>
        /// <param name="environment">given ide environment</param>
        /// <returns>tools version entry</returns>
        public string this[IDE environment]
        {
            get
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        return "11.00";
                    //case IDE.VS2012:
                    //    return "11.00";
                    case IDE.VS20131517:
                        return "12.00";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    /// <summary>
    /// Solution Format to string helper
    /// </summary
    internal class ToolsVersions
    {
        /// <summary>
        /// Converts given argument to solution file(.sln) tools entry
        /// </summary>
        /// <param name="environment">given ide environment</param>
        /// <returns>tools version entry</returns>
        public string this[IDE environment]
        {
            get
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        return "4.0";
                    //case IDE.VS2012:
                    //    return "4.0";
                    case IDE.VS20131517:
                        return "12.00";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    /// <summary>
    /// .Net Framework Version to string helper
    /// </summary>
    internal class RuntimeVersions
    {
        /// <summary>
        /// Converts given framework version to csproj or vbproj file entry
        /// </summary>
        /// <param name="runtime">given runtime</param>
        /// <returns>framework string</returns>
        public string this[NetVersion runtime]
        {
            get
            {
                switch (runtime)
                {
                    //case NetVersion.Net2:
                    //    return "    <TargetFrameworkVersion>v2.0</TargetFrameworkVersion>";
                    //case NetVersion.Net3:
                    //    return "    <TargetFrameworkVersion>v3.0</TargetFrameworkVersion>";
                    //case NetVersion.Net35:
                    //    return "    <TargetFrameworkVersion>v3.5</TargetFrameworkVersion>";
                    case NetVersion.Net4:
                        return "    <TargetFrameworkVersion>v4.0</TargetFrameworkVersion>";
                    case NetVersion.Net4Client:
                        return "    <TargetFrameworkVersion>v4.0</TargetFrameworkVersion>\r\n    <TargetFrameworkProfile>Client</TargetFrameworkProfile>";
                    case NetVersion.Net45:
                        return "    <TargetFrameworkVersion>v4.5</TargetFrameworkVersion>";
                    case NetVersion.Net451:
                        return "    <TargetFrameworkVersion>v4.5.1</TargetFrameworkVersion>";
                    case NetVersion.Net452:
                        return "    <TargetFrameworkVersion>v4.5.2</TargetFrameworkVersion>";
                    case NetVersion.Net46:
                        return "    <TargetFrameworkVersion>v4.6</TargetFrameworkVersion>";
                    case NetVersion.Net461:
                        return "    <TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>";
                    default:
                        throw new ArgumentOutOfRangeException("runtime");
                }
            }
        }
    }

    /// <summary>
    /// Filesystem operation helper
    /// </summary>
    internal static class FileSystem
    {
        /// <summary>
        /// Moves a file to target. The use File.Move if source and target has the same root directory.
        /// Otherwise its use File.Copy and File.Delete (File.Move can't handle different root directories)
        /// </summary>
        /// <param name="sourceFileName">file to move</param>
        /// <param name="destFileName">new file target</param>
        /// <param name="overwrite">overwrite target if exists</param>
        internal static void FileMove(string sourceFileName, string destFileName, bool overwrite = false)
        {
            string sourceRoot = Path.GetPathRoot(sourceFileName);
            string destRoot = Path.GetPathRoot(destFileName);
            if (sourceRoot.Equals(destRoot, StringComparison.InvariantCultureIgnoreCase))
            {
                if (overwrite && File.Exists(destFileName))
                    File.Delete(destFileName);
                File.Move(sourceFileName, destFileName); 
            }
            else
            {
                File.Copy(sourceFileName, destFileName, overwrite);
                File.Delete(sourceFileName);
            }
        }

        /// <summary>
        /// Moves a directory to target. The use Directory.Move if source and target has the same root directory.
        /// Otherwise its use Directory.Copy and Directory.Delete (Directory.Move can't handle different root directories)
        /// </summary>
        /// <param name="sourceDirName">file to move</param>
        /// <param name="destDirName">new file target</param>
        internal static void DirectoryMove(string sourceDirName, string destDirName)
        {
            string sourceRoot = Path.GetPathRoot(sourceDirName);
            string destRoot = Path.GetPathRoot(destDirName);
            if (sourceRoot.Equals(destRoot, StringComparison.InvariantCultureIgnoreCase))
            {
                Directory.Move(sourceDirName, destDirName);
            }
            else
            {
                CopyTo(new DirectoryInfo(sourceDirName), new DirectoryInfo(destDirName));
                Directory.Delete(sourceDirName, true);
            }
        }

        /// <summary>
        /// Copy directory helper
        /// </summary>
        /// <param name="source">source directory</param>
        /// <param name="target">destination directory</param>
        private static void CopyTo(DirectoryInfo source, DirectoryInfo target)
        {
            if (!Directory.Exists(target.FullName))
                Directory.CreateDirectory(target.FullName);
            
            foreach (FileInfo fileInfo in source.GetFiles())
                fileInfo.CopyTo(Path.Combine(target.ToString(), fileInfo.Name), true);

            foreach (DirectoryInfo sourceSubDir in source.GetDirectories())
            {
                DirectoryInfo targetSubDir = target.CreateSubdirectory(sourceSubDir.Name);
                CopyTo(sourceSubDir, targetSubDir);
            }
        }
    }
}
