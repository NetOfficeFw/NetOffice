using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NetOffice.Exceptions;

namespace NetOffice.IO
{
    /// <summary>
    /// Wrapper arround System.IO.Path to throw NetOfficeIOException if something failed
    /// </summary>
    public static class IOPath
    {
        /// <summary>
        ///  Returns the directory information for the specified path string
        /// </summary>
        /// <param name="path">the path of a file or directory</param>
        /// <returns>Directory information for path, or null if path denotes a root directory or is null. Returns System.String.Empty if path does not contain directory information</returns>
        public static string GetDirectoryName(string path)
        {
            try
            {
                return Path.GetDirectoryName(path);
            }
            catch (Exception exception)
            {
                throw new NetOfficeIOException(exception);
            }
        }

        /// <summary>
        /// Returns the file name and extension of the specified path string
        /// </summary>
        /// <param name="path">the path string from which to obtain the file name and extension</param>
        /// <returns>The characters after the last directory character in path</returns>
        public static string GetFileName(string path)
        {
            try
            {
                return Path.GetFileName(path);
            }
            catch (Exception exception)
            {
                throw new NetOfficeIOException(exception);
            }
        }
    }
}
