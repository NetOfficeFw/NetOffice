using System;

namespace NetOffice.Resolver
{
    /// <summary>
    /// Uri Conversion methods because System.Uri doesnt handle special characters as well
    /// </summary>
    internal static class UriResolver
    {
        /// <summary>
        /// Convert file: path to local
        /// </summary>
        /// <param name="path">target path as any</param>
        /// <returns>converted path</returns>
        public static string ResolveLocalPath(string path)
        {
            if (null == path)
                throw new ArgumentNullException();

            if (path.StartsWith("file:///"))
            {
                path = path.Substring("file:///".Length);
                path = path.Replace("/", "\\");
                return path;
            }
            else
                return path.Replace("/", "\\");
        }
    }
}
