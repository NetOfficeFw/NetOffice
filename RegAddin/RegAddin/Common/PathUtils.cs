using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace RegAddin.Common
{
    internal class PathUtils
    {
        internal static bool TryCheckForValidLocalAndAbsoluteFileSystemPath(string path)
        {
            try
            {
                var invalidPathChars = Path.GetInvalidPathChars();
                var invalidFileChars = Path.GetInvalidFileNameChars();

                foreach (var item in invalidPathChars)
                {
                    if (path.IndexOf(item, 0) > -1)
                        return false;
                }

                foreach (var item in invalidFileChars)
                {
                    var testString = item.ToString() + item.ToString();
                    if (path.IndexOf(testString, 0) > -1)
                        return false;
                }

                // regex ^([a-zA-Z]\:)(\\[^\\/:*?<>"|]*(?<![ ]))*(\.[a-zA-Z]{2,6})$

                Path.Combine(path);

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
