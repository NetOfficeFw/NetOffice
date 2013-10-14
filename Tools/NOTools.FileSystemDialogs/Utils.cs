using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    internal static class Utils
    {
        public static bool IsTrue(bool defaultBool, DefaultBoolean val)
        {
            if (true == defaultBool && IsTrueOrDefault(val))
                return true;
            else
                return false;
        }

        public static bool IsTrueOrDefault(DefaultBoolean val)
        {
            if (val == DefaultBoolean.Default || val == DefaultBoolean.True)
                return true;
            else
                return false;
        }
    }
}
