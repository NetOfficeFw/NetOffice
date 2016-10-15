using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal class TokenConversion
    {
        internal static string ConvertToString(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return "null";

            var publicKeyToken = string.Empty;
            for (int i = 0; i < bytes.GetLength(0); i++)
                publicKeyToken += string.Format("{0:x2}", bytes[i]);

            return publicKeyToken;
        }
    }
}
