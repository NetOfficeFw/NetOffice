using System;
using System.ComponentModel;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal class CultureInfoConversion
    {
        internal static string ConvertToString(CultureInfo cultureInfo)
        {
            string result = new CultureInfoConverter().ConvertToInvariantString(cultureInfo);
            if (String.IsNullOrWhiteSpace(result))
                result = "neutral";
            return result;
        }
    }
}
