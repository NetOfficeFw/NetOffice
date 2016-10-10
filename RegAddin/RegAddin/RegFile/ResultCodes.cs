using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.RegFile
{
    internal enum ResultCodes
    {
        Okay = 0,
        InvalidRegfilePath = -1,
        AssemblyNotComVisible = -2,
        NothingFound = -3
    }
}
