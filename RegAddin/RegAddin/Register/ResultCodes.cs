using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Register
{
    internal enum ResultCodes
    {
        Okay = 0,
        AssemblyNotComVisible = -2,
        NothingFound = -3,
        RegisterCallFailed = -4,
        AssemblyNotSigned = -5        
    }
}
