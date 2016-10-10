using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Unregister
{
    internal enum ResultCodes
    {
        Okay = 0,      
        NothingFound = -3,
        UnRegisterCallFailed = -4,
        AssemblyNotSigned = -5
    }
}
