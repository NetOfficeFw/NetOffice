using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Unregister
{
    internal class UnregisterException : Exception
    {
        internal UnregisterException(string message) : base(message)
        {

        }
    }
}
