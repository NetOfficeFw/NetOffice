using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Register
{
    internal class RegisterException : Exception
    {
        internal RegisterException(string message) : base(message)
        {

        }
    }
}
