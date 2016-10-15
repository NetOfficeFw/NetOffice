using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.RegFile
{
    internal class RegFileException : Exception
    {
        internal RegFileException(string message) : base(message)
        {

        }
    }
}
