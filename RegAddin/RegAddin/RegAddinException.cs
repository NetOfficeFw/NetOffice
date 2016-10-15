using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class RegAddinException : Exception
    {
        internal RegAddinException(string errorName) : base(new ErrorCodes().MessageFromName(errorName))
        {
            ReturnCode = new ErrorCodes().CodeFromName(errorName);
        }

        internal RegAddinException(int returnCode) : base(new ErrorCodes().MessageFromCode(returnCode))
        {
            ReturnCode = returnCode;
        }

        internal int ReturnCode { get; private set; }
    }
}
