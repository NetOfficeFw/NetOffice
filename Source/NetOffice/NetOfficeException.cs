using System;

namespace NetOffice
{
    /// <summary>
    /// signals an exception occured in NetOffice.dll, not in corresonding NetOffice assembly
    /// </summary>
    public class NetOfficeException : Exception 
    {
        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="message"></param>
        public NetOfficeException(string message) : base(message)
        { }
    }
}
