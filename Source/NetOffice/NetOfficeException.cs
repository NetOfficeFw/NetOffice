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
        /// <param name="message">given exception info</param>
        public NetOfficeException(string message) : base(message)
        { }

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public NetOfficeException(string message, Exception innerException): base(message, innerException)
        {
        }
    }
}
