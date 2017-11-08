using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates an I/O Error
    /// </summary>
    public class NetOfficeIOException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public NetOfficeIOException(Exception innerException) : base(innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">inner exception</param>
        public NetOfficeIOException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
