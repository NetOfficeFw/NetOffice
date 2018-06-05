using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates NetOffice failed to initialize
    /// </summary>
    public class NetOfficeInitializeException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public NetOfficeInitializeException(Exception innerException) : base("Failed to initialize.", innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public NetOfficeInitializeException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public NetOfficeInitializeException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
