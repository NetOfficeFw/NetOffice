using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates NetOffice.Core failed to create a COMProxyShare instance
    /// </summary>
    public class CreateCOMProxyShareException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public CreateCOMProxyShareException(string message) : base(message)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public CreateCOMProxyShareException(Exception innerException) : base("Failed to create a new COMProxyShare instance")
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public CreateCOMProxyShareException(string message, Exception innerException) : base(message)
        {

        }
    }
}
