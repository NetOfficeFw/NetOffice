using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates NetOffice.Core failed to recieve required factory info
    /// </summary>
    public class FactoryException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public FactoryException(string message) : base(message)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public FactoryException(Exception innerException) : base("Failed to recieve required factory info")
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public FactoryException(string message, Exception innerException) : base(message)
        {

        }
    }
}
