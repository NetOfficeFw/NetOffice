using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates NetOffice.Core failed to create an instance
    /// </summary>
    public class CreateInstanceException : NetOfficeCOMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public CreateInstanceException(string message) : base(message)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public CreateInstanceException(Exception innerException) : base("Failed to create a new instance")
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public CreateInstanceException(string message, Exception innerException) : base(message)
        {

        }
    }
}
