using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// ComActivator Exception
    /// </summary>
    public class ActivationException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public ActivationException(Exception innerException) : base("Failed to complete instance activation.", innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public ActivationException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public ActivationException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
