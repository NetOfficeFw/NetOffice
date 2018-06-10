using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a NetOffice API factory failed to create an instance
    /// </summary>
    public class CreateFactoryInstanceException : CreateInstanceException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public CreateFactoryInstanceException(string message) : base(message)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public CreateFactoryInstanceException(Exception innerException) : base("Failed to create a new factory instance")
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public CreateFactoryInstanceException(string message, Exception innerException) : base(message)
        {

        }
    }
}
