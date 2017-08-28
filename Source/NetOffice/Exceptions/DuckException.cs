using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates NetOffice.Core failed to compile a duck type implementation
    /// </summary>
    public class DuckException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public DuckException(string message) : base(message)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>     
        /// <param name="innerException">inner exception</param>
        public DuckException(Exception innerException) : base("Failed to compile a duck type implementation")
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public DuckException(string message, Exception innerException) : base(message)
        {

        }
    }
}
