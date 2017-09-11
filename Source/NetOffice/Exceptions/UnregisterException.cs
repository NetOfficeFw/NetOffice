using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates an uregister error
    /// </summary>
    public class UnregisterException : NetOfficeException
    {
        /// <summary>
        /// Default Error Message
        /// </summary>
        private static string _exceptionMessage = "An error occured while calling unregister.";
            
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public UnregisterException() : base(_exceptionMessage, null)
        {
            
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public UnregisterException(Exception innerException) : base(_exceptionMessage, innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public UnregisterException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public UnregisterException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
