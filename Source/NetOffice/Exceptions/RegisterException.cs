using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a register error
    /// </summary>
    public class RegisterException : NetOfficeException
    {
        /// <summary>
        /// Default Error Message
        /// </summary>
        private static string _exceptionMessage = "An error occured while calling register.";

        /// <summary>
        /// Codeblock
        /// </summary>
        public readonly int ErrorBlock;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="errorBlock">codeblock</param>
        public RegisterException(int errorBlock) : base(_exceptionMessage, null)
        {
            ErrorBlock = errorBlock;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="errorBlock">codeblock</param>
        /// <param name="innerException">inner exception</param>
        public RegisterException(int errorBlock, Exception innerException) : base(_exceptionMessage, null)
        {
            ErrorBlock = errorBlock;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public RegisterException(Exception innerException) : base(_exceptionMessage, innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public RegisterException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public RegisterException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
