using System;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using System.Security;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Extend System.Runtime.InteropServices.COMException
    /// </summary>
    public class NetOfficeCOMException : COMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public NetOfficeCOMException() : base()
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="inner">the exception that is the cause of the current exception</param>
        public NetOfficeCOMException(Exception inner) : base(null != inner ? inner.Message : "<Error>", inner)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        public NetOfficeCOMException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="inner">the exception that is the cause of the current exception</param>
        public NetOfficeCOMException(string message, Exception inner) : base(message, inner)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="errorCode">The error code (HRESULT) value associated with this exception</param>
        public NetOfficeCOMException(string message, int errorCode) : base(message, errorCode)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="info">the SerializationInfo object that holds the serialized object data</param>
        /// <param name="context">the StreamingContext object that supplies the contextual information about the source or destination</param>
        /// <exception cref="ArgumentNullException">info is null</exception>
        [SecuritySafeCritical]
        public NetOfficeCOMException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
