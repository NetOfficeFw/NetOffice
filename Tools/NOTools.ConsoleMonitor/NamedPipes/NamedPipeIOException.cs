using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// This exception is thrown by named pipes communication methods.
    /// </summary>
    #endregion
    public class NamedPipeIOException : InterProcessIOException
    {
        #region Comments
        /// <summary>
        /// Creates a NamedPipeIOException instance.
        /// </summary>
        /// <param name="text">The error message text.</param>
        #endregion
        public NamedPipeIOException(String text)
            : base(text)
        {
        }
        #region Comments
        /// <summary>
        /// Creates a NamedPipeIOException instance.
        /// </summary>
        /// <param name="text">The error message text.</param>
        /// <param name="errorCode">The native error code.</param>
        #endregion
        public NamedPipeIOException(String text, uint errorCode)
            : base(text)
        {
            this.ErrorCode = errorCode;
            if (errorCode == NamedPipeNative.ERROR_CANNOT_CONNECT_TO_PIPE)
            {
                this.IsServerAvailable = false;
            }
        }
        #region Comments
        /// <summary>
        /// Creates a NamedPipeIOException instance.
        /// </summary>
        /// <param name="info">The serialization information.</param>
        /// <param name="context">The streaming context.</param>
        #endregion
        protected NamedPipeIOException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
