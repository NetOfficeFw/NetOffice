using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// Notify Addins about any exception in base class COMAddin
    /// </summary>
    public interface IErrorHandler
    {
        /// <summary>
        /// Error Handler method
        /// </summary>
        /// <param name="exception">the occured exception</param>
        /// <returns>true when the error is handled by the client</returns>
        bool OnError(System.Exception exception);
    }
}
