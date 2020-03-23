using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Exceptions;

namespace NetOffice.ExcelApi.Tools.Exceptions
{
    /// <summary>
    /// Indicates which method in Rtd server implementation cause an error
    /// </summary>
    public enum RTDMethods
    {
        /// <summary>
        /// Error occured not IRtdServer method - see stacktrace for further information 
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// IRtdServer.ServerStart
        /// </summary>
        ServerStart = 1,

        /// <summary>
        /// IRtdServer.ConnectData
        /// </summary>
        ConnectData = 2,

        /// <summary>
        /// IRtdServer.RefreshData
        /// </summary>
        RefreshData = 3,

        /// <summary>
        /// IRtdServer.DisconnectData
        /// </summary>
        DisconnectData = 4,

        /// <summary>
        /// IRtdServer.Heartbeat
        /// </summary>
        Heartbeat = 5,

        /// <summary>
        /// IRtdServer.ServerTerminate
        /// </summary>
        ServerTerminate = 6
    }

    /// <summary>
    /// An exception occured in RTD Server implementation
    /// </summary>
    public class COMRtdServerException : NetOfficeException
    {
        /// <summary>
        /// Indicates which method in Rtd server implementation cause an error
        /// </summary>
        public readonly RTDMethods Method;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="inner">inner exception</param>   
        public COMRtdServerException(Exception inner) : base(inner)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="method">indicates which method in Rtd server implementation cause an error</param>   
        /// <param name="inner">inner exception</param>   
        public COMRtdServerException(RTDMethods method, Exception inner) : base(inner)
        {
            Method = method;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        ///  <param name="method">indicates which method in Rtd server implementation cause an error</param> 
        /// <param name="message">given error message as any</param>
        /// <param name="inner">inner exception</param>
        public COMRtdServerException(RTDMethods method, string message, Exception inner) : base(message, inner)
        {
            Method = method;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given error message as any</param>
        /// <param name="inner">inner exception</param>
        public COMRtdServerException(string message, Exception inner) : base(message, inner)
        {

        }
    }
}
