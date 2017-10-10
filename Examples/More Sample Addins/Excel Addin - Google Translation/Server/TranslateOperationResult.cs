using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    /// <summary>
    /// Result object for a WebTranslationService.Translate
    /// </summary>
    [Serializable]
    public class TranslateOperationResult
    {

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="state">operation state</param>
        /// <param name="result">result text</param>
        /// <param name="exception">exception (if occured)</param>
        internal TranslateOperationResult(TranslateOperationState state, string requested, string result, Exception exception, bool cached = false)
        {
            State = state;
            Requested = requested;
            Result = result;
            Exception = exception;
            Cached = cached;
        }

        /// <summary>
        /// Operation State
        /// </summary>
        public TranslateOperationState State { get; private set; }
        
        /// <summary>
        /// Search Text
        /// </summary>
        public string Requested { get; private set; }

        /// <summary>
        /// Found result in local cache
        /// </summary>
        public bool Cached { get; private set; }

        /// <summary>
        /// Result Text
        /// </summary>
        public string Result { get; private set; }

        /// <summary>
        /// Error Exception (if ocurred)
        /// </summary>
        public Exception Exception { get; private set; }
    }

    /// <summary>
    /// indicate the operation state
    /// </summary>
    [Serializable]
    public enum TranslateOperationState
    {
        /// <summary>
        /// evertything is fine
        /// </summary>
        Sucseed = 0,

        /// <summary>
        /// An error is occured
        /// </summary>
        Error = 2
    }

}
