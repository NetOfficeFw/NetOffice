using System;

namespace NetOffice
{
    /// <summary>
    /// signals an exception occured in LateBindingApi, not in corresond latebinded assembly
    /// </summary>
    public class LateBindingApiException : Exception 
    {
        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="message"></param>
        public LateBindingApiException(string message): base(message)
        { }
    }
}
