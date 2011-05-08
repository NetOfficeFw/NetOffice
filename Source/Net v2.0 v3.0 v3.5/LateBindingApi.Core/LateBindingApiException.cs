using System;

namespace LateBindingApi.Core
{
    public class LateBindingApiException : Exception 
    {
        public LateBindingApiException(string message): base(message)
        { }
    }
}
