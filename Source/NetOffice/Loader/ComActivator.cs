using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// 
    /// </summary>
    public static class ComActivator
    {
        /// <summary>
        /// 
        /// </summary>
        public static ICOMObject CreateInitializeInstance(Type type, ICOMObject parentObject, object comProxy, Type comProxyType)
        {
            var newInstance = (ICOMObject)Activator.CreateInstance(type);
            ICOMObjectInitialize init = (ICOMObjectInitialize)newInstance;
            init.InitializeCOMObject(parentObject, comProxy, comProxyType);
            return newInstance;
        }
    }
}
