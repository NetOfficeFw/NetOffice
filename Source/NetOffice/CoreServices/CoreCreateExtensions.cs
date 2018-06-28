using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.CoreServices.Internal;
using NetOffice.Exceptions;

namespace NetOffice.CoreServices
{
    internal static class CoreCreateExtensions
    {
        internal static ICOMObject CreateInstance(Core value, TypeInformation typeInfo, ICOMObject caller, object comProxy)
        {
            ICOMObject newInstance = null;
            try
            {
                newInstance = ComActivator.CreateInitializeInstance(value, typeInfo.Implementation, typeInfo.Factory, caller, comProxy, typeInfo.Proxy);
                newInstance = value.InternalObjectActivator.TryReplaceInstance(caller, newInstance);
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }
            return newInstance;
        }

        internal static ICOMObject TryCreateObjectByResolveEvent(Core value, ICOMObject caller, Type contractType, object comProxy)
        {
            ICOMObject result = value.InternalObjectResolver.RaiseResolve(caller, contractType, comProxy);
            if (null != result)
            {
                ICOMObjectInitialize init = result as ICOMObjectInitialize;
                if (null != init && false == init.IsInitialized)
                {
                    init.InitializeCOMObject(caller, comProxy);
                }
            }
            return result;
        }
    }
}
