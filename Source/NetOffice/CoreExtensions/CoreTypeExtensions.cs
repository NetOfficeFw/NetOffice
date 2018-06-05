using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.CoreServices.Internal;

namespace NetOffice
{
    /// <summary>
    /// Provides type convert extension
    /// </summary>
    public static class CoreTypeExtensions
    {
        /// <summary>
        /// Analyze an object and create wrapper arround if necessary
        /// </summary>
        /// <param name="factory">core to extend</param>
        /// <param name="value">value as any</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>value or wrapped value</returns>
        public static object WrapObject(this Core factory, object value, bool allowDynamicObject)
        {
            if ((null != value) && (value is MarshalByRefObject))
            {
                ICOMObject newObject = factory.CreateObjectFromComProxy(null, value, allowDynamicObject);
                return newObject;
            }
            else
            {
                return value;
            }
        }

        internal static TypeInformation GetTypeInformation(Core value, object comProxy, Type contractWrapperType)
        {
            TypeInformation typeInfo = null;
            if (false == value.TypeCache.TryGetTypeInfo(contractWrapperType, ref typeInfo))
            {
                Type comProxyType = comProxy.GetType();
                Type implementationType = value.Assemblies.GetImplementationType(contractWrapperType, false);
                if (null != implementationType)
                {
                    typeInfo = new TypeInformation(contractWrapperType, implementationType, comProxyType);
                    value.TypeCache.Add(typeInfo);
                }
            }
            return typeInfo;
        }

        internal static TypeInformation GetTypeInformation(Core value, object comProxy, string contractWrapperNamespace, string contractWrapperTypeName)
        {
            TypeInformation typeInfo = null;
            if (false == value.TypeCache.TryGetTypeInfo(contractWrapperNamespace + "." + contractWrapperTypeName, ref typeInfo))
            {
                Type comProxyType = comProxy.GetType();
                Type contractType = null;
                Type implementationType = null;
                if (value.Assemblies.GetContractAndImplementationType(contractWrapperNamespace, contractWrapperTypeName, ref contractType, ref implementationType, false))
                {
                    typeInfo = new TypeInformation(contractType, implementationType, comProxyType);
                    value.TypeCache.Add(typeInfo);
                }
            }
            return typeInfo;
        }
    }
}
