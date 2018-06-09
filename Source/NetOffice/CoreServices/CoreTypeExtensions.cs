using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using NetOffice.ComTypes;
using NetOffice.Exceptions;

namespace NetOffice.CoreServices
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

        internal static void GetComponentAndTypeId(Core value, object comProxy, ref Guid componentId, ref Guid typeId)
        {
            typeId = TypeGuid(comProxy);
            Guid parentGuid = Guid.Empty;
            if (!value.InternalCache.TypeComponentIdCache.TryGetValue(typeId, out parentGuid))
            {
                parentGuid = GetParentLibraryGuid(value, comProxy);
                value.InternalCache.TypeComponentIdCache.Add(typeId, parentGuid);
            }
        }

        /// <summary>
        /// Returns parent library(COM component) id
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>parent library/component id</returns>
        /// <exception cref="ArgumentNullException">comProxy is null</exception>
        internal static Guid GetParentLibraryGuid(this Core value, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException();

            Guid parentGuid = Guid.Empty;

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                throw new IDispatchNotImplementedException();
            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo();
            COMTypes.ITypeLib parentTypeLib = null;


            int i = 0;
            typeInfo.GetContainingTypeLib(out parentTypeLib, out i);

            IntPtr attributesPointer = IntPtr.Zero;
            parentTypeLib.GetLibAttr(out attributesPointer);

            COMTypes.TYPELIBATTR attributes =
                (COMTypes.TYPELIBATTR)Marshal.PtrToStructure(attributesPointer,
                typeof(COMTypes.TYPELIBATTR));
            parentGuid = attributes.guid;
            parentTypeLib.ReleaseTLibAttr(attributesPointer);
            Marshal.ReleaseComObject(parentTypeLib);
            
            Marshal.ReleaseComObject(typeInfo);

            return parentGuid;
        }

        /// <summary>
        /// Get type id from IDispatch GetTypeInfo
        /// </summary>
        /// <param name="comProxy">target proxy</param>
        /// <returns>type id</returns>
        internal static Guid TypeGuid(this object comProxy)
        {
            Guid typeGuid = Guid.Empty;
            if (null == comProxy)
                throw new ArgumentNullException();

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                throw new IDispatchNotImplementedException();

            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo();
            typeGuid = typeInfo.GetTypeGuid();
            Marshal.ReleaseComObject(typeInfo);

            return typeGuid;
        }

        internal static TypeInformation GetTypeInformationForUnknownObject(Core value, IFactoryInfo factoryInfo, Guid typeId, object comProxy)
        {
            TypeInformation result = value.InternalCache.TypeCache.TryGetTypeInfo(factoryInfo.ComponentGuid, typeId);
            if (null == result)
            {
                Type contract = null;
                Type implementation = null;
                if(factoryInfo.ContractAndImplementation(typeId, ref contract, ref implementation))
                    result = value.InternalCache.TypeCache.Add(contract, implementation, comProxy.GetType(), factoryInfo.ComponentGuid, typeId);
            }
            return result;
        }

        internal static TypeInformation GetTypeInformationForKnownObject(Core value, Type contractType, object comProxy)
        {
            TypeInformation result = value.InternalCache.TypeCache.TryGetTypeInfo(contractType);
            if (null != result)
            {
                Type implementationType = value.InternalFactories.FactoryAssemblies.GetImplementationType(contractType, false);
                if (null != implementationType)
                {
                    Guid typeId = Guid.Empty;
                    Guid componentId = Guid.Empty;
                    GetComponentAndTypeId(value, componentId, ref componentId, ref typeId); 
                    result = value.InternalCache.TypeCache.Add(contractType, implementationType, comProxy.GetType(), componentId, typeId);
                }
            }
            return result;
        }

        //internal static TypeInformation GetTypeInformation(Core value, object comProxy, Guid typeId, Type contractWrapperType)
        //{
        //    TypeInformation typeInfo = null;
        //    if (false == value.InternalCache.TypeCache.TryGetTypeInfo(contractWrapperType, ref typeInfo))
        //    {
        //        Type comProxyType = comProxy.GetType();
        //        Type implementationType = value.InternalFactories.FactoryAssemblies.GetImplementationType(contractWrapperType, false);
        //        if (null != implementationType)
        //        {
        //            typeInfo = new TypeInformation(contractWrapperType, implementationType, comProxyType, typeId);
        //            value.InternalCache.TypeCache.Add(typeInfo);
        //        }
        //    }
        //    return typeInfo;
        //}

        //internal static TypeInformation GetTypeInformation(Core value, object comProxy, string contractWrapperNamespace, string contractWrapperTypeName, Guid typeId)
        //{
        //    TypeInformation typeInfo = null;
        //    if (false == value.InternalCache.TypeCache.TryGetTypeInfo(contractWrapperNamespace + "." + contractWrapperTypeName, ref typeInfo))
        //    {
        //        Type comProxyType = comProxy.GetType();
        //        Type contractType = null;
        //        Type implementationType = null;
        //        if (value.InternalFactories.FactoryAssemblies.GetContractAndImplementationType(contractWrapperNamespace, contractWrapperTypeName, ref contractType, ref implementationType, false))
        //        {
        //            typeInfo = new TypeInformation(contractType, implementationType, comProxyType, typeId);
        //            value.InternalCache.TypeCache.Add(typeInfo);
        //        }
        //    }
        //    return typeInfo;
        //}
    }
}
