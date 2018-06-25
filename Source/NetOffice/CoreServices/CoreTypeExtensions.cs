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
        /// <exception cref="ArgumentNullException">core or value is null </exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public static object WrapObject(this Core factory, object value, bool allowDynamicObject)
        {
            if (null == factory)
                throw new ArgumentNullException("factory");
            if (null == value)
                throw new ArgumentNullException("value");

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

        /// <summary>
        /// Retrieve component and type id for given com proxy
        /// </summary>
        /// <param name="value">core to extend</param>
        /// <param name="comProxy">com proxy as any</param>
        /// <param name="componentId">component id from com proxy</param>
        /// <param name="typeId">type id from com proxy</param>
        /// <exception cref="ArgumentNullException">core or comProxy is null </exception>
        internal static void GetComponentAndTypeId(this Core value, object comProxy, ref Guid componentId, ref Guid typeId)
        {
            if (null == value)
                throw new ArgumentNullException("value");
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");

            typeId = TypeGuid(comProxy);
            if (!value.InternalCache.TypeComponentIdCache.TryGetValue(typeId, out componentId))
            {
                componentId = GetParentLibraryGuid(value, comProxy);
                value.InternalCache.TypeComponentIdCache.Add(typeId, componentId);
            }
        }

        /// <summary>
        /// Resolve type informations for an unknown proxy
        /// </summary>
        /// <param name="value">Core to extend</param>
        /// <param name="typeFactory">corresponding factory from com proxy</param>
        /// <param name="typeId">type id from com proxy</param>
        /// <param name="comProxy">target com proxy</param>
        /// <returns>type information or null if its not a known type</returns>
        /// <exception cref="ArgumentNullException">core or typeFactory or com proxy is null </exception>
        internal static TypeInformation GetTypeInformationForUnknownObject(this Core value, ITypeFactory typeFactory, Guid typeId, object comProxy)
        {
            if (null == value)
                throw new ArgumentNullException("value");
            if (null == typeFactory)
                throw new ArgumentNullException("typeFactory");
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");

            TypeInformation result = value.InternalCache.TypeCache.TryGetTypeInfo(typeFactory.ComponentID, typeId);
            if (null == result)
            {
                Type contract = null;
                Type implementation = null;
                if(value.InternalFactories.FactoryAssemblies.GetContractAndImplementationType(typeFactory, typeId, ref contract, ref implementation))
                    result = value.InternalCache.TypeCache.Add(typeFactory, contract, implementation, comProxy.GetType(), typeFactory.ComponentID, typeId);
            }
            return result;
        }

        /// <summary>
        /// Resolve type informations for a proxy by known result contract
        /// </summary>
        /// <param name="value">Core to extend</param>
        /// <param name="contractType">known contract</param>
        /// <param name="comProxy">target com proxy</param>
        /// <returns>type information or null if its failed to resolve</returns>
        /// <exception cref="ArgumentNullException">core or contractType or com proxy is null</exception>
        internal static TypeInformation GetTypeInformationForKnownObject(this Core value, Type contractType, object comProxy)
        {
            if (null == value)
                throw new ArgumentNullException("value");
            if (null == contractType)
                throw new ArgumentNullException("contractType");
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");

            TypeInformation result = value.InternalCache.TypeCache.TryGetTypeInfo(contractType);
            if (null == result)
            {
                ITypeFactory typeFactory = null;
                Type implementationType = value.InternalFactories.FactoryAssemblies.GetImplementationType(contractType, ref typeFactory, false);
                if (null != implementationType)
                {
                    Guid typeId = Guid.Empty;
                    Guid componentId = Guid.Empty;
                    GetComponentAndTypeId(value, comProxy, ref componentId, ref typeId);
                    result = value.InternalCache.TypeCache.Add(typeFactory, contractType, implementationType, comProxy.GetType(), componentId, typeId);
                }
            }
            return result;
        }


        /// <summary>
        /// Resolve type informations for a proxy by known result contract
        /// </summary>
        /// <param name="value">Core to extend</param>
        /// <param name="contractType">known contract</param>
        /// <returns>type information or null if its failed to resolve</returns>
        /// <exception cref="ArgumentNullException">core or contractType or com proxy is null</exception>
        internal static Type GetImplementationTypeForKnownObject(this Core value, Type contractType)
        {
            if (null == value)
                throw new ArgumentNullException("value");
            if (null == contractType)
                throw new ArgumentNullException("contractType");

            value.CheckInitialize();

            Type result = null;
            
            TypeInformation typeInfo = value.InternalCache.TypeCache.TryGetTypeInfo(contractType);
            if (null != typeInfo)
            {
                result = typeInfo.Implementation;
            }
            else
            {
                ITypeFactory typeFactory = null;
                result = value.InternalFactories.FactoryAssemblies.GetImplementationType(contractType, ref typeFactory, false);
            }
            return result;
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
    }
}
