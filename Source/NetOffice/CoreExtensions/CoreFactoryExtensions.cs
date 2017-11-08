using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using NetOffice.ComTypes;
using NetOffice.Exceptions;
    
namespace NetOffice
{
    internal static class CoreFactoryExtensions
    {
        /// <summary>
        /// See Attributes\DuplicateAttribute.cs for further informations
        /// </summary>
        private static Guid[] _duplicateTypes = new Guid[]
            {
                new Guid("000C0310-0000-0000-C000-000000000046"),
                new Guid("000C0311-0000-0000-C000-000000000046"),
                new Guid("000C0312-0000-0000-C000-000000000046"),
                new Guid("000C0314-0000-0000-C000-000000000046"),
                new Guid("000C0317-0000-0000-C000-000000000046"),
                new Guid("000C0318-0000-0000-C000-000000000046"),
                new Guid("000C0319-0000-0000-C000-000000000046"),
                new Guid("000C031A-0000-0000-C000-000000000046"),
                new Guid("000C031B-0000-0000-C000-000000000046"),
                new Guid("000C031F-0000-0000-C000-000000000046"),
                new Guid("000C0321-0000-0000-C000-000000000046"),
                new Guid("000C036E-0000-0000-C000-000000000046"),
                new Guid("000C036F-0000-0000-C000-000000000046"),
                new Guid("000C0370-0000-0000-C000-000000000046"),
                new Guid("000C0398-0000-0000-C000-000000000046")
            };
        
        /// <summary>
        /// Get wrapper class factory info
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="hostCache">core host cache</param>
        /// <param name="caller">calling instance</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wantTheDuck">want duck implementation</param>
        /// <param name="throwException">throw exception if no info found or return null</param>
        /// <returns>factory info from corresponding assembly</returns>
        internal static IFactoryInfo GetFactoryInfo(this Core value, 
            Dictionary<Guid, Guid> hostCache, ICOMObject caller,
            object comProxy, bool wantTheDuck, bool throwException)
        {
            if (value.Assemblies.Count == 0)
                return null;

            string className = ComTypes.TypeDescriptor.GetClassName(comProxy);
            Guid typeid = TypeGuid(comProxy);
            Guid hostGuid = GetParentLibraryGuid(value, comProxy, typeid);

            if (null != caller && typeid.IsDuplicateType())
            {             
                foreach (IFactoryInfo item in value.Assemblies)
                {                   
                    if (item.IsDuck != wantTheDuck || item.AssemblyName != caller.InstanceComponentName)
                        continue;
                    foreach (var guid in item.ComponentGuid)
                        if (true == guid.Equals(hostGuid))
                            return item;
                }
            }
            else
            {
                foreach (IFactoryInfo item in value.Assemblies)
                {
                    if (item.IsDuck != wantTheDuck)
                        continue;
                    foreach (var guid in item.ComponentGuid)
                        if (true == guid.Equals(hostGuid))
                            return item;
                }
            }

            // failback because some types was multiple defined by its class id (not allowed in COM but in fact MS do this)
            // list of known multiple defined types is available on netoffice.codeplex.com or Attributes\DuplicateAttribute.cs
            foreach (IFactoryInfo item in value.Assemblies)
            {
                if (item.IsDuck != wantTheDuck)
                    continue;
                bool hasComponentID = null != item.ComponentGuid ? item.ComponentGuid.Contains(hostGuid) : false;
                if (item.Contains(className) && hasComponentID)
                {
                    value.Console.WriteLine("Failback factory {0}=>{1}recieved.", item.Assembly.FullName, className);
                    return item;
                }
            }

            if (throwException)
            {
                string message = string.Format("Class {0}:{1} not found in loaded NetOffice Assemblies{2}", hostGuid, className, Environment.NewLine);
                message += string.Format("Currently loaded NetOfficeApi Assemblies{0}", Environment.NewLine);
                foreach (IFactoryInfo item in value.Assemblies)
                    message += string.Format("Loaded NetOffice Assembly:{0} {1}{2}", item.ComponentGuid, item.Assembly.FullName, Environment.NewLine);

                throw new FactoryException(message);
            }
            else
                return null;
        }
        
        /// <summary>
        /// Returns parent library id
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="typeGuid">type id from comProxy</param>
        /// <returns>parent library/component id</returns>
        internal static Guid GetParentLibraryGuid(this Core value, object comProxy, Guid typeGuid)
        {
            if (null == comProxy)
                throw new ArgumentNullException();

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                throw new IDispatchNotImplementedException();
            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo();
            COMTypes.ITypeLib parentTypeLib = null;
            Guid parentGuid = Guid.Empty;

            if (!value.HostCache.TryGetValue(typeGuid, out parentGuid))
            {
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

                value.HostCache.Add(typeGuid, parentGuid);
            }

            Marshal.ReleaseComObject(typeInfo);

            return parentGuid;
        }

        /// <summary>
        /// Returns parent library id
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>parent library/component id</returns>
        internal static Guid GetParentLibraryGuid(this Core value, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException();

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                throw new IDispatchNotImplementedException();

            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo();
            COMTypes.ITypeLib parentTypeLib = null;
            Guid typeGuid = typeInfo.GetTypeGuid();
            Guid parentGuid = Guid.Empty;

            if (!value.HostCache.TryGetValue(typeGuid, out parentGuid))
            {
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

                value.HostCache.Add(typeGuid, parentGuid);
            }

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

        /// <summary>
        /// Returns information the type is a known external duplicate
        /// </summary>
        /// <param name="value">type id</param>
        /// <returns>true if duplicate, otherwise false</returns>
        internal static bool IsDuplicateType(this Guid value)
        {
            return _duplicateTypes.Contains(value);
        }

        /// <summary>
        /// Performs GetTypeInfo on IDispatch
        /// Handle the strange cast behavior - see remarks. 
        /// </summary>
        /// <remarks>
        /// Seems to be cast to IDispatch never failed
        /// even the instance behind comProxy doesnt implement the interface.
        /// If its failed to cast, an InvalidCastException occurs 
        /// while first use the interface.
        /// The method catch arround here and throws IDispatchNotImplementedException 
        /// to signalize the missing IDispatch suport.    
        /// </remarks>
        /// <param name="dispatcher">given IDispatch as any </param>
        /// <returns>type informations or null if dispatcher argument is null</returns>
        internal static COMTypes.ITypeInfo GetTypeInfo(this IDispatch dispatcher)
        {
            if (null == dispatcher)
                return null;
            try
            {
                return dispatcher.GetTypeInfo(0, 0);
            }
            catch (InvalidCastException exception)
            {
                throw new IDispatchNotImplementedException(exception);
            }
            catch
            {
                throw;
            }
        }
    }
}