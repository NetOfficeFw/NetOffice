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
        /// Get wrapper class factory info
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="hostCache">core host cache</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wantTheDuck">want duck implementation</param>
        /// <param name="throwException">throw exception if no info found or return null</param>
        /// <returns>factory info from corresponding assembly</returns>
        internal static IFactoryInfo GetFactoryInfo(this Core value, 
            Dictionary<Guid, Guid> hostCache,
            object comProxy, bool wantTheDuck, bool throwException)
        {
            if (value.Assemblies.Count == 0)
                return null;

            string className = TypeDescriptor.GetClassName(comProxy);
            Guid hostGuid = GetParentLibraryGuid(value, comProxy);

            foreach (IFactoryInfo item in value.Assemblies)
            {
                if (item.IsDuck != wantTheDuck)
                    continue;
                foreach (var guid in item.ComponentGuid)
                    if (true == guid.Equals(hostGuid))
                        return item;
            }

            // failback because some types was multiple defined (not allowed in COM but in fact MS do this)
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
        /// <returns>parent library/component id</returns>
        internal static Guid GetParentLibraryGuid(this Core value, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException();

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                return Guid.Empty;

            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo(0, 0);
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
                return Guid.Empty;

            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo(0, 0);
            typeGuid = typeInfo.GetTypeGuid();
            Marshal.ReleaseComObject(typeInfo);

            return typeGuid;
        }
    }
}
