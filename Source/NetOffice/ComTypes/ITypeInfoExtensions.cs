using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace NetOffice.ComTypes
{
    internal static class ITypeInfoExtensions
    {
        /// <summary>
        /// Returns id of an interface
        /// </summary>
        /// <param name="typeInfo">com type information</param>
        /// <returns>interface id(iid)</returns>
        internal static Guid GetTypeGuid(this COMTypes.ITypeInfo typeInfo)
        {
            IntPtr attribPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out attribPtr);
            COMTypes.TYPEATTR Attributes = (COMTypes.TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(COMTypes.TYPEATTR));
            Guid typeGuid = Attributes.guid;
            typeInfo.ReleaseTypeAttr(attribPtr);
            return typeGuid;
        }
    }
}
