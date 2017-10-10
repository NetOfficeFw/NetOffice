using System;
using System.Runtime.InteropServices;

namespace NetOffice.Filtering
{
    /// <summary>
    /// http://msdn.microsoft.com/en-us/library/windows/desktop/ms683793%28v=vs.85%29.aspx
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct InterfaceInfo
    {
        /// <summary>
        /// A pointer to the IUnknown interface on the object
        /// </summary>
        [MarshalAs(UnmanagedType.IUnknown)]
        public object punk;

        /// <summary>
        /// The identifier of the requested interface
        /// </summary>
        public Guid iid;

        /// <summary>
        /// The interface method
        /// </summary>
        public ushort wMethod;
    }
}
