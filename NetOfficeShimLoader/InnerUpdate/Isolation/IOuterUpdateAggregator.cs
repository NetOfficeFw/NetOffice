/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    ///  Represents an outer shim host for a managed update handler
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("E20B53FD-03C8-4977-8725-7E0C89657960")]
    public interface IOuterUpdateAggregator
    {
        /// <summary>
        /// Publish managed addin as IUnknown* to the shim host
        /// </summary>
        /// <param name="innnerHandlerAsUnknown">managed addin as IUnknown*</param>
        void SetInnerHandler(IntPtr innnerHandlerAsUnknown);
    }
}
