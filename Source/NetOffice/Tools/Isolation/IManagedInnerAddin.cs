/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// To implement by a managed addin to recieve an outer shim host
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("EF261BCD-3078-459E-9448-13845BEED136")]
    public interface IManagedInnerAddin
    {
        /// <summary>
        /// Set an outer shim host to the managed addin
        /// </summary>
        /// <param name="shim">outer shim host</param>
        void SetParent([In, Out, MarshalAs(UnmanagedType.IUnknown)] IShimHost shim);

        /// <summary>
        /// Notifies the adddin that its loaded by an update request
        /// </summary>
        /// <param name="custom">custom data given in prev managed addin instance (possibly modified by an update handler)</param>
        void ReloadNotification([In, MarshalAs(UnmanagedType.BStr)]string custom);
    }
}
