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
    [ComImport, Guid("EF2F0985-2D4F-45AA-ADB6-510271A6EFC3")]
    public interface IShimUpdateHost
    {
        /// <summary>
        /// Signalize the update is done
        /// The host want unload the update handler and reload the managed addin
        /// </summary>
        void Done();
    }
}
