/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    ///  Represents an outer aggregator by a shim that handle update/reload possibilites
    /// </summary>
    [ComImport, Guid("F7BCF161-FCB2-4880-9C33-78C456B1F291")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOuterUpdateAggregator
    {
        /// <summary>
        /// Recreate the managed appdomain and create a new instance of the managed addin
        /// </summary>
        void Reload();
    }
}
