/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// Represents an outer aggregator by a COM Shim
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("E8E14A9B-6FB4-45A6-BFF2-47610F68D075")]
    public interface IOuterComAggregator
    {
        /// <summary>
        /// Publish managed addin as IUnknown* to the Shim
        /// </summary>
        /// <param name="innnerAddinAsUnknown">managed addin as IUnknown*</param>
        void SetInnerAddin(IntPtr innnerAddinAsUnknown);
    }
}
