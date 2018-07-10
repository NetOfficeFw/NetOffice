/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.Text;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// To implement by a managed update handler to recieve an outer shim host
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("BA23F519-0F53-4EC7-A416-2681BE22150F")]
    public interface IManagedInnerUpdateHandler
    {
        /// <summary>
        /// Set an outer shim host to the update handler
        /// </summary>
        /// <param name="shim">outer shim host</param>
        void SetParent([In, MarshalAs(UnmanagedType.IUnknown)] IShimUpdateHost shim);

        /// <summary>
        /// Set custom data from addin instance to the update handler
        /// </summary>
        /// <param name="custom">custom data as any</param>
        void SetCustomData([In, MarshalAs(UnmanagedType.BStr)] string custom);

        /// <summary>
        /// Set the host application to the update handler if OnConnection is already passed
        /// </summary>
        /// <param name="application">host application </param>
        void SetApplication([In, MarshalAs(UnmanagedType.IDispatch)] object application);

        /// <summary>
        /// Determines the update handler supports direct execution
        /// Outer update aggregator want execute and close the handler if execution is supported
        /// </summary>
        /// <param name="canExecute">true if direct execution is supported, otherwise false</param>
        void CanExecute([In, Out, MarshalAs(UnmanagedType.Bool)] ref bool canExecute);

        /// <summary>
        /// Execute the handler
        /// </summary>
        void Execute();

        /// <summary>
        /// Called before the outer shim is unload the AppDomain
        /// </summary>
        void Close();
    }
}
