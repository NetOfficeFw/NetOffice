using System;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace COMAddinRibbonExample
{

    /* 
     * why i have re-defined the ribbon interfaces here ? 
     *
     * NetOffice.OfficeApi.dll includes the 3 Ribbon Interfaces but if you implement these interfaces from the NetOffice assembly
     * in your Addin class then Visual Studio try to register the NetOffice.OfficeApi.dll for COM Interop while compiling. 
     *
     * I have no solution for these problem at the moment. (re)define the interfaces here in the project works fine!
     * Do you know a better way?
     * contact me: public.sebastian@web.de
     */

    #region Ribbon Interfaces
    [ComImport, Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType((short)0x1040)]
    public interface IRibbonUI
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Invalidate();
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        void InvalidateControlMso([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        void ActivateTab([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
        void ActivateTabMso([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
        void ActivateTabQ([In, MarshalAs(UnmanagedType.BStr)] string ControlID, [In, MarshalAs(UnmanagedType.BStr)] string Namespace);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType((short)0x1040)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)] get; }
        [DispId(2)]
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)] get; }
        [DispId(3)]
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)] get; }
    }

    [ComImport, TypeLibType((short)0x1040), Guid("000C0396-0000-0000-C000-000000000046")]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    #endregion
}
