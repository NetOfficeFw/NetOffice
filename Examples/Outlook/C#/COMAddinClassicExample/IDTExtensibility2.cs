using System;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace COMAddinClassicExample
{
    /*
     * taken from Extensibility2
     * you can also use the reference to extensibility2.dll but i like to keep the assembly references low
     */
    [ComVisible(true)]
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup,
        ext_cm_Startup,
        ext_cm_External,
        ext_cm_CommandLine,
        ext_cm_Solution,
        ext_cm_UISetup
    }

    [ComVisible(true)]
    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown,
        ext_dm_UserClosed,
        ext_dm_UISetupComplete,
        ext_dm_SolutionClosed
    }

    [ComVisible(true)]
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType((short)0x1040)]
    public interface IDTExtensibility2
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application, [In] ext_ConnectMode ConnectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
}
