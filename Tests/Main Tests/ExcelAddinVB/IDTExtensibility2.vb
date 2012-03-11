Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")> _
Public Enum ext_ConnectMode
    ext_cm_AfterStartup = 0
    ext_cm_CommandLine = 3
    ext_cm_External = 2
    ext_cm_Solution = 4
    ext_cm_Startup = 1
    ext_cm_UISetup = 5
End Enum

<Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")> _
Public Enum ext_DisconnectMode
    ext_dm_HostShutdown = 0
    ext_dm_SolutionClosed = 3
    ext_dm_UISetupComplete = 2
    ext_dm_UserClosed = 1
End Enum

<ComImport(), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(CShort(&H1040))> _
Public Interface IDTExtensibility2
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(1)> _
    Sub OnConnection(<[In](), MarshalAs(UnmanagedType.IDispatch)> ByVal Application As Object, <[In]()> ByVal ConnectMode As ext_ConnectMode, <[In](), MarshalAs(UnmanagedType.IDispatch)> ByVal AddInInst As Object, <[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(2)> _
    Sub OnDisconnection(<[In]()> ByVal RemoveMode As ext_DisconnectMode, <[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(3)> _
    Sub OnAddInsUpdate(<[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(4)> _
    Sub OnStartupComplete(<[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(5)> _
    Sub OnBeginShutdown(<[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
End Interface
