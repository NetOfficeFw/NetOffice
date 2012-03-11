Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport(), Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType(CShort(&H1040))> _
Public Interface IRibbonUI
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(1)> _
    Sub Invalidate()
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(2)> _
    Sub InvalidateControl(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ControlID As String)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(3)> _
    Sub InvalidateControlMso(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ControlID As String)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(4)> _
    Sub ActivateTab(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ControlID As String)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(5)> _
    Sub ActivateTabMso(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ControlID As String)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(6)> _
    Sub ActivateTabQ(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ControlID As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal [Namespace] As String)
End Interface

<ComImport(), Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType(CShort(&H1040))> _
Public Interface IRibbonControl
    <DispId(1)> _
    ReadOnly Property Id() As <MarshalAs(UnmanagedType.BStr)> String
    <DispId(2)> _
    ReadOnly Property Context() As <MarshalAs(UnmanagedType.IDispatch)> Object
    <DispId(3)> _
    ReadOnly Property Tag() As <MarshalAs(UnmanagedType.BStr)> String
End Interface

<ComImport(), TypeLibType(CShort(&H1040)), Guid("000C0396-0000-0000-C000-000000000046")> _
Public Interface IRibbonExtensibility
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(1)> _
    Function GetCustomUI(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal RibbonID As String) As <MarshalAs(UnmanagedType.BStr)> String
End Interface
