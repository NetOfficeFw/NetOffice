Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport(), Guid("000C033E-0000-0000-C000-000000000046"), TypeLibType(CShort(&H10C0))> _
Public Interface ICustomTaskPaneConsumer
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(1)> _
    Sub CTPFactoryAvailable(<[In](), MarshalAs(UnmanagedType.Interface)> ByVal CTPFactoryInst As Object)
End Interface
