Imports System.Runtime.InteropServices
Imports Word06AddinVB4

<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual), Guid("CD610128-44B7-4AFD-AF34-F8AD9FF1BD8E")>
Public Interface ITimeComponent

    <DispId(1)>
    Function GetTime() As String

End Interface

Public Class TimeComponent
    Implements ITimeComponent

    Public Function GetTime() As String Implements ITimeComponent.GetTime
        Return DateTime.Now.ToString()
    End Function

End Class
