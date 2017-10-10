Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
'
'Minimum Addin Example
'
<COMAddin("Word01 Sample Addin VB4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Word01AddinVB4.Connect"), Guid("E419BD32-242F-4931-A8A2-35460B0535EB"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        Console.WriteLine("Word Version is {0}", Application.Version)

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection

    End Sub

End Class