Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.AccessApi.Tools
'
'Minimum Addin Example
'
<COMAddin("Access01 Sample Addin VB4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Access01AddinVB4.Connect"), Guid("B5CBBECD-4DEE-4A61-AD34-E9B8618452B0"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        If (Application.EntityIsAvailable("Version")) Then
            Console.WriteLine("Access Version is {0}", Application.Version)
        Else
            Console.WriteLine("Access Version is {0}", "9(2000) or below")
        End If

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class