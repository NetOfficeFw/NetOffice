Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.OutlookApi.Tools
'
'Minimum Addin Example
'
<COMAddin("Outlook01 Sample Addin VB4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Outlook01AddinVB4.Connect"), Guid("CFBD53D0-6C6B-4310-A2B4-92FC72D34225"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        Console.WriteLine("Outlook Version is {0}", Application.Version)

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class