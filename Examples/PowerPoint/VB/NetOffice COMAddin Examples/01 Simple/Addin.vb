Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.PowerPointApi.Tools
'
'Minimum Addin Example
'
<COMAddin("PowerPoint01 Sample Addin VB4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("PowerPoint01AddinVB4.Connect"), Guid("A9FAD74D-BFEE-4C80-89E5-53690BBC7C81"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        Console.WriteLine("PowerPoint Version is {0}", Application.Version)

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class