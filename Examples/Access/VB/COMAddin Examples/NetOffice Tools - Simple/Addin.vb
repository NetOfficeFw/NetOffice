Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.AccessApi.Tools
Imports NetOffice.AccessApi

'/*
'  * This project shows you the COMAddin base class from the NetOffice tools.
'  * Its designed to reduce infrastructure code from your own.
'  * You have to set some attributes and thats all. 
'  * You see also the host application is available as class instance property. no need for dispose here because the base class do this for you while shutdown.
'*/

<COMAddin("NetOfficeVB4 Sample Access Addin", "This Addin shows you the COMAddin base class from the NetOffice Tools", 3)> _
<Guid("BD432A44-89E6-492e-9CB7-35D57CBC2B64"), ProgId("SimpleAccessVB4.Addin"), Tweak(True)> _
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' get the host application version
        ' we check at runtime the property available because Access 2000 doesnt have the Version property
        If Application.EntityIsAvailable("Version", NetOffice.SupportEntityType.Property) Then
            Dim hostVersion As String = Me.Application.Version
            Console.WriteLine("Host Application Version is:{0}", hostVersion)
        End If

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class
