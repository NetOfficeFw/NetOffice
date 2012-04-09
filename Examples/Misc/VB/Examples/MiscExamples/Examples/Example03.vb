Imports ExampleBase
Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi

'  in some situations you want check for a running office application instance.
'  this example shows you how to use the Marshal.GetActiveObject method to get a running application and create a NetOffice wrapper instance.
'  for this example we use outlook. please note the Marshal.GetActiveObject method throws a COMException if no running instance available
Public Class Example03
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost
    
    Public Sub New()

        InitializeComponent()
 
    End Sub

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example03", "Beispiel03")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "How to access a running Outlook application", "Eine laufene Outlook Instanz automatisieren")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Me
        End Get
    End Property

#End Region

#Region "UI Trigger"

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        Try

            Dim application As Outlook.Application = Nothing
            Dim nativeProxy As Object = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application")

            application = New Outlook.Application(Nothing, nativeProxy)
            textBoxLog.Clear()
            textBoxLog.AppendText("we got running outlook instance" + vbNewLine)
            textBoxLog.AppendText("outlook version is " + application.Version)

            'instance was already running at start. we dispose references but not quit application
            application.Dispose()

        Catch ex As System.Runtime.InteropServices.COMException

            _hostApplication.ShowErrorDialog(Nothing, ex)

        End Try

    End Sub

#End Region

End Class
