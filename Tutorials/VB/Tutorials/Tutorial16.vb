Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.GlobalHelperModules.GlobalModule

Public Class Tutorial16
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this example demonstrate the global helper module(static class)
        ' the module is a vba compatibility workarround and contains static methods and properties from the coresponding Application class.

        ' start excel and add a new workbook
        Dim application As New Excel.Application()
        application.Visible = False
        application.DisplayAlerts = False
        application.Workbooks.Add()

        ' GlobalModule contains the well known globals and is located in NetOffice.ExcelApi.GlobalHelperModules
        ' In VB.NET you can do now: ActiveCell.Value = "ActiveCellValue" and this is helpful 
        ' to bring code from VBA to VB/NetOffice
        ActiveCell.Value = "ActiveCellValue"

        ' close and dispose excel
        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial16"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Globals in NetOffice"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return FormMain.DocumentationBase & "Tutorial16_EN_VB.html"
        End Get
    End Property

End Class