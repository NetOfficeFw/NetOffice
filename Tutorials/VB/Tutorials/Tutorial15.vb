Imports System.Runtime.InteropServices
Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial15
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this example demonstrate the NetOffice low-level interface for latebinding calls

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook
        Dim book As Excel.Workbook = application.Workbooks.Add()

        Dim sheet As Excel.Worksheet = application.Workbooks(1).Worksheets(1)
        Dim sampleRange As Excel.Range = sheet.Cells(1, 1)

        'we set the COMVariant ColorIndex from Font of ouer sample range with the invoker class
        Invoker.Default.PropertySet(sampleRange.Font, "ColorIndex", 1)

        ' creates a native unmanaged ComProxy with the invoker
        Dim comProxy As Object = Invoker.Default.PropertyGet(application, "Workbooks")
        Marshal.ReleaseComObject(comProxy)

        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial15"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Using the Invoker"
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
            Return FormMain.DocumentationBase & "Tutorial15_EN_VB.html"
        End Get
    End Property

End Class
