Imports System.Windows.Forms
Imports System.Drawing
Imports Excel = NetOffice.ExcelApi
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports Tests.Core

Public Class Test08
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Create custom UI."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test08"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Excel"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Excel.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            Dim iconBitmap As New Bitmap(System.Reflection.Assembly.GetAssembly(Me.GetType()).GetManifestResourceStream("ExcelTestsVB.Test08.bmp"))
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False
            application.Workbooks.Add()
 
            ' add a commandbar popup
            Dim commandBarPopup As Office.CommandBarPopup = application.CommandBars("Worksheet Menu Bar").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarPopup.Caption = "commandBarPopup"

            ' add a button to the popup
            Dim commandBarBtn As Office.CommandBarButton = commandBarPopup.Controls.Add(MsoControlType.msoControlButton)
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
            commandBarBtn.Caption = "commandBarButton"
            Clipboard.SetDataObject(iconBitmap)
            commandBarBtn.PasteFace()
            Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
            AddHandler commandBarBtn.ClickEvent, clickHandler

            'add a new toolbar
            Dim commandBar As Office.CommandBar = application.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, System.Type.Missing, True)
            commandBar.Visible = True

            ' add a button to the toolbar
            commandBarBtn = commandBar.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
            commandBarBtn.Caption = "commandBarButton"
            commandBarBtn.FaceId = 3
            clickHandler = AddressOf Me.commandBarBtn_Click
            AddHandler commandBarBtn.ClickEvent, clickHandler

            ' add a dropdown box to the toolbar
            commandBarPopup = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarPopup.Caption = "commandBarPopup"

            ' add a button to the popup, we use an own icon for the button
            commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
            commandBarBtn.Caption = "commandBarButton"
            Clipboard.SetDataObject(iconBitmap)
            commandBarBtn.PasteFace()
            clickHandler = AddressOf Me.commandBarBtn_Click
            AddHandler commandBarBtn.ClickEvent, clickHandler

            ' create context menu
            commandBarPopup = application.CommandBars("Cell").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarPopup.Caption = "commandBarPopup"

            ' add a button to the popup
            commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
            commandBarBtn.Caption = "commandBarButton"
            commandBarBtn.FaceId = 9
            clickHandler = AddressOf Me.commandBarBtn_Click
            AddHandler commandBarBtn.ClickEvent, clickHandler

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function

    Private Sub commandBarBtn_Click(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        Ctrl.Dispose()

    End Sub
End Class
