Imports System.Reflection
Imports Tests.Core
Imports System.Windows.Forms
Imports System.Drawing

Imports NetOffice
Imports System.Data.OleDb
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports NetOffice.AccessApi.Constants

Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

Imports DAO = NetOffice.DAOApi
Imports NetOffice.DAOApi.Enums
Imports NetOffice.DAOApi.Constants

Public Class Test04
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
            Return "Test04"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Access"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Access.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            Dim iconBitmap As New Bitmap(System.Reflection.Assembly.GetAssembly(Me.GetType()).GetManifestResourceStream("AccessTestsVB.Test04.bmp"))
            application = New NetOffice.AccessApi.Application()

            Dim commandBar As Office.CommandBar = Nothing
            Dim commandBarBtn As Office.CommandBarButton = Nothing

            ' add a commandbar popup
            Dim commandBarPopup As Office.CommandBarPopup = application.CommandBars("Menu Bar").Controls.Add(MsoControlType.msoControlPopup)
            commandBarPopup.Caption = "commandBarPopup"

            ' add a button to the popup
            commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton)
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
            commandBarBtn.Caption = "commandBarButton"
            Clipboard.SetDataObject(iconBitmap)
            commandBarBtn.PasteFace()
            Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
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

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".mdb" or ".accdb"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Access.Application) As String

        ' Access 2000 doesnt have the Version property(unfortunately)
        ' we check for support with the SupportEntity method, implemented by NetOffice
        If (Not application.EntityIsAvailable("Version")) Then
            Return ".mdb"
        End If

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".accdb"
        Else
            Return ".xls"
        End If

    End Function

    Private Sub commandBarBtn_Click(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        Ctrl.Dispose()

    End Sub

End Class
