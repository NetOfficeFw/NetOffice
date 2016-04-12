Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Tools
Imports NetOffice.ExcelApi.Tools.Utils

Public Class Tutorial13
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' Any MS-Office application in NetOffice has a custom utils provider for common tasks
        ' Moreover its available as instance property in NetOffice.Tools.COMAddin
        ' If you have suggestions for the utils please feel free to contact the project
        ' This tutorial shows only few features in MS-Excel

        ' start excel and disable alerts
        Dim application As New Excel.Application()
        application.DisplayAlerts = False


        ' Create an instance of excel utils
        Dim utils As CommonUtils = New CommonUtils(application, GetType(Tutorial13).Assembly)


        ' the file part of the utils makes it easier to deal with file extensions depedent on the current version


        ' get default(xls or xlsx) , template with macros(xlt or xltm) - extension and build a valid file path
        Dim extensionNormal As String = utils.File.FileExtension(DocumentFormat.Normal)

        Dim extensionTemplateWithMacros As String = utils.File.FileExtension(DocumentFormat.TemplateMacros)
        Dim exampleFilePath As String = utils.File.Combine("C:\MyFiles", "MyWorkbook", DocumentFormat.Normal)


        ' the dialog part of the utils allows you to show default dialogs/messageboxes or you own dialogs

        ' dialogs want be suppressed by default if the office application is currently in automation or not visible
        ' you can also trigger the DialogShow and DialogShown event to observe dialog popups
        ' we disable any suppress behavior here
        utils.Dialog.SuppressOnAutomation = False
        utils.Dialog.SuppressOnHide = False


        ' show a simple message box. Have a look at the last argument. Its a default result and used if the messagebox is not shown.
        ' In this tutorial, excel is in automation and hidden. Remove one or both of the 2 code lines above and the message box is not shown.
        ' We got the default result in this case
        Dim userResult As DialogResult = utils.Dialog.ShowMessageBox("Hello World from NetOffice tutorial", "NO tutorial", MessageBoxButtons.YesNo, DialogResult.No)


        Application.Quit()
        Application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial13"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "NetOffice Utils", "NetOffice Utils")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub ChangeLanguage(ByVal lcid As Integer) Implements TutorialsBase.ITutorial.ChangeLanguage

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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial13_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial13_DE_VB")
        End Get
    End Property

#End Region

End Class
