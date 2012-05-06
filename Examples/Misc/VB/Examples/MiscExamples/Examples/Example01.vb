Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

'  NetOffice gives you as additional service a debug console.
'  Essential NetOffice system steps and any occured exception related with
'  your office application(or NetOffice himself maybe) are stored here. if an error occureed and you need help,
'  please use the NetOffice discussion board: http://netoffice.codeplex.com/discussions
'  describe your problem and post the content of the DebugConsole as below your message.
'  the following infos are also helpful: operating system 32 or 64 bit, office version 32 or 64 bit, assembly runs as administrator or not
'
'  the following options are available
' 
'   ConsoleMode.None       = Console is deactivated (default)
'   ConsoleMode.Console    = redirect all messages to System.Console
'   ConsoleMode.Trace      = redirect all messages to System.Diagnostics.Trace
'   ConsoleMode.MemoryList = keep all messages in memory. use DebugConsole.Messages and DebugConsole.ClearMessagesList() with these option
'   ConsoleMode.LogFile    = writes all messages immediately to a file. you have to set DebugConsole.FileName before use

Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        Dim application As Excel.Application = Nothing
        Try

            ' create excel instance
            application = New Excel.Application()

            ' activate the DebugConsole. the default value is: ConsoleMode.None
            DebugConsole.Mode = ConsoleMode.MemoryList

            ' create excel instance
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False

            ' we open a non existing file to produce an error
            application.Workbooks.Open("z:\\NotExistingFile.exe")

        Catch

            Dim messages As String = ""

            For Each item As String In DebugConsole.Messages
                messages += item + Environment.NewLine
            Next

            MessageBox.Show(messages, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            ' quit and dispose
            application.Quit()
            application.Dispose()

        End Try

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example01", "Beispiel01")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Use the NetOffice Debug Console", "Benutzen der NetOffice Debug Console")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

End Class
