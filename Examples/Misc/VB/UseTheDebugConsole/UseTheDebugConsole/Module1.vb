Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi

Module Module1

    Sub Main()

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
        '   ConsoleMode.MemoryList = keep all messages in memory. use DebugConsole.Messages and DebugConsole.ClearMessagesList() with these option
        '   ConsoleMode.LogFile    = writes all messages immediately to a file. you have to set DebugConsole.FileName before use

        Dim application As Excel.Application = Nothing
        Try

            ' Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize()

            ' create excel instance
            application = New Excel.Application()

            ' activate the DebugConsole. the default value is: ConsoleMode.None
            DebugConsole.Mode = ConsoleMode.MemoryList

            ' Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize()

            ' create excel instance
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False

            ' we open a non existing file to produce an error
            application.Workbooks.Open("z:\\NotExistingFile.exe")

        Catch ex As Exception

            Console.WriteLine("An error is occured. NetOffice DebugConsole content below:")

            For Each item As String In DebugConsole.Messages
                Console.WriteLine(item)
            Next  

            Console.ReadKey()

        Finally

            ' quit and dispose
            application.Quit()
            application.Dispose()

        End Try

    End Sub

End Module
