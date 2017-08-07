Imports NetOffice

Module Module1

    Sub Main()

        Dim excelType As Type = System.Type.GetTypeFromProgID("Excel.Application", True)
        Dim interopProxy As Object = Activator.CreateInstance(excelType)

        NetOffice.Settings.Default.EnableAutomaticQuit = True

        Dim application As Object = New COMDynamicObject(interopProxy)
        application.Visible = True
        application.Workbooks.Add()
        application.Workbooks.Add()
        application.Workbooks.Add()

        Dim book1Active = application.ActiveWorkbook = application.Workbooks(1)
        Dim book3Active = application.ActiveWorkbook = application.Workbooks(3)

        Console.WriteLine("book1Active {0} book3Active {1}", book1Active, book3Active)

        application.Dispose()

    End Sub

End Module
