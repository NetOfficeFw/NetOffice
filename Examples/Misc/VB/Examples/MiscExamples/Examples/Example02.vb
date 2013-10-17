Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

' in some situations of version independent developement, its necessary to check for
' the support of a specific entity at runtime. for this reason any object in NetOffice
' has the following method:
'
'  bool EntityIsAvailable(string name);
'  bool EntityIsAvailable(string name, SupportEntityType searchType);
'  
' this example shows you how to use them.
Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' create excel instance
        Dim application As New Excel.Application()

        ' ask the application object for Quit method support
        Dim supportQuitMethod As Boolean = application.EntityIsAvailable("Quit")

        ' ask the application object for Visible property support
        Dim supportVisbibleProperty As Boolean = application.EntityIsAvailable("Visible")

        'ask the application object for SmartArtColors property support (only available in Excel 2010)
        Dim supportSmartArtColorsProperty As Boolean = application.EntityIsAvailable("SmartArtColors")

        ' ask the application object for XYZ property or method support (not exists of course)
        Dim supportTestXYZProperty As Boolean = application.EntityIsAvailable("TestXYZ")

        ' print result
        Dim messageBoxContent As String = ""
        messageBoxContent += String.Format("Your installed Excel Version supports the Quit Method: {0}{1}", supportQuitMethod, Environment.NewLine)
        messageBoxContent += String.Format("Your installed Excel Version supports the Visible Property: {0}{1}", supportVisbibleProperty, Environment.NewLine)
        messageBoxContent += String.Format("Your installed Excel Version supports the SmartArtColors Property: {0}{1}", supportSmartArtColorsProperty, Environment.NewLine)
        messageBoxContent += String.Format("Your installed Excel Version supports the TestXYZ Property: {0}{1}", supportTestXYZProperty, Environment.NewLine)
        MessageBox.Show(messageBoxContent, "EntityIsAvailable Result", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' quit and dispose
        application.Quit()
        application.Dispose()


    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example02", "Beispiel02")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Check entity support at runtime", "Zur Laufzeit prüfen ob eine Methode oder Property unterstützt wird")
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
