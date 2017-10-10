Imports ExampleBase
Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.WordApi.Tools.Contribution

''' <summary>
''' Create a document write text and save
''' </summary>
Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application
        wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        ' create a utils instance, not need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(wordApplication)

        ' add a new document
        Dim newDocument As Word.Document
        newDocument = wordApplication.Documents.Add()

        ' insert some text
        wordApplication.Selection.TypeText("This text is written by automation")

        wordApplication.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend)
        wordApplication.Selection.Font.Color = WdColor.wdColorSeaGreen
        wordApplication.Selection.Font.Bold = 1
        wordApplication.Selection.Font.Size = 18

        'save document
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example01", DocumentFormat.Normal)
        newDocument.SaveAs(documentFile)

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example01"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create a document write text and save"
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

End Class
