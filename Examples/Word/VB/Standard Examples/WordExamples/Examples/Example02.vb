Imports ExampleBase
Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.WordApi.Tools.Contribution

''' <summary>
''' Example 2 - Insert a table to document
''' </summary>
Public Class Example02
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

        ' add a table
        Dim table As Word.Table
        table = newDocument.Tables.Add(wordApplication.Selection.Range, 3, 2)

        'insert some text into the cells
        table.Cell(1, 1).Select()
        wordApplication.Selection.TypeText("This")

        table.Cell(1, 2).Select()
        wordApplication.Selection.TypeText("table")

        table.Cell(2, 1).Select()
        wordApplication.Selection.TypeText("was")

        table.Cell(2, 2).Select()
        wordApplication.Selection.TypeText("created")

        table.Cell(3, 1).Select()
        wordApplication.Selection.TypeText("by")

        table.Cell(3, 2).Select()
        wordApplication.Selection.TypeText("NetOffice")

        'save document
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example02", DocumentFormat.Normal)
        newDocument.SaveAs(documentFile)

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example02"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Insert a table to document"
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
