Imports ExampleBase
Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports VB = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums
Imports NetOffice.WordApi.Tools.Contribution

''' <summary>
''' Example 5 - Create vba macros
''' </summary>
Public Class Example05
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

        ' add new module and insert macro
        ' the option "Trust access to Visual Basic Project" must be set
        Dim newModule As VB.CodeModule
        newModule = newDocument.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule

        ' set the modulename
        newModule.Name = "NetOfficeTestModule"

        ' add the macro
        Dim codeLines As String
        codeLines = String.Format("Sub NetOfficeTestMacro()" & Environment.NewLine & "   {0}" & Environment.NewLine & _
                                  "End Sub", "Selection.TypeText (""This text is written by a automatic created macro with NetOffice..."")")
        newModule.InsertLines(1, codeLines)

        'start the macro
        wordApplication.Run("NetOfficeTestModule!NetOfficeTestMacro")

        ' save the document
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example05", DocumentFormat.Macros)
        If (utils.ApplicationIs2007OrHigher) Then
            newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault)
        Else
            newDocument.SaveAs(documentFile)
        End If

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example05"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create vba macros. The option Trust access to Visual Basic Project must be set."
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
