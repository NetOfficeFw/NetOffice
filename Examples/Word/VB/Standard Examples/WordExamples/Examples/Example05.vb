Imports ExampleBase
Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports VB = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums

Public Class Example05
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application
        wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        ' create a utils instance, not need for but helpful to keep the lines of code low
        Dim utils As Word.Tools.CommonUtils = New Word.Tools.CommonUtils(wordApplication)

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
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example05", Word.Tools.DocumentFormat.Macros)
        If (utils.ApplicationIs2007OrHigher) Then
            newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault)
        Else
            newDocument.SaveAs(documentFile)
        End If

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example05", "Beispiel05")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create vba macros. The option Trust access to Visual Basic Project must be set.", "Erstellen von VBA Macros. Die Option Visual Basic Projekten vertrauen muss aktiviert sein.")
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
