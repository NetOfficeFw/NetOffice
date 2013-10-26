Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.VBIDEApi.Enums
Imports Tests.Core

Public Class Test05
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using a VBE Macros."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test05"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Word"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Word.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.WordApi.Application()
            application.DisplayAlerts = WdAlertLevel.wdAlertsNone

            ' add a new document
            Dim newDocument As Word.Document
            newDocument = application.Documents.Add()

            ' add new module and insert macro
            ' the option "Trust access to Visual Basic Project" must be set
            Dim newModule As NetOffice.VBIDEApi.CodeModule
            newModule = newDocument.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule

            ' set the modulename
            newModule.Name = "NetOfficeTestModule"

            ' add the macro
            Dim codeLines As String
            codeLines = String.Format("Sub NetOfficeTestMacro()" & Environment.NewLine & "   {0}" & Environment.NewLine & _
                                      "End Sub", "Selection.TypeText (""This text is written by a automatic created macro with NetOffice..."")")
            newModule.InsertLines(1, codeLines)

            'start the macro
            application.Run("NetOfficeTestModule!NetOfficeTestMacro")

            newDocument.Close(False)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit(WdSaveOptions.wdDoNotSaveChanges)
                application.Dispose()
            End If

        End Try

    End Function

End Class
