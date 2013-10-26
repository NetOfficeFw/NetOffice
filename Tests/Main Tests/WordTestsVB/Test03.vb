Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test03
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using List templates."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test03"
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

            'create a new listtemplate
            Dim template As Word.ListTemplate
            template = newDocument.ListTemplates.Add(True, "NetOfficeListTemplate")

            'get the predefined listlevels (9)
            Dim levels As Word.ListLevels
            levels = template.ListLevels

            'customize the first level of the list
            levels(1).NumberFormat = "%1."

            'tab is used to change the level
            levels(1).TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            levels(1).NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            levels(1).NumberPosition = 0
            levels(1).Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            levels(1).TextPosition = application.CentimetersToPoints(0.63F)
            levels(1).TabPosition = application.CentimetersToPoints(0.63F)
            levels(1).ResetOnHigher = 0
            levels(1).StartAt = 1
            levels(1).LinkedStyle = ""
            levels(1).Font.Bold = 1

            'customize the second level of the list
            levels(2).NumberFormat = "%1.%2."

            'tab is used to change the level
            levels(2).TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            levels(2).NumberStyle = WdListNumberStyle.wdListNumberStyleArabic

            'we want the numbers to appear under the first letter of the higher level
            levels(2).NumberPosition = application.CentimetersToPoints(0.63F)
            levels(2).Alignment = WdListLevelAlignment.wdListLevelAlignLeft

            'and the text should indent a tab more on the right
            levels(2).TextPosition = application.CentimetersToPoints(1.4F)
            levels(2).TabPosition = application.CentimetersToPoints(1.4F)
            levels(2).ResetOnHigher = 0
            levels(2).StartAt = 1
            levels(2).LinkedStyle = ""
            levels(2).Font.Italic = 1

            'apply the defined listtemplate to the selection
            application.Selection.Range.ListFormat.ApplyListTemplate(template, False, _
                            WdListApplyTo.wdListApplyToWholeList, WdDefaultListBehavior.wdWord9ListBehavior)

            'create a list
            application.Selection.TypeText("Welcoming")
            application.Selection.TypeParagraph()

            application.Selection.TypeText("Introduction")
            application.Selection.TypeParagraph()

            application.Selection.TypeText("Presentation")
            application.Selection.TypeParagraph()

            'execute the indent so the second level gets activated
            application.Selection.Range.ListFormat.ListIndent()

            application.Selection.TypeText("Top 1")
            application.Selection.TypeParagraph()

            application.Selection.TypeText("Top 2")
            application.Selection.TypeParagraph()

            application.Selection.TypeText("Top 3")
            application.Selection.TypeParagraph()

            ' execute the outdent so the first level gets reactivated
            application.Selection.Range.ListFormat.ListOutdent()
            application.Selection.TypeText("Questions & Answers")

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
