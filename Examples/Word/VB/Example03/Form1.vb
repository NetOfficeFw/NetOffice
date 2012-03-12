Imports System.Reflection

Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.VBIDEApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' initialize api
        LateBindingApi.Core.Factory.Initialize()

        'start word and turn off msg boxes
        Dim wordApplication As Word.Application
        wordApplication = New Word.Application()
        wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        'add a new document
        Dim newDocument As Word.Document
        newDocument = wordApplication.Documents.Add()

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
        levels(1).TextPosition = wordApplication.CentimetersToPoints(0.63F)
        levels(1).TabPosition = wordApplication.CentimetersToPoints(0.63F)
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
        levels(2).NumberPosition = wordApplication.CentimetersToPoints(0.63F)
        levels(2).Alignment = WdListLevelAlignment.wdListLevelAlignLeft

        'and the text should indent a tab more on the right
        levels(2).TextPosition = wordApplication.CentimetersToPoints(1.4F)
        levels(2).TabPosition = wordApplication.CentimetersToPoints(1.4F)
        levels(2).ResetOnHigher = 0
        levels(2).StartAt = 1
        levels(2).LinkedStyle = ""
        levels(2).Font.Italic = 1

        'apply the defined listtemplate to the selection
        wordApplication.Selection.Range.ListFormat.ApplyListTemplate(template, False, _
                        WdListApplyTo.wdListApplyToWholeList, WdDefaultListBehavior.wdWord9ListBehavior)

        'create a list
        wordApplication.Selection.TypeText("Welcoming")
        wordApplication.Selection.TypeParagraph()

        wordApplication.Selection.TypeText("Introduction")
        wordApplication.Selection.TypeParagraph()

        wordApplication.Selection.TypeText("Presentation")
        wordApplication.Selection.TypeParagraph()

        'execute the indent so the second level gets activated
        wordApplication.Selection.Range.ListFormat.ListIndent()

        wordApplication.Selection.TypeText("Top 1")
        wordApplication.Selection.TypeParagraph()

        wordApplication.Selection.TypeText("Top 2")
        wordApplication.Selection.TypeParagraph()

        wordApplication.Selection.TypeText("Top 3")
        wordApplication.Selection.TypeParagraph()

        ' execute the outdent so the first level gets reactivated
        wordApplication.Selection.Range.ListFormat.ListOutdent()
        wordApplication.Selection.TypeText("Questions & Answers")

        ' save the document
        Dim fileExtension As String = GetDefaultExtension(wordApplication)
        Dim documentFile As String = String.Format("{0}\Example03{1}", Application.StartupPath, fileExtension)

        ' newer word versions try to save the document in .odt(open document format) by default
        ' we dont want this, we want .doc or .docx !!!
        Dim version As Double = Convert.ToDouble(wordApplication.Version)
        If (version >= 120.0) Then
            newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault)
        Else
            newDocument.SaveAs(documentFile)
        End If

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        Dim fDialog As New FinishDialog("Document saved.", documentFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".doc" or ".docx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Word.Application) As String

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".docx"
        Else
            Return ".doc"
        End If

    End Function

#End Region

End Class
