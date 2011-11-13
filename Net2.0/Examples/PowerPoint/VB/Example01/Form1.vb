Imports LateBindingApi.Core
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums


Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start powerpoint
        Dim powerApplication As PowerPoint.Application = New PowerPoint.Application()

        ' add a new presentation with one new slide
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText)

        ' save the document 
        Dim fileExtension As String = GetDefaultExtension(powerApplication)
        Dim documentFile As String = String.Format("{0}\\Example01{1}", Application.StartupPath, fileExtension)
        presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        Dim fDialog As New FinishDialog("Presentation saved.", documentFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".ppt" or ".pptx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As PowerPoint.Application) As String

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".pptx"
        Else
            Return ".ppt"
        End If

    End Function

#End Region

End Class
