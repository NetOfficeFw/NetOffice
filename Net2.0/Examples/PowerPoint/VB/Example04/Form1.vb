Imports LateBindingApi.Core
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start powerpoint and turn off msg boxes
        Dim powerApplication As New PowerPoint.Application()
        powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone

        ' add a new presentation with two new slides
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        Dim slide1 As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)
        Dim slide2 As PowerPoint.Slide = presentation.Slides.Add(2, PpSlideLayout.ppLayoutBlank)

        ' add shapes
        slide1.Shapes.AddShape(MsoAutoShapeType.msoShape4pointStar, 100, 100, 200, 200)
        slide2.Shapes.AddShape(MsoAutoShapeType.msoShapeDoubleWave, 200, 200, 200, 200)

        ' change blend animation
        slide1.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverDown
        slide1.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast

        slide2.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverLeftDown
        slide2.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast

        ' save the document 
        Dim fileExtension As String = GetDefaultExtension(powerApplication)
        Dim documentFile As String = String.Format("{0}\\Example02{1}", Environment.CurrentDirectory, fileExtension)
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
