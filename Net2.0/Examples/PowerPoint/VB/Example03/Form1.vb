Imports LateBindingApi.Core
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports VBE = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        Dim powerApplication As PowerPoint.Application = Nothing
        Dim documentFile As String = Nothing
        Try
            ' Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize()

            ' start powerpoint
            powerApplication = New PowerPoint.Application()

            ' add a new presentation with one new slide
            Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
            Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

            ' add new module and insert macro
            ' the option "Trust access to Visual Basic Project" must be set
            Dim vbeModule As VBE.CodeModule = presentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule

            Dim macro As String = String.Format("Sub NetOfficeTestMacro()" & vbNewLine & "   {0}" & vbNewLine & "End Sub", "MsgBox ""Thanks for click!""")
            vbeModule.InsertLines(1, macro)

            ' add button and connect with macro
            Dim button As PowerPoint.Shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeActionButtonForwardorNext, 100, 100, 200, 200)
            button.ActionSettings(PpMouseActivation.ppMouseClick).AnimateAction = MsoTriState.msoTrue
            button.ActionSettings(PpMouseActivation.ppMouseClick).Action = PpActionType.ppActionRunMacro
            button.ActionSettings(PpMouseActivation.ppMouseClick).Run = "NetOfficeTestMacro"

            ' save the document 
            Dim fileExtension As String = GetDefaultExtension(powerApplication)
            documentFile = String.Format("{0}\\Example03{1}", Application.StartupPath, fileExtension)
            presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue)

        Catch throwedException As Exception

            ' not trusted
            Dim message As String = String.Format("An error is occured.{0}ExceptionTrace:{0}", Environment.NewLine)

            Dim exception As Exception = throwedException
            While (Not IsNothing(exception))
                message += String.Format("{0}{1}", exception.Message, Environment.NewLine)
                exception = exception.InnerException
            End While

            MessageBox.Show(message)

        Finally

            ' close excel and dispose reference
            powerApplication.Quit()
            powerApplication.Dispose()

            If (Not IsNothing(documentFile)) Then
                Dim fDialog As FinishDialog = New FinishDialog("Presentation saved.", documentFile)
                fDialog.ShowDialog(Me)
            End If

        End Try

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
