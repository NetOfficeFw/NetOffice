Imports ExampleBase
Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums
Imports VB = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums

Public Class Example03
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        Dim powerApplication As PowerPoint.Application = Nothing
        Dim documentFile As String = Nothing
        Dim isFailed = False

        Try

            ' start powerpoint
            powerApplication = New PowerPoint.Application()

            ' add a new presentation with one new slide
            Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
            Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

            ' add new module and insert macro. the option "Trust access to Visual Basic Project" must be set
            Dim vbeModule As VB.CodeModule = presentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule

            Dim macro As String = String.Format("Sub NetOfficeTestMacro()" & vbNewLine & "   {0}" & vbNewLine & "End Sub", "MsgBox ""Thanks for click!""")
            vbeModule.InsertLines(1, macro)

            ' add button and connect with macro
            Dim button As PowerPoint.Shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeActionButtonForwardorNext, 100, 100, 200, 200)
            button.ActionSettings(PpMouseActivation.ppMouseClick).AnimateAction = MsoTriState.msoTrue
            button.ActionSettings(PpMouseActivation.ppMouseClick).Action = PpActionType.ppActionRunMacro
            button.ActionSettings(PpMouseActivation.ppMouseClick).Run = "NetOfficeTestMacro"

            ' save the document 
            Dim fileExtension As String = GetDefaultExtension(powerApplication)
            documentFile = String.Format("{0}\\Example03{1}", _hostApplication.RootDirectory, fileExtension)
            presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue)

        Catch throwedException As Exception

            isFailed = True
            _hostApplication.ShowErrorDialog("VBA Error", throwedException)

        Finally

            ' close excel and dispose reference
            powerApplication.Quit()
            powerApplication.Dispose()

            If (Not IsNothing(documentFile) And Not isFailed) Then
                ' show dialog for the user(you!)
                _hostApplication.ShowFinishDialog(Nothing, documentFile)
            End If

        End Try

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example03", "Beispiel03")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create an run macros. the option 'Trust access to Visual Basic Project' must be set", "Makros erstellen und ausführen. Die Option 'Visual Basic Projekten vertrauen' muss aktiviert sein.")
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

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".ppt" or ".pptm"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As PowerPoint.Application) As String

        Dim version As Double = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture)
        If (version >= 12.0) Then
            Return ".pptm"
        Else
            Return ".ppt"
        End If

    End Function

#End Region

End Class