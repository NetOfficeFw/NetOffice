Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports Tests.Core

Public Class Test05
    Implements ITestPackage

    Dim _presentationClose As Boolean
    Dim _afterNewPresentation As Boolean

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using events."
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
            Return "PowerPoint"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As PowerPoint.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New PowerPoint.Application()
            application.Visible = MsoTriState.msoTrue

            ' PowerPoint 2000 doesnt support DisplayAlerts, we check at runtime its available and set
            If (application.EntityIsAvailable("DisplayAlerts")) Then
                application.DisplayAlerts = PpAlertLevel.ppAlertsNone
            End If

            ' we register some events. note: the event trigger was called from power point, means an other Thread
            ' remove the Quit() call below and check out more events if you want

            Dim newCloseHandler As PowerPoint.Application_PresentationCloseEventHandler = AddressOf Me.powerApplication_PresentationCloseEvent
            AddHandler application.PresentationCloseEvent, newCloseHandler

            Dim newAfterNewHandler As PowerPoint.Application_AfterNewPresentationEventHandler = AddressOf Me.powerApplication_AfterNewPresentationEvent
            AddHandler application.AfterNewPresentationEvent, newAfterNewHandler

            ' add a new presentation with one new slide
            Dim presentation As PowerPoint.Presentation = application.Presentations.Add(MsoTriState.msoTrue)
            Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

            System.Threading.Thread.Sleep(2000)

            ' close the document
            presentation.Close()

            If (_afterNewPresentation And _presentationClose) Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Return New TestResult(False, DateTime.Now.Subtract(startTime), String.Format("AfterNewPresentation:{0} , PresentationClose:{1}", _afterNewPresentation, _presentationClose), Nothing, "")
            End If

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function
     
    Private Sub powerApplication_PresentationCloseEvent(ByVal Pres As NetOffice.PowerPointApi.Presentation)

        _presentationClose = True
        Pres.Dispose()

    End Sub


    Private Sub powerApplication_AfterNewPresentationEvent(ByVal Pres As NetOffice.PowerPointApi.Presentation)

        _afterNewPresentation = True
        Pres.Dispose()

    End Sub

End Class
