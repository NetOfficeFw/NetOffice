Imports LateBindingApi.Core
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start powerpoint and turn off msg boxes
        Dim powerApplication As New PowerPoint.Application()
        powerApplication.Visible = MsoTriState.msoTrue

        ' PowerPoint 2000 doesnt support DisplayAlerts, we check at runtime its available and set
        If (powerApplication.EntityIsAvailable("DisplayAlerts")) Then
            powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone
        End If

        ' we register some events. note: the event trigger was called from power point, means an other Thread
        ' remove the Quit() call below and check out more events if you want

        Dim newCloseHandler As PowerPoint.Application_PresentationCloseEventHandler = AddressOf Me.powerApplication_PresentationCloseEvent
        AddHandler powerApplication.PresentationCloseEvent, newCloseHandler

        Dim newAfterNewHandler As PowerPoint.Application_AfterNewPresentationEventHandler = AddressOf Me.powerApplication_AfterNewPresentationEvent
        AddHandler powerApplication.AfterNewPresentationEvent, newAfterNewHandler

        ' add a new presentation with one new slide
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

        ' close the document
        presentation.Close()

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

    End Sub


    Private Sub powerApplication_PresentationCloseEvent(ByVal Pres As NetOffice.PowerPointApi.Presentation)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event PresentationClose called."})
        Pres.Dispose()

    End Sub


    Private Sub powerApplication_AfterNewPresentationEvent(ByVal Pres As NetOffice.PowerPointApi.Presentation)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event AfterNewPresentation called."})
        Pres.Dispose()

    End Sub


    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

End Class
