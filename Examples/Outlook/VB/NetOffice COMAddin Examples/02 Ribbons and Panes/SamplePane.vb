Imports NetOffice.OutlookApi.Tools
Imports System.Drawing

Public Class SamplePane
    Implements ITaskPane ' Not necessary to implement ITaskPane but its helpful

    Private Counter As PerformanceCounter

    Public Sub OnConnection(ByVal application As NetOffice.OutlookApi.Application, ByVal parentPane As NetOffice.OfficeApi._CustomTaskPane, ByVal customArguments() As Object) Implements NetOffice.OutlookApi.Tools.ITaskPane.OnConnection


    End Sub

    Public Sub OnDisconnection() Implements NetOffice.OutlookApi.Tools.ITaskPane.OnDisconnection

        UsageTimer.Enabled = False
        If Not IsNothing(Counter) Then
            Counter.Dispose()
            Counter = Nothing
        End If

    End Sub

    Public Sub OnVisibleStateChanged(ByVal visible As Boolean) Implements NetOffice.OutlookApi.Tools.ITaskPane.OnVisibleStateChanged

        ' Create the performance counter is expensive in performance
        ' To avoid slow down the Outlook startup sequence - we create them on demand when user want show the pane
        ' (Real world code want doing that async)
        ' Note: There is an issue here in Outlook 2007 which support taskpanes but no ribbons
        If (visible And IsNothing(Counter)) Then
            Counter = New PerformanceCounter("Process", "% Processor Time", "OUTLOOK")
            UsageTimer.Enabled = True
        ElseIf True = visible Then
            UsageTimer.Enabled = True
        ElseIf False = visible Then
            UsageTimer.Enabled = False
        End If

    End Sub

    Public Sub OnDockPositionChanged(ByVal position As NetOffice.OfficeApi.Enums.MsoCTPDockPosition) Implements NetOffice.OutlookApi.Tools.ITaskPane.OnDockPositionChanged


    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)

        MyBase.OnResize(e)
        UsageLabel.Location = New Point(
                                    (Width / 2 - UsageLabel.Width / 2),
                                    (Height / 2 - UsageLabel.Height / 2))

    End Sub

    Private Sub UsageTimer_Tick(sender As Object, e As EventArgs) Handles UsageTimer.Tick

        If Not IsNothing(Counter) Then

            Dim value As Single = Counter.NextValue()
            Dim barValue As Int32 = Convert.ToInt32(value)
            If (barValue < 0) Then barValue = 0
            If (barValue > 100) Then barValue = 100
            UsageLabel.Text = String.Format("{0} %", barValue)
            UsageBar.Value = barValue

        Else

            UsageLabel.Text = String.Empty

        End If

    End Sub

End Class