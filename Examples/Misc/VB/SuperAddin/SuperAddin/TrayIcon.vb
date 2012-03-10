
Public Class TrayIcon
    Implements IDisposable

    Private disposedValue As Boolean = False        ' To detect redundant calls
    Private _trayIcon As NotifyIcon

    Public Sub New(ByVal visible As Boolean)

        _trayIcon = New NotifyIcon(New System.ComponentModel.Container())
        Dim iconStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("SuperAddinVB.AddinIcon.ico")
        _trayIcon.Icon = New System.Drawing.Icon(iconStream)
        iconStream.Close()
        _trayIcon.Text = "SuperAdddin loaded."
        _trayIcon.Visible = visible

    End Sub

    Public Property Visible() As Boolean

        Get
            Return _trayIcon.Visible
        End Get
        Set(ByVal value As Boolean)
            _trayIcon.Visible = value
        End Set

    End Property


    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            If (Not IsNothing(_trayIcon)) Then
                _trayIcon.Dispose()
                _trayIcon = Nothing
            End If

        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
