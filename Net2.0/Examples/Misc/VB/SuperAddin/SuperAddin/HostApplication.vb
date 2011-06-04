Imports LateBindingApi.Core
Imports System.ComponentModel
Imports Extensibility

Public Class HostApplication
    Implements IDisposable

#Region "Fields"

    Private disposedValue As Boolean = False        ' To detect redundant calls
    Private _hostApp As COMObject

#End Region

#Region "Properties"

    Public ReadOnly Property Application() As COMObject
        Get
            Return _hostApp
        End Get
    End Property

    Public ReadOnly Property ComponentName() As String
        Get
            Return TypeDescriptor.GetComponentName(_hostApp.UnderlyingObject)
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _hostApp.UnderlyingTypeName
        End Get
    End Property

    Public ReadOnly Property Version() As String
        Get
            Return Invoker.PropertyGet(_hostApp.UnderlyingObject, "Version")
        End Get
    End Property

#End Region

#Region "Construction"

    Public Sub New(ByVal comProxy As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As Array)

        Dim typeComponent As String = System.ComponentModel.TypeDescriptor.GetComponentName(comProxy)
        Select Case typeComponent
            Case "Microsoft Excel"
                _hostApp = New NetOffice.ExcelApi.Application(Nothing, comProxy)
            Case "Excel"
                _hostApp = New NetOffice.ExcelApi.Application(Nothing, comProxy)


            Case "Microsoft Word"
                _hostApp = New NetOffice.WordApi.Application(Nothing, comProxy)
            Case "Word"
                _hostApp = New NetOffice.WordApi.Application(Nothing, comProxy)

            Case "Microsoft Outlook"
                _hostApp = New NetOffice.OutlookApi.Application(Nothing, comProxy)
            Case "Outlook"
                _hostApp = New NetOffice.OutlookApi.Application(Nothing, comProxy)

            Case "Microsoft PowerPoint"
                _hostApp = New NetOffice.PowerPointApi.Application(Nothing, comProxy)
            Case "PowerPoint"
                _hostApp = New NetOffice.PowerPointApi.Application(Nothing, comProxy)

            Case "Microsoft Access"
                _hostApp = New NetOffice.AccessApi.Application(Nothing, comProxy)
            Case "Access"
                _hostApp = New NetOffice.AccessApi.Application(Nothing, comProxy)
        End Select

    End Sub

#End Region

#Region " IDisposable Members "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub


    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
