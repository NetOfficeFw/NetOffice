Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial05
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()

        ' ActiveSheet is defined as unkown Proxy in Excel Type Library, it can have multiple times at runtime
        ' but its always a COM Proxy, never a scalar type like bool or int. 
        ' In VBA oder PIA its converted to object, in NetOffice its also represents as COMObject
        Dim sheet As Object = application.ActiveSheet
        If (TypeName(sheet) = "Worksheet") Then
            Dim activeSheet As Excel.Worksheet = sheet
        End If

        ' all classes inherites from the common base type COMObject
        ' you can use also:
        Dim anonymSheet As COMObject = application.ActiveSheet

        '3 basic properties of COMObject
        Dim proxy As Object = anonymSheet.UnderlyingObject ' the real COM proxy, be carefull !
        Dim proxyClassName As String = anonymSheet.UnderlyingTypeName ' the class name of the COM proxy, for example "Worksheet"
        Dim isDisposed As Boolean = anonymSheet.IsDisposed ' info about the object is already disposed

        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial05"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Understanding unkown Types", "Richtiges verwenden von unbekannten Typen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub ChangeLanguage(ByVal lcid As Integer) Implements TutorialsBase.ITutorial.ChangeLanguage

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Nothing
        End Get
    End Property


    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial05_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial05_DE_VB")
        End Get
    End Property

#End Region

End Class
