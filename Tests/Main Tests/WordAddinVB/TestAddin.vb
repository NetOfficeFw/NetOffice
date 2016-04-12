Imports System.Reflection
Imports System.Windows.Forms
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports NetOffice
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
Imports Access = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<COMAddin("NOTestsMain.WordTestAddinVB", "This is a test addin from NOTests.Main", 3)> _
<CustomUI("WordAddinVB.RibbonUI.xml"), RegistryLocation(RegistrySaveLocation.LocalMachine)> _
<Guid("D101EE86-9BAA-41D8-A9CE-093687BD2E46"), ProgId("NOTestsMain.WordTestAddinVB"), Tweak(True)> _
Public Class TestAddin
    Inherits COMAddin

    Public Sub New()

        TaskPanes.Add(GetType(SampleControl), "NOTestsMain - VB Word Pane")
        TaskPanes(0).DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
        TaskPanes(0).DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
        TaskPanes(0).Width = 150
        TaskPanes(0).Visible = True
        TaskPanes(0).Arguments = New Object() {Me}

    End Sub

    Public ReadOnly Property StatusOkay As String

        Get

            If (True = RibbonUIOkay And True = TaskPaneOkay And True = TweakOkay And IsNothing(GeneralError)) Then
                Return True
            Else
                Return False
            End If

        End Get

    End Property

    Public ReadOnly Property StatusDescription
        Get
            Dim result As String = ""

            If (False = TaskPaneOkay) Then
                result += "Taskpane is not loaded"
            End If
            If (False = RibbonUIOkay) Then
                result += "RibbonUI is not loaded"
            End If
            If (False = TweakOkay) Then
                result += "Tweak is not set " + NetOffice.Settings.Default.ExceptionMessage
            End If
            If (Not IsNothing(GeneralError)) Then
                result += "General Error:" + GeneralError
            End If

            Return result

        End Get
    End Property

    Private ReadOnly Property TweakOkay As Boolean
        Get
            If (Factory.Settings.ExceptionMessage.StartsWith("WordTweakVB")) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Private ReadOnly Property RibbonUIOkay As Boolean
        Get
            Return Not IsNothing(RibbonUI)
        End Get
    End Property

    Private GeneralError As String

    Private RibbonUI As Office.IRibbonUI

    Friend TaskPaneOkay As Boolean


    Public Sub OnLoadRibbonUI(ByVal ribbUI As Office.IRibbonUI)

        RibbonUI = ribbUI

    End Sub

    Public Function GetLabel(ByVal control As Office.IRibbonControl)

        Return Factory.Settings.ExceptionMessage

    End Function

    Private Sub Addin_OnConnection(ByVal Application As Object, ByVal ConnectMode As NetOffice.Tools.ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Handles Me.OnConnection

        Dim addin As New Office.COMAddIn(Nothing, AddInInst)
        addin.Object = Me
        addin.Dispose()

    End Sub

    Protected Overrides Function AllowApplyTweak(ByVal name As String, ByVal value As String) As Boolean

        Factory.Console.SendPipeConsoleMessage("WordTestAddinVB", String.Format("AllowApplyTweak {0}:{1}", name, value))
        Return True

    End Function

    Protected Overrides Sub OnError(ByVal methodKind As NetOffice.Tools.ErrorMethodKind, ByVal exception As System.Exception)

        If (IsNothing(GeneralError)) Then
            GeneralError = ""
        End If

        GeneralError += methodKind.ToString() + Environment.NewLine + exception.GetType().Name + Environment.NewLine + exception.Message

    End Sub

    <RegisterFunction(RegisterMode.CallAfter)> _
    Public Shared Sub Register(ByVal type As Type, ByVal registerCall As RegisterCall)

        SetTweakPersistenceEntry(type, "NOExceptionMessage", "WordTweakVB", False)

    End Sub

End Class
