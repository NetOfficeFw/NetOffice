Public Class Tutorial11
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        Dim message As String = IIf(_hostApplication.LCID = 1033, "This tutorial doens't contain example code", "Dieses Tutorial enthält keinen Beispielcode")
        MessageBox.Show(message, "Tutorial11", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial11"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Addin Deployment", "Addins auf anderen System installieren")
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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial11_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial11_DE_VB")
        End Get
    End Property

#End Region

End Class