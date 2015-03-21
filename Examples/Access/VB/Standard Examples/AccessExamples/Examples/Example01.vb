Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports NetOffice.AccessApi.Constants
Imports DAO = NetOffice.DAOApi
Imports NetOffice.DAOApi.Enums
Imports NetOffice.DAOApi.Constants

Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start access 
        Dim accessApplication As New Access.Application()

        ' create a utils instance, not need for but helpful to keep the lines of code low
        Dim utils As Access.Tools.CommonUtils = New Access.Tools.CommonUtils(accessApplication)

        ' create database file name 
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example01", Access.Tools.DocumentFormat.Normal)

        ' delete old database if exists
        If (System.IO.File.Exists(documentFile)) Then
            System.IO.File.Delete(documentFile)
        End If

        ' create database 
        Dim newDatabase As DAO.Database = accessApplication.DBEngine.Workspaces(0).CreateDatabase(documentFile, LanguageConstants.dbLangGeneral)

        ' close access and dispose reference
        accessApplication.Quit(AcQuitOption.acQuitSaveAll)
        accessApplication.Dispose()

        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example01", "Beispiel01")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create new Database", "Eine neue Datenbank erstellen")
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

End Class
