Imports System.Data
Imports System.Data.OleDb

Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports NetOffice.AccessApi.Constants
Imports DAO = NetOffice.DAOApi
Imports NetOffice.DAOApi.Enums
Imports NetOffice.DAOApi.Constants

Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start access 
        Dim accessApplication As New Access.Application()

        ' create database name 
        Dim fileExtension As String = GetDefaultExtension(accessApplication)
        Dim documentFile As String = String.Format("{0}\\Example02{1}", _hostApplication.RootDirectory, fileExtension)

        ' delete old database if exists
        If (System.IO.File.Exists(documentFile)) Then
            System.IO.File.Delete(documentFile)
        End If

        ' create database 
        Dim newDatabase As DAO.Database = accessApplication.DBEngine.Workspaces(0).CreateDatabase(documentFile, LanguageConstants.dbLangGeneral)
        accessApplication.DBEngine.Workspaces(0).Close()

        ' setup database connection              'Provider=Microsoft.Jet.OLEDB.4.0;Data Source= < access2007
        Dim oleConnection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;Data Source=" + documentFile)
        oleConnection.Open()

        ' create table
        Dim oleCreateCommand As New OleDbCommand("CREATE TABLE NetOfficeTable(Column1 Text, Column2 Text)", oleConnection)
        oleCreateCommand.ExecuteReader().Close()

        ' write some data with plain sql & close
        For i As Integer = 0 To 1000

            Dim insertCommand As String = String.Format("INSERT INTO NetOfficeTable(Column1, Column2) VALUES(""{0}"", ""{1}"")", i, DateTime.Now.ToShortTimeString())
            Dim oleInsertCommand As New OleDbCommand(insertCommand, oleConnection)
            oleInsertCommand.ExecuteReader().Close()

        Next
        oleConnection.Close()

        ' close access and dispose reference
        accessApplication.Quit(AcQuitOption.acQuitSaveAll)
        accessApplication.Dispose()

        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example02", "Beispiel02")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create new table", "Eine Tabelle erstellen")
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
    ''' returns the valid file extension for the instance. for example ".mdb" or ".accdb"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Access.Application) As String

        ' Access 2000 doesnt have the Version property(unfortunately)
        ' we check for support with the SupportEntity method, implemented by NetOffice
        If (Not application.EntityIsAvailable("Version")) Then
            Return ".mdb"
        End If

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".accdb"
        Else
            Return ".mdb"
        End If

    End Function

#End Region

End Class
