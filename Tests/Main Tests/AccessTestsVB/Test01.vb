Imports NetOffice
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports Tests.Core

Imports DAO = NetOffice.DAOApi
Imports NetOffice.DAOApi.Enums
Imports NetOffice.DAOApi.Constants

Public Class Test01
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Insert text."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test01"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Access"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Access.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.AccessApi.Application()

            ' create database name 
            Dim fileExtension As String = GetDefaultExtension(application)
            Dim documentFile As String = String.Format("{0}\\Test01{1}", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), fileExtension)

            ' delete old database if exists
            If (System.IO.File.Exists(documentFile)) Then
                System.IO.File.Delete(documentFile)
            End If

            ' create database 
            Dim newDatabase As DAO.Database = application.DBEngine.Workspaces(0).CreateDatabase(documentFile, LanguageConstants.dbLangGeneral)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function

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
            Return ".xls"
        End If

    End Function

End Class
