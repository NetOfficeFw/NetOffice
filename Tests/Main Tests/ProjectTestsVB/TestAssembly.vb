Imports Tests.Core

Public Class TestAssembly
    Implements ITestAssembly

    Private _listPackages As List(Of ITestPackage)

    Public ReadOnly Property Language As String Implements Tests.Core.ITestAssembly.Language
        Get
            Return "VB"
        End Get
    End Property

    Public Function LoadTestPackages() As Tests.Core.ITestPackage() Implements Tests.Core.ITestAssembly.LoadTestPackages

        If IsNothing(_listPackages) Then

            _listPackages = New List(Of ITestPackage)
            _listPackages.Add(New Test01())
            _listPackages.Add(New Test02())

        End If

        Return _listPackages.ToArray()

    End Function

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestAssembly.OfficeProduct
        Get
            Return "Project"
        End Get
    End Property

End Class
