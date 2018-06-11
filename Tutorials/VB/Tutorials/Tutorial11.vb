Imports System.Linq
Imports System.Runtime.CompilerServices
Imports NetOffice.CollectionsGeneric
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.Extensions

Public Class Tutorial11
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' Best practice to write own IEnumerable<T> extensions.
        ' See sample extension at the end of these file.

        ' NetOffice spend some extensions on IEnumerable<T> you may know from Linq2Objects.
        ' These extensions take care to free unused/unwanted COM Proxies immediately.
        ' However, the extensions doesnt works like Linq which means calling the result
        ' execute the method on demand. Its just ordinary extensions.

        Dim application As New Excel.ApplicationClass()
        application.DisplayAlerts = False
        application.Workbooks.Add()

        ' Here we use "First()" and "FirstOrDefault()"
        Dim sheet As Excel.Worksheet = application.Workbooks.First().Sheets.FirstOrDefault(Function(e) e.Name = "Sheet1")
        If Not IsNothing(sheet) Then

            sheet.Cells(1, 1).Value = "Test123"
            sheet.Cells(5, 5).Value = "Test123"
            sheet.Cells(10, 10).Value = "Test123"

            ' iterate over 10x10 used range and return the 3 cells we filled
            ' Linq2Objects would create 101 new proxies(10x10 + enumerator) here without any chance to free them.
            ' In NetOffice exensions - you have just 4 new managed proxies.
            Dim ranges As IEnumerable(Of Excel.Range) = sheet.UsedRange.Where(Function(e) False = IsNothing(e.Value))

            ' doing the same here again with the tutorial sample extension (scroll down)
            ranges = sheet.UsedRange.AllCellsWithValues()

        End If

        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial11"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Extensions and IEnumerable(Of T)"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

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
            Return FormMain.DocumentationBase & "Tutorial11_EN_VB.html"
        End Get
    End Property

End Class

Public Module Tutorial11Sample

    ' -- Best Practice sample extension to create extensions for IEnumerable(Of T) in NetOffice
    '
    ' In order to prevent ambiguous conflicts 
    ' you need to target NetOffice.CollectionsGeneric.IEnumerableProvider(Of T)
    ' All collections in NetOffice implement these interface
    <Extension()>
    Public Function AllCellsWithValues(source As IEnumerableProvider(Of Excel.Range)) As IEnumerable(Of Excel.Range)

        Dim result As List(Of Excel.Range) = New List(Of Excel.Range)
        Dim enumerator As ICOMObject = source.GetComObjectEnumerator(Nothing)

        Try

            For Each item As Excel.Range In source.FetchVariantComObjectEnumerator(source, enumerator)

                If False = IsNothing(item.Value) Then
                    result.Add(item)
                Else
                    item.Dispose()
                End If

            Next

            AllCellsWithValues = result

        Catch

            Throw

        Finally

            If False = IsNothing(enumerator) AndAlso Not Object.ReferenceEquals(enumerator, source) Then
                enumerator.Dispose()
            End If

        End Try

    End Function

End Module