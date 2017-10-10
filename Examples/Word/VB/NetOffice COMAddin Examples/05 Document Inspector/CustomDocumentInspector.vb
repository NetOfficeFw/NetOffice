Imports System.Net
Imports System.Collections.Generic
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.WordApi.Tools
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.WordApi
''' <summary>
''' Deobfuscate all bit.ly URL's in a document
''' 
''' Please note: it is a sample implementation,
''' real code want handle more url shorteners and is aware of cascading shorteners (and possible endless cascade).
''' Moreover professional addins handle remote web communication outside(exe/service) of MS-Office
''' to deal friendly with desktop firewalls.
''' </summary>
Public Class CustomDocumentInspector
    Inherits ToolsDocumentInspector

    Private Shared Bitly() As String = New String() {"http://bit.ly", "https://bit.ly"}
    Private Shared LinkEnds() As String = New String() {" ", "^t", "\r"}

    ''' <summary>
    ''' What we found in Inspect
    ''' </summary>
    Private InspectResult As New Dictionary(Of Integer, String)

    ''' <summary>
    ''' Short/Long Url Cache to avoid resolve the same short link twice
    ''' </summary>
    Private Cache As New Dictionary(Of String, String)

    Public Overrides ReadOnly Property Name() As String

        Get
            Name = "Deobfuscate bit.ly Short Links(VB4)"
        End Get

    End Property

    Public Overrides ReadOnly Property Description() As String

        Get
            Description = "Performs HTTP calls to resolve bit.ly url's."
        End Get

    End Property

    Public Overrides Sub Inspect(doc As Document, ByRef status As MsoDocInspectorStatus, ByRef result As String, ByRef action As String)

        InspectResult.Clear()
        Cache.Clear()
        Dim range As Word.Range = doc.Content
        Dim find As Word.Find = range.Find
        find.Forward = True
        find.Text = "http*"
        find.MatchWildcards = True

        If find.Execute() Then

            Dim start As Integer = range.Start
            Do While start > 0

                Dim text As String = String.Empty
                Dim character As Word.Range = range.Characters(1)
                Dim characterText As String = character.Text
                Dim isEndLink As Boolean = False
                For Each item As String In LinkEnds
                    If characterText = item Then
                        isEndLink = True
                        Exit For
                    End If
                Next

                If False = isEndLink Then
                    text = text & character.Text
                    character = character.Next()
                End If

                For Each item As String In Bitly
                    If text.StartsWith(item) Then
                        InspectResult.Add(start, text)
                        Exit For
                    End If
                Next

                If False = find.Execute() Then
                    Exit Do
                End If

                start = range.Start

            Loop

        End If

        If InspectResult.Count > 0 Then

            status = MsoDocInspectorStatus.msoDocInspectorStatusIssueFound
            result = String.Format("{0} link(s) found.", InspectResult.Count)
            action = "Deobfuscate Links."

        Else

            status = MsoDocInspectorStatus.msoDocInspectorStatusDocOk
            result = "No links found."
            action = "No links to change."

        End If

    End Sub

    Public Overrides Sub Fix(doc As Document, hwnd As Integer, ByRef status As MsoDocInspectorStatus, ByRef result As String)

        Dim range As Word.Range = doc.Content
        Dim find As Word.Find = range.Find

        Dim replacedLinks As Integer = 0
        For Each item As KeyValuePair(Of Integer, String) In InspectResult
            Dim uri As String = TryGetBitlyRedirectUrl(item.Value)
            If False = String.IsNullOrWhiteSpace(uri) Then
                If find.Execute(item.Value, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, uri) Then
                    replacedLinks = replacedLinks + 1
                End If
            End If
        Next

        If (replacedLinks = InspectResult.Count) Then

            status = MsoDocInspectorStatus.msoDocInspectorStatusDocOk
            result = "All links have been replaced."

        Else

            status = MsoDocInspectorStatus.msoDocInspectorStatusError
            result = "Unable to replace one or more link(s)."
        End If

    End Sub

    Private Function TryGetBitlyRedirectUrl(uri As String) As String

        If Cache.ContainsKey(uri) Then
            Return Cache(uri)
        End If

        Dim request As HttpWebRequest = WebRequest.Create(uri)
        request.Timeout = 5000
        request.Method = "HEAD"
        request.AllowAutoRedirect = False
        Dim response As HttpWebResponse
        Try
            response = request.GetResponse()
            If Not IsNothing(response) Then
                Dim result As String = response.GetResponseHeader("Location")
                If False = String.IsNullOrWhiteSpace(result) & True = result.EndsWith("/") Then
                    result = result.Substring(0, result.Length - 1)
                End If
                Cache.Add(uri, result)
                Return result
            Else
                Return Nothing
            End If
        Catch exception As WebException
            ' 404 - invalid bit.ly link or timeout because network issues
            Return Nothing
        Catch execption As Exception
            Return Nothing
        End Try

    End Function

End Class