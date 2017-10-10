Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Tools
Imports NetOffice.ExcelApi.Tools.Attributes
'
'Sample RTD Component
'
<COMRtdServer("A sample rtd server", 1)>
<ProgId("MyRtdServerVB4.Server"), Guid("BD578911-742E-4853-B793-084B64A4CB09"), Codebase, Programmable>
Public Class Server
    Inherits RealtimeDataServer

    Dim TopicID As Integer

    Protected Overrides Function ConnectData(topicID As Integer, strings As Object, getNewValues As Boolean) As Object

        topicID = topicID
        ConnectData = GetTime()

    End Function

    Protected Overrides Function RefreshData(topicCount As Integer) As Object

        Dim data(,) As Object = New Object(2, 1) {}
        data(0, 0) = TopicID
        data(1, 0) = GetTime()
        topicCount = 1
        RefreshData = data

    End Function

    Private Function GetTime() As String

        GetTime = "GetTime " & DateTime.Now.ToString()

    End Function

End Class