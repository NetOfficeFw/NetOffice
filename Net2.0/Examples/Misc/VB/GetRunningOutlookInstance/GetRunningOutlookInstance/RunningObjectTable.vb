Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.ComponentModel

''' <summary>
''' helper class for accessing Running Object Table 
''' </summary>
''' <remarks>taken(and modified) from http://dotnet-snippets.de/dns/laufende-com-objekte-abfragen-SID526.aspx</remarks>
Public Class RunningObjectTable

    ' Win32-API-call for reading ROT
    <DllImport("ole32.dll")> Public Shared Function GetRunningObjectTable(ByVal reserved As Short, ByRef pprot As IRunningObjectTable) As Integer
    End Function

    ' Win32-API-call to create binding
    <DllImport("ole32.dll")> Public Shared Function CreateBindCtx(ByVal reserved As Short, ByRef pctx As IBindCtx) As Integer
    End Function

    ''' <summary>
    ''' returns native com proxy from outlook application object in Running Object Table
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function GetRunningOutlookInstanceFromROT() As Object

        Dim monikerList As IEnumMoniker = Nothing
        Dim runningObjectTable As IRunningObjectTable = Nothing

        Try

            ' query table and returns null if no objects runnings
            If (Not GetRunningObjectTable(0, runningObjectTable) = 0 And Not IsNothing(runningObjectTable)) Then
                Return Nothing
            End If

            ' query moniker & reset 
            runningObjectTable.EnumRunning(monikerList)
            monikerList.Reset()

            Dim monikerContainer(1) As IMoniker
            Dim pointerFetchedMonikers As IntPtr = IntPtr.Zero

            ' fetch all moniker
            Do While (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) = 0)

                ' create binding object
                Dim bindInfo As IBindCtx = Nothing
                CreateBindCtx(0, bindInfo)

                ' query com proxy info       
                Dim comInstance As Object = Nothing
                runningObjectTable.GetObject(monikerContainer(0), comInstance)

                Dim name As String = TypeDescriptor.GetClassName(comInstance)
                Dim component As String = TypeDescriptor.GetComponentName(comInstance, False)
                If ((component = "Outlook") And (name = "Application")) Then
                    Marshal.ReleaseComObject(bindInfo)
                    Return comInstance
                Else
                    Marshal.ReleaseComObject(comInstance)
                End If

                Marshal.ReleaseComObject(bindInfo)

            Loop

            ' outlook is not running 
            Return Nothing

        Finally

            'release proxies
            If (Not IsNothing(runningObjectTable)) Then
                Marshal.ReleaseComObject(runningObjectTable)
            End If

            If (Not IsNothing(monikerList)) Then
                Marshal.ReleaseComObject(monikerList)
            End If

        End Try

    End Function

End Class
