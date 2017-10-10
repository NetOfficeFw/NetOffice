Imports Microsoft.Win32

''' <summary>
''' office addin keys and create/delete methods
''' </summary>
''' <remarks></remarks>
Public Class OfficeRegistry

    '
    ' office addin registry keys 
    '

    Public Shared Office As String = "Software\\Microsoft\\Office\\"
    Public Shared Excel As String = "Software\\Microsoft\\Office\\Excel\\AddIns\\"
    Public Shared Word As String = "Software\\Microsoft\\Office\\Word\\AddIns\\"
    Public Shared Outlook As String = "Software\\Microsoft\\Office\\Outlook\\AddIns\\"
    Public Shared PowerPoint As String = "Software\\Microsoft\\Office\\PowerPoint\\AddIns\\"
    Public Shared Access As String = "Software\\Microsoft\\Office\\Access\\AddIns\\"

    ''' <summary>
    ''' creates addin key
    ''' </summary>
    ''' <param name="officeApp"></param>
    ''' <param name="progId"></param>
    ''' <param name="name"></param>
    ''' <param name="description"></param>
    ''' <remarks></remarks>
    Public Shared Sub CreateAddinKey(ByVal officeApp As String, ByVal progId As String, ByVal name As String, ByVal description As String)

        Dim regKey As String = GetRegistryKey(officeApp)

        Dim rk As RegistryKey = Registry.CurrentUser.CreateSubKey(regKey + progId)
        rk.Close()
        rk = Registry.CurrentUser.OpenSubKey(regKey + progId, True)
        rk.SetValue("LoadBehavior", Convert.ToInt32(3))
        rk.SetValue("FriendlyName", name)
        rk.SetValue("Description", description)

    End Sub

    ''' <summary>
    ''' deletes addin key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteAddinKey(ByVal key As String)

        Registry.CurrentUser.DeleteSubKey(key, False)

    End Sub

    Public Shared Sub LogErrorMessage(ByVal officeApp As String, ByVal progId As String, ByVal message As String, ByVal exception As Exception)

        Dim regKey As String = GetRegistryKey(officeApp)
        Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(regKey + progId, True)

        rk.SetValue("ErrorTimestamp", DateTime.Now.ToString())
        rk.SetValue("ErrorMessage", message)
        rk.SetValue("ErrorException", exception.Message)
        rk.Close()

    End Sub

    Private Shared Function GetRegistryKey(ByVal officeApp As String) As String

        Select Case officeApp
            Case "Excel"
                Return Excel
            Case "Word"
                Return Word
            Case "Outlook"
                Return Outlook
            Case "PowerPoint"
                Return PowerPoint
            Case "Access"
                Return Access
            Case Else
                Throw New ArgumentOutOfRangeException("officeApp")
        End Select

    End Function

End Class
