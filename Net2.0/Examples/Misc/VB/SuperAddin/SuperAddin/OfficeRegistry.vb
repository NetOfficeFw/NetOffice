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
    ''' <param name="name"></param>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Shared Sub CreateAddinKey(ByVal name As String, ByVal key As String)

        Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey(key)
        regKey.Close()
        regKey = Registry.CurrentUser.OpenSubKey(key, True)
        regKey.SetValue("LoadBehavior", Convert.ToInt32(3))
        regKey.SetValue("FriendlyName", name)
        regKey.SetValue("Description", "example for versionindependent addin loaded in all office products")

    End Sub

    ''' <summary>
    ''' deletes addin key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteAddinKey(ByVal key As String)

        Registry.CurrentUser.DeleteSubKey(key)

    End Sub

End Class
