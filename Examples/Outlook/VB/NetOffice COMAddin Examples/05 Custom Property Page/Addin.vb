Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Tools
'
'Custom Property Page Addin Example
'
<COMAddin("Outlook05 Sample Addin VB4", "Custom Property Page Example", LoadBehavior.LoadAtStartup)>
<ProgId("Outlook05AddinVB4.Connect"), Guid("DAFB59F0-F7AB-49CA-ADC0-1C2755C95E03"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As Array) Handles Me.OnStartupComplete

        Dim handler As Outlook.Application_OptionsPagesAddEventHandler = AddressOf Me.Application_OptionsPagesAddEvent
        AddHandler Application.OptionsPagesAddEvent, handler

        ' This is another way to bring a custom option page to
        ' the Mail Folder TreeView on the left - context menu/properties
        ' --------------------------------------------------------------
        ' Dim mapi As Outlook.NameSpace = Application.GetNamespace("MAPI")
        ' Dim mapiHandler As Outlook.NameSpace_OptionsPagesAddEventHandler = AddressOf Me.Mapi_OptionsPagesAddEvent
        ' AddHandler mapi.OptionsPagesAddEvent, mapiHandler

    End Sub

    Private Sub Application_OptionsPagesAddEvent(pages As Outlook.PropertyPages)

        'we show the NetOffice Core settings in the option page
        pages.Add(New OptionPage(Factory), "Outlook05 Sample Addin VB4")

    End Sub


    Private Sub Mapi_OptionsPagesAddEvent(pages As Outlook.PropertyPages, folder As Outlook.MAPIFolder)

    End Sub

End Class