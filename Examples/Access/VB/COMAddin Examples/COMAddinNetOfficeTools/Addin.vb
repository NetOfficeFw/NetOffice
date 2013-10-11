Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports NetOffice
Imports NetOffice.Tools
Imports NetOffice.AccessApi.Tools
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

'/*
'  * This project shows you the COMAddin base class from the NetOffice tools.
'  * Its designed to reduce infrastructure code from your own.
'  * this addin looks a bit strange of course because the explanation
'  * check the NetOffice download section for NetOffice Tools based Addins
'  * Wikipedia Addin  - Word
'  * Twitter Addin    - Outlook
'  * Google Addin     - Excel
'*/

'
' as you can see, the needed registry informations was given as annotation, no need for Register/Unregister methods
' CurrentUser is default, no need for this attribute if you want HKEY_CURRENTUSER (just for example)
' you can specify a path to an embedded xml ressource file with your ribbon schema, otherwise you can override the GetCustomUI method from COMAddin base class
<COMAddin("NetOfficeVB4 Sample Addin", "This Addin shows you the COMAddin class from the NetOffice Tools", 3)> _
<RegistryLocation(RegistrySaveLocation.CurrentUser)> _
<CustomUI("COMAddinNetOfficeTools.RibbonUI.xml")> _
<GuidAttribute("C5ED5AD1-D1C2-45b8-B836-0F3D966D063F"), ProgIdAttribute("COMAddinNetOfficeToolsVB4.AccessSampleAddin")> _
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        'wen can add our own taskpanes here, if you dont want that then overwrite the CTPFactoryAvailable method
        ' show into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application
        TaskPanes.Add(GetType(SamplePane), "NetOffice Tools - Sample Pane")
        TaskPanes(0).DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
        TaskPanes(0).DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
        TaskPanes(0).Width = 150
        TaskPanes(0).Visible = True
        'TaskPanes(0).Arguments = new object[] { this }

    End Sub

    ' trigger the well known IExtensibility2 methods, this is very similar to VSTO
    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' you see the host application is accessible as property from the class instance
        ' the application property was disposed automaticly while shutdown
        MessageBox.Show("Version is:" + Me.Application.Version)

    End Sub


    '/*
    ' * now you see the way to exend or modify the register/unregister process if you want
    ' * we define 2 static methods with the RegisterFunction attribute, we use CallBeforeAndAfter as parameter
    ' * this means the register method in the base class call our method as first (before registry modification) and as last(after registry modification) 
    ' * the register call parameter give you the info what is is. Replace means the method in the base class does nothing and its your task to create the registry keys
    ' * same thing with Unregister method. 
    ' */
    <RegisterFunction(RegisterMode.CallBeforeAndAfter)> _
    Public Shared Sub Register(ByVal type As Type, ByVal registerCall As RegisterCall)

        Select Case registerCall

            Case registerCall.CallBefore

            Case registerCall.CallAfter

            Case registerCall.Replace


        End Select

    End Sub

    <RegisterFunction(RegisterMode.CallBeforeAndAfter)> _
    Public Shared Sub UnRegister(ByVal type As Type, ByVal registerCall As RegisterCall)

        Select Case registerCall

            Case registerCall.CallBefore

            Case registerCall.CallAfter

            Case registerCall.Replace


        End Select

    End Sub

    '/*
    '* at last you see some options for troubleshooting. the COMAddin base class is not a blackbox.
    '*/

    '// for example you have an security issues while register or something like that
    '// then you can implement a static errorhandler method.
    '// the first parameter shows you the error occurs in Register or Unregister
    '// the second parameter is the thrown exception. rethrow the exception in this method to signalize an error to the environment
    <RegisterErrorHandler()> _
    Public Shared Sub RegisterErrorHandler(ByVal methodKind As RegisterErrorMethodKind, ByVal exception As Exception)

    End Sub


    ' this non-static error handler is used for IExtensibility2 methods (your code) and the COMAddin methods GetCustomUI and CTPFactoryAvailable        
    <ErrorHandler()> _
    Public Shared Sub GeneralErrorHandler(ByVal methodKind As ErrorMethodKind, ByVal exception As Exception)

    End Sub

End Class
