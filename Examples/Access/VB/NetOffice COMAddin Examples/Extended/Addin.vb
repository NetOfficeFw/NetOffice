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
' * This project shows you in depth the COMAddin base class from the NetOffice tools.
' * The COMAddin base class is designed to reduce infrastructure code.
' * This addin looks a bit strange of course because the explanation.
' * Check the NetOffice download section for more NetOffice Tools based Addins
' * Wikipedia Addin  - Word
' * Twitter Addin    - Outlook
' * Google Addin     - Excel
'*/

'
'/*
'* As you can see, the necessary registry informations was given as annotation, no need for Register/Unregister methods
'* The RegistryLocation attribute is not always necessary. CurrentUser is default, no need for this attribute if you want HKEY_CURRENTUSER (just for example here)
'* You see also the CustomUI attribute. You can specify a path to an embedded xml ressource file with your ribbon schema. If you want this then you can override the GetCustomUI method from the base class.
'* The Tweak attribute allows to set various NetOffice options at runtime with custom values entries in the current office addin key(helpful for troubleshooting). Learn more about in the Tweaks sample addin project.
'*/
<COMAddin("NetOfficeVB4 Extended Access Addin", "This Addin shows you the COMAddin  baseclass from the NetOffice Tools", 3)> _
<CustomUI("NetOfficeTools.ExtendedAccessVB4.RibbonUI.xml"), RegistryLocation(RegistrySaveLocation.CurrentUser)> _
<Guid("A3FF00C1-6894-46F8-A8BA-EEB863FBBBAF"), ProgId("ExtendedAccessVB4.Addin"), Tweak(True)> _
Public Class Addin
Inherits COMAddin

    Public Sub New()

        ' enable shared debug output and send a load message(use NOTools.ConsoleMonitor.exe to observe the shared console output)
        Factory.Console.Name = "ExtendedAccessVB4.Addin"
        Factory.Console.EnableSharedOutput = True
        Factory.Console.SendPipeConsoleMessage("ExtendedAccessVB4.Addin", "ExtendedAccessVB4.Addin has been loaded.")

        ' We add our own taskpane here, if you dont want this way then overwrite the CTPFactoryAvailable method and create your panes in this method.
        ' Taskpanes in Netoffice can implement the ITaskPane interface with the OnConnection/OnDisconnection to avoid the singleton pattern.
        ' Take a look into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application.
        TaskPanes.Add(GetType(SamplePane), "NetOffice Tools - Sample Pane(VB4)")
        TaskPanes(0).DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom
        TaskPanes(0).DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
        TaskPanes(0).Height = 50
        TaskPanes(0).Visible = True
        TaskPanes(0).Arguments = New Object() {Me}
        Dim handler As Office.CustomTaskPane_VisibleStateChangeEventHandler = AddressOf Me.TaskPane_VisibleStateChange
        AddHandler TaskPanes(0).VisibleStateChange, handler

    End Sub

    ' ouer ribbon instance
    Private RibbonUI As Office.IRibbonUI

    ' say hello in console at startup
    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' You see the host application is accessible as property from the class instance.
        ' The application property was disposed automaticly while shutdown.
        '  We check at runtime (with a NetOffice special service) the property is available because Access 2000 and below doesn't have the Version property.
        If Application.EntityIsAvailable("Version", NetOffice.SupportEntityType.Property) Then
            Factory.Console.WriteLine("Host Application Version is:{0}", Me.Application.Version)
        Else
            Factory.Console.WriteLine("Host Application Version 2000(9) or below.")
        End If

    End Sub

    ' trigger taskpane visibility has been changed and update the checkbutton in the ribbon ui for show/hide taskpane
    Private Sub TaskPane_VisibleStateChange(ByVal CustomTaskPaneInst As Office._CustomTaskPane)

        ' ouer taskpane visibility has been changed. we send a message to the host application
        ' and say please refresh the checkbutton state. now the host application want call ouer OnGetPressedPanelToggle method to update the checkstate.
        If Not IsNothing(RibbonUI) Then
            RibbonUI.InvalidateControl("paneVisibleToogleButton")
        End If

    End Sub

    ' defined in RibbonUI.xml to get a instance for ouer ribbon ui.
    Public Sub OnLoadRibonUI(ByVal ribbUI As Office.IRibbonUI)

        RibbonUI = ribbUI

    End Sub


    '  defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
    Public Function OnGetPressedPanelToggle(ByVal control As Office.IRibbonControl) As Boolean

        Return TaskPanes(0).Visible

    End Function


    ' defined in RibbonUI.xml to track the user clicked ouer checkbutton. then we upate the panel visibility at hand.
    Public Sub OnCheckPanelToggle(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)

        TaskPanes(0).Visible = pressed

    End Sub

    ' defined in RibbonUI.xml to catch the user click for the about button
    Public Sub OnClickAboutButton(ByVal control As Office.IRibbonControl)

        MessageBox.Show("NetOffice Tools - Extended Sample Addin.", "ExtendedAccessVB4.Addin")

    End Sub

    '/*
    '* Now you see the way to exend or modify the register/unregister process if you want.
    '* We define 2 static methods with the RegisterFunction attribute, we use CallBeforeAndAfter as condition.
    '* This condition means the register method in the base class call our method as first (before registry modification) and as last(after registry modification).
    '* The register call argument give you the info what is is currently. Replace means the method in the base class does nothing and its your task to create the registry keys.
    '* Same thing with Unregister method. 
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

    ' This error handler is used for IExtensibility2 events (your code) and the COMAddin methods GetCustomUI, CTPFactoryAvailable and CreateFactory(also overwrites).
    ' the first argument shows in which method the error is occured. The second argument is the detailed exception info. 
    ' Rethrow the exception otherwise the exception is marked as handled.
    Protected Overrides Sub OnError(ByVal methodKind As NetOffice.Tools.ErrorMethodKind, ByVal exception As System.Exception)

        MessageBox.Show("An error occured in " & methodKind.ToString(), "ExtendedAccessVB4.Addin")

    End Sub

    ' This method demonstrate an error handle for the register/unregister process.
    ' For example you have an security issues while register or something like that then you can implement a static errorhandler method.
    ' The first argument shows you the error occurs in Register or Unregister.
    ' The second argument is the thrown exception. Rethrow the exception to signalize an error to the environment otherwise the exception is handled.
    <RegisterErrorHandler()> _
    Public Shared Sub RegisterErrorHandler(ByVal methodKind As RegisterErrorMethodKind, ByVal exception As Exception)

        MessageBox.Show("An error occured in " & methodKind.ToString(), "ExtendedAccessVB4.Addin")

    End Sub

End Class
