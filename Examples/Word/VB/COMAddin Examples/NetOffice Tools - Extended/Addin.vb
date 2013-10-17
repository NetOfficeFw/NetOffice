Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Imports NetOffice
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
Imports Access = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
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
<COMAddin("NetOfficeVB4 Extended Word Addin", "This Addin shows you the COMAddin  baseclass from the NetOffice Tools", 3)> _
<RegistryLocation(RegistrySaveLocation.CurrentUser)> _
<CustomUI("NetOfficeTools.ExtendedWordVB4.RibbonUI.xml")> _
<Guid("A7F57B00-924A-4C61-BFCD-7F1620645A65"), ProgId("ExtendedWordVB4.Addin"), Tweak(True)> _
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        ' we can add our own taskpanes here, if you dont want that then overwrite the CTPFactoryAvailable method
        ' show into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application
        TaskPanes.Add(GetType(SamplePane), "NetOffice Tools - Sample Pane")
        TaskPanes(0).DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
        TaskPanes(0).DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
        TaskPanes(0).Width = 150
        TaskPanes(0).Visible = True
        TaskPanes(0).Arguments = New Object() {Me}

    End Sub

    ' trigger the well known IExtensibility2 methods, this is very similar to VSTO
    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' you see the host application is accessible as property from the class instance
        ' the application property was disposed automaticly while shutdown
        Console.WriteLine("Host Application Version is:{0}", Me.Application.Version)

    End Sub

    ' Ribbon UI Trigger
    Public Sub OnAction(ByVal control As Office.IRibbonControl)
        Try

            Select Case control.Id
                Case "customButton1"
                    MessageBox.Show("This is the first sample button.", "ExtendedWordVB4.Addin")
                Case "customButton2"
                    MessageBox.Show("This is the second sample button.", "ExtendedWordVB4.Addin")
                Case Else
                    MessageBox.Show("Unkown Control Id: " + control.Id, "ExtendedWordVB4.Addin")

            End Select

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "ExtendedAccessVB4.Addin", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
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

    '// this error handler is used for IExtensibility2 methods (your code) and the COMAddin methods GetCustomUI and CTPFactoryAvailable
    Protected Overrides Sub OnError(ByVal methodKind As NetOffice.Tools.ErrorMethodKind, ByVal exception As System.Exception)

        MessageBox.Show("An error occured in " & methodKind.ToString())

    End Sub

    '// for example you have an security issues while register or something like that
    '// then you can implement a static errorhandler method.
    '// the first parameter shows you the error occurs in Register or Unregister
    '// the second parameter is the thrown exception. rethrow the exception in this method to signalize an error to the environment
    <RegisterErrorHandler()> _
    Public Shared Sub RegisterErrorHandler(ByVal methodKind As RegisterErrorMethodKind, ByVal exception As Exception)

        MessageBox.Show("An error occured in " & methodKind.ToString())

    End Sub

End Class
