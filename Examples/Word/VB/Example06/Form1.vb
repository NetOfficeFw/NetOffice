Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application()
        wordApplication.DisplayAlerts = False

        ' we register some events. note: the event trigger was called from word, means an other Thread
        ' remove the Quit() call below and check out more events if you want

        Dim newHandler As Word.Application_NewDocumentEventHandler = AddressOf Me.wordApplication_NewDocumentEvent
        AddHandler wordApplication.NewDocumentEvent, newHandler

        Dim newCloseHandler As Word.Application_DocumentBeforeCloseEventHandler = AddressOf Me.wordApplication_DocumentBeforeCloseEvent
        AddHandler wordApplication.DocumentBeforeCloseEvent, newCloseHandler

        ' add a document and close
        Dim document As Word.Document = wordApplication.Documents.Add()
        document.Close()

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

    End Sub

    Private Sub wordApplication_NewDocumentEvent(ByVal Doc As Word.Document)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event NewDocumentEvent called."})
        Doc.Dispose()

    End Sub


    Private Sub wordApplication_DocumentBeforeCloseEvent(ByVal Doc As Word.Document, ByRef Cancel As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event DocumentBeforeClose called."})
        Doc.Dispose()

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

End Class
