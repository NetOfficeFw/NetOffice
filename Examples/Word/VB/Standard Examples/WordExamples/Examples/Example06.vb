Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Example06
    Implements IExample

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Dim _hostApplication As ExampleBase.IHost

    Public Sub New()

        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example06"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Using Events"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Me
        End Get
    End Property


#End Region

#Region "UI Trigger"

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application()
        wordApplication.DisplayAlerts = False

        ' we register some events. note: the event trigger was called from word, means an other Thread
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

#End Region

#Region "Word Trigger"

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

#End Region
   
End Class

