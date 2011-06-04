Imports System.Reflection

Public NotInheritable Class FormShowError

#Region "Fields"

    Private _errorHeader As String
    Private _errorFooter As String
    Private _exception As Exception
    Private _isExtended As Boolean

#End Region

#Region "Construction"

    Public Sub New(ByVal exceptionToShow As Exception)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Initialize("An error is occured.", "", exceptionToShow)
        labelErrorHeader.Visible = True

    End Sub

    Private Sub Initialize(ByVal errorHeader As String, ByVal errorFooter As String, ByVal exceptionToShow As Exception)

        Me.Width = 440
        Me.Height = 160
        _isExtended = False

        _errorHeader = errorHeader
        _errorFooter = errorFooter

        labelErrorHeader.Text = errorHeader
        labelErrorFooter.Text = errorFooter

        _exception = exceptionToShow
        Dim i As Integer = 1
        Do While IsNothing(_exception)

            Dim lviException As ListViewItem = listViewExceptions.Items.Insert(0, i.ToString())
            lviException.SubItems.Add(_exception.Source)
            lviException.SubItems.Add(exceptionToShow.GetType().ToString())
            lviException.SubItems.Add(_exception.Message)
            _exception = _exception.InnerException
            i = i + 1

        Loop

        _exception = exceptionToShow

    End Sub

#End Region

#Region "GuiTrigger"

    Private Sub listViewExceptions_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listViewExceptions.Resize

        listViewExceptions.Columns(0).Width = (listViewExceptions.Width / 100) * 10
        listViewExceptions.Columns(1).Width = (listViewExceptions.Width / 100) * 20
        listViewExceptions.Columns(2).Width = (listViewExceptions.Width / 100) * 20
        listViewExceptions.Columns(3).Width = (listViewExceptions.Width / 100) * 60

    End Sub

    Private Sub buttonOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonOk.Click

        Me.Close()

    End Sub

    Private Sub buttonDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonDetails.Click

        If (True = _isExtended) Then
            Me.Width = 440
            Me.Height = 160
            _isExtended = False
        Else
            Me.Width = 440
            Me.Height = 300
            _isExtended = True
        End If

    End Sub

#End Region



    ''' <summary>
    ''' writes message to logfile in dll folder
    ''' </summary>
    ''' <param name="header"></param>
    ''' <param name="throwedException"></param>
    ''' <remarks></remarks>
    Public Shared Sub LogError(ByVal header As String, ByVal throwedException As Exception)

        'dll path
        Dim codeBase As String = Assembly.GetCallingAssembly().CodeBase

        If (True = codeBase.StartsWith("file:///", StringComparison.InvariantCultureIgnoreCase)) Then
            codeBase = codeBase.Substring(8)
        End If

        codeBase = codeBase.Replace("/", "\")
        codeBase = codeBase.Substring(0, codeBase.LastIndexOf("\"))

        'write message
        Dim logPath As String = System.IO.Path.Combine(codeBase, "SuperAddinErrors.log")
        Dim message As String = ""

        If (Not IsNothing(throwedException)) Then
            message = throwedException.Message + vbNewLine
        End If

        System.IO.File.AppendAllText(logPath, header + "\r\n" + message)

    End Sub


End Class
