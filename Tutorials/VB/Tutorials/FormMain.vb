Public Class FormMain

    Public Shared ReadOnly DocumentationBase As String = "https://netoffice.io/documentation/"

    Public Sub New()

        InitializeComponent()
        Me.Text = "NetOffice Tutorials in VB"
        LoadTutorials()

    End Sub

    Private Sub LoadTutorials()

        LoadTutorial(New Tutorial01())
        LoadTutorial(New Tutorial02())
        LoadTutorial(New Tutorial03())
        LoadTutorial(New Tutorial04())
        LoadTutorial(New Tutorial05())
        LoadTutorial(New Tutorial06())
        LoadTutorial(New Tutorial07())
        LoadTutorial(New Tutorial08())
        LoadTutorial(New Tutorial09())
        LoadTutorial(New Tutorial10())
        LoadTutorial(New Tutorial11())
        LoadTutorial(New Tutorial12())
        LoadTutorial(New Tutorial13())
        LoadTutorial(New Tutorial14())
        LoadTutorial(New Tutorial15())
        LoadTutorial(New Tutorial16())
        LoadTutorial(New Tutorial17())

    End Sub

End Class
