Imports System.ComponentModel
Imports NetOffice
Imports NetOffice.OutlookApi.Native
Imports Outlook = NetOffice.OutlookApi

Public Class OptionPage
    Implements Outlook.Native.PropertyPage

    Private DataSource As Settings
    Private EditSource As Settings
    Private PageContainer As Outlook.Native.PropertyPageSite

    Public Sub New(core As Core)

        InitializeComponent()
        DataSource = core.Settings
        EditSource = New Settings(core.Settings)
        SettingsGrid.SelectedObject = EditSource
        Dim handler As PropertyChangedEventHandler = AddressOf Me.EditSource_PropertyChanged
        AddHandler EditSource.PropertyChanged, handler

    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)

        MyBase.OnLoad(e)
        PageContainer = Outlook.Tools.Contribution.ApplicationUtils.TryGetPageContainer(Me)

    End Sub

    Public ReadOnly Property Dirty As Boolean Implements PropertyPage.Dirty

        Get
            Return False = DataSource.IsEqualTo(EditSource)
        End Get

    End Property

    Public Sub Apply() Implements PropertyPage.Apply

        If Dirty Then
            DataSource.LoadFrom(EditSource)
        End If

    End Sub

    Public Sub GetPageInfo(ByRef HelpFile As String, ByRef HelpContext As Integer) Implements PropertyPage.GetPageInfo


    End Sub

    Private Sub EditSource_PropertyChanged(sender As Object, args As PropertyChangedEventArgs)

        If Not IsNothing(PageContainer) Then
            PageContainer.OnStatusChange()
        End If

    End Sub

End Class