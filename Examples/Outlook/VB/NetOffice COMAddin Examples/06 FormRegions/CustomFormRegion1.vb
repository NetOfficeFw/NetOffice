Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Tools
Imports MSForms = NetOffice.MSFormsApi
Imports NetOffice.OutlookApi.Tools.Contribution

Public Class CustomFormRegion1
    Inherits OpenFormRegion

    Private textBox1 As Outlook.OlkTextBox
    Private commandButton1 As Outlook.OlkCommandButton

    Public Sub New(formRegion As Outlook.FormRegion)

        MyBase.New(formRegion)

        Dim form As MSForms.UserForm = formRegion.Form
        textBox1 = form.Controls("TextBox1").To(Of Outlook.OlkTextBox)()
        commandButton1 = form.Controls("CommandButton1").To(Of Outlook.OlkCommandButton)()

        If Not IsNothing(commandButton1) Then
            Dim handler As Outlook.OlkCommandButton_ClickEventHandler = AddressOf Me.CommandButton1_ClickEvent
            AddHandler commandButton1.ClickEvent, handler
        End If

    End Sub

    Private Sub CommandButton1_ClickEvent()

        If Not IsNothing(textBox1) Then
            OutlookDialogUtils.ShowMessageBox(textBox1.Text, "Outlook06AddinVB4")
        End If

    End Sub

End Class