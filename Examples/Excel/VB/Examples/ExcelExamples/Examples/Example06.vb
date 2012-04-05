Imports ExampleBase
Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.VBIDEApi.Enums

Public Class Example06
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example06", "Beispiel06")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Dialogs in Excel", "Dialoge in Excel")
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

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False
        excelApplication.Visible = True

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        'show selected window and display user clicks ok or cancel
        Dim returnValue As Boolean
        Dim radioSelectButton As RadioButton = GetSelectedRadioButton()

        Select Case radioSelectButton.Text

            Case "xlDialogAddinManager"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogAddinManager).Show()

            Case "xlDialogFont"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogFont).Show()

            Case "xlDialogEditColor"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogEditColor).Show()

            Case "xlDialogGallery3dBar"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogGallery3dBar).Show()

            Case "xlDialogSearch"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogSearch).Show()

            Case "xlDialogPrinterSetup"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogPrinterSetup).Show()

            Case "xlDialogFormatNumber"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogFormatNumber).Show()

            Case "xlDialogApplyStyle"

                returnValue = excelApplication.Dialogs(XlBuiltInDialog.xlDialogApplyStyle).Show()

        End Select

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        Dim message As String = String.Format("The dialog returns {0}.", returnValue)
        MessageBox.Show(Me, message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

#Region "Helper"

    Private Function GetSelectedRadioButton() As RadioButton

        Dim itemControl As Control
        For Each itemControl In panelSelection.Controls

            If (TypeName(itemControl) = "RadioButton") Then

                Dim radioSelectButton As RadioButton = itemControl
                If (radioSelectButton.Checked) Then
                    Return radioSelectButton
                End If
            End If

        Next itemControl

        Throw (New Exception("No Dialog selected."))

    End Function

#End Region

End Class
