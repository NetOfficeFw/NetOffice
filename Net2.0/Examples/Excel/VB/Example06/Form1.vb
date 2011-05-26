Imports System.Reflection

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        'dont show dialogs with an invisible excel
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

        Dim message As String = String.Format("The dialog returns {0}.", returnValue)
        MessageBox.Show(Me, message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

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
