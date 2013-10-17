Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' draw back color and perform the BorderAround method
        workSheet.Range("$B2:$B5").Interior.Color = ToDouble(Color.DarkGreen)
        workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic)

        ' draw back color and border the range explicitly
        workSheet.Range("$D2:$D5").Interior.Color = ToDouble(Color.DarkGreen)
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlDouble
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Weight = 4
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Color = ToDouble(Color.Black)

        workSheet.Cells(1, 1).Value = "We have 2 simple shapes created."

        ' save the book
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example01{1}", _hostApplication.RootDirectory, fileExtension)
        workBook.SaveAs(workbookFile)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example01", "Beispiel01")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Background Colors and Borders for Cells", "Hintergrundfarben und Rahmen in Zellen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

#Region "Helper"

    ''' <summary>
    ''' Translate a color to double
    ''' </summary>
    ''' <param name="color">expression to convert</param>
    ''' <returns>color</returns>
    ''' <remarks></remarks>
    Private Function ToDouble(ByVal color As System.Drawing.Color) As Double

        Dim returnValue As UInteger = color.B
        returnValue = returnValue << 8
        returnValue += color.G
        returnValue = returnValue << 8
        returnValue += color.R
        Return returnValue

    End Function

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".xls" or ".xlsx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Excel.Application) As String

        Dim version As Double = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture)
        If (version >= 12.0) Then
            Return ".xlsx"
        Else
            Return ".xls"
        End If

    End Function

#End Region

End Class
