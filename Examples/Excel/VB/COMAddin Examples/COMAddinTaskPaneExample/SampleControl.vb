Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports Office = NetOffice.OfficeApi

Public Class SampleControl

    Dim _customers As List(Of Customer)

    Public Sub New()

        InitializeComponent()
        LoadSampleCustomerData()
        UpdateSearchResult()

    End Sub

#Region "Private Methods"

    Private Sub LoadSampleCustomerData()

        _customers = New List(Of Customer)

        Dim embeddedCustomerXmlContent As String = ReadString("CustomerData.xml")
        Dim document As New XmlDocument
        document.LoadXml(embeddedCustomerXmlContent)
        For Each customerNode As XmlNode In document.DocumentElement.ChildNodes

            Dim id As Integer = Convert.ToInt32(customerNode.Attributes("ID").Value)
            Dim name As String = customerNode.Attributes("Name").Value
            Dim company As String = customerNode.Attributes("Company").Value
            Dim city As String = customerNode.Attributes("City").Value
            Dim postalCode As String = customerNode.Attributes("PostalCode").Value
            Dim country As String = customerNode.Attributes("Country").Value
            Dim phone As String = customerNode.Attributes("Phone").Value

            _customers.Add(New Customer(id, name, company, city, postalCode, country, phone))

        Next

    End Sub

    Private Sub UpdateSearchResult()

        listViewSearchResults.Items.Clear()
        For Each item As Customer In _customers

            If (item.Name.IndexOf(textBoxSearch.Text.Trim(), StringComparison.InvariantCultureIgnoreCase) > -1) Then

                Dim viewItem As ListViewItem = listViewSearchResults.Items.Add("")
                viewItem.SubItems.Add(item.ID.ToString())
                viewItem.SubItems.Add(item.Name)
                viewItem.ImageIndex = 0
                viewItem.Tag = item

            End If

        Next

    End Sub

    Private Sub UpdateDetails()

        If (listViewSearchResults.SelectedItems.Count > 0) Then

            Dim selectedCustomer As Customer = listViewSearchResults.SelectedItems(0).Tag
            propertyGridDetails.SelectedObject = selectedCustomer

        Else

            propertyGridDetails.SelectedObject = Nothing

        End If

    End Sub

    Private Function ToRangeAddress(ByVal rowIndex As Integer, ByVal columnIndex As Integer)

        If (columnIndex < 1) Then Throw (New ArgumentOutOfRangeException("Invalid Argument. columnIndex must be > 0"))
        If (rowIndex < 1) Then Throw (New ArgumentOutOfRangeException("Invalid Argument. rowIndex must be > 0"))

        Dim columnChars() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

        If (columnIndex <= columnChars.Length) Then Return columnChars(columnIndex - 1) + rowIndex.ToString()

        Dim multi As Integer = columnIndex / columnChars.Length
        Dim pre As String = columnChars(multi - 1)
        Dim newx As Integer = columnIndex
        newx -= (multi * columnChars.Length)
        Return pre + columnChars(newx - 1) + rowIndex.ToString()

    End Function

    Private Function CalculateRangeArea(ByVal rowIndex As Integer, ByVal columnIndex As Integer, ByVal countOfProperties As Integer) As String

        Dim startRangeAddress As String = ToRangeAddress(rowIndex, columnIndex)
        Dim endEndRangeAddress As String = ToRangeAddress(rowIndex + countOfProperties - 1, columnIndex + 1)
        Return startRangeAddress + ":" + endEndRangeAddress

    End Function

    Private Function ToStringArray(ByVal customer As Customer) As Object(,)

        Dim customerPropertiesArray(7, 2) As Object

        customerPropertiesArray(0, 0) = "ID:"
        customerPropertiesArray(0, 1) = customer.ID.ToString()

        customerPropertiesArray(1, 0) = "Name:"
        customerPropertiesArray(1, 1) = customer.Name

        customerPropertiesArray(2, 0) = "Company:"
        customerPropertiesArray(2, 1) = customer.Company

        customerPropertiesArray(3, 0) = "City:"
        customerPropertiesArray(3, 1) = customer.City

        customerPropertiesArray(4, 0) = "Postal Code:"
        customerPropertiesArray(4, 1) = customer.PostalCode

        customerPropertiesArray(5, 0) = "Country:"
        customerPropertiesArray(5, 1) = customer.Country

        customerPropertiesArray(6, 0) = "Phone:"
        customerPropertiesArray(6, 1) = customer.Phone

        Return customerPropertiesArray

    End Function

    ''' <summary>
    ''' reads text from ressource
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReadString(ByVal fileName As String) As String

        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim ressourceStream As System.IO.Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + fileName)
        If (IsNothing(ressourceStream)) Then
            Throw (New System.IO.IOException("Error accessing resource Stream."))
        End If

        Dim textStreamReader As System.IO.StreamReader = New System.IO.StreamReader(ressourceStream)
        If (IsNothing(textStreamReader)) Then
            Throw (New System.IO.IOException("Error accessing resource File."))
        End If

        Dim text As String = textStreamReader.ReadToEnd()
        ressourceStream.Close()
        textStreamReader.Close()
        Return text

    End Function

#End Region

#Region "UI Trigger"

    Private Sub listViewSearchResults_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listViewSearchResults.DoubleClick

        If (listViewSearchResults.SelectedItems.Count > 0) Then

            Dim activeSheet As Excel.Worksheet = Addin.Application.ActiveSheet
            Dim activeCell As Excel.Range = Addin.Application.ActiveCell

            If Not IsNothing(activeCell) Then

                Dim rowIndex As Integer = activeCell.Row
                Dim columnIndex As Integer = activeCell.Column

                Dim targetRangeAddress As String = CalculateRangeArea(rowIndex, columnIndex, 7)

                Dim selectedCustomer As Customer = listViewSearchResults.SelectedItems(0).Tag

                Dim targetRange As Excel.Range = activeSheet.Range(targetRangeAddress)
                targetRange.Value2 = ToStringArray(selectedCustomer)
                targetRange.HorizontalAlignment = XlHAlign.xlHAlignLeft

                activeSheet.Columns(targetRange.Column).AutoFit()

                activeCell.Dispose()
                activeSheet.Dispose()

            End If

        End If

    End Sub

    Private Sub listViewSearchResults_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles listViewSearchResults.ItemSelectionChanged

        UpdateDetails()

    End Sub

    Private Sub textBoxSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textBoxSearch.TextChanged

        UpdateSearchResult()
        UpdateDetails()

    End Sub

#End Region

End Class
