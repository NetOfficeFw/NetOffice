Imports LateBindingApi.Core
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
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

        Try

            If (listViewSearchResults.SelectedItems.Count > 0) Then

                ' any action...

            End If

        Catch ex As Exception

            MessageBox.Show(Me, ex.Message, "An error is occured", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub listViewSearchResults_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles listViewSearchResults.ItemSelectionChanged

        Try

            UpdateDetails()

        Catch ex As Exception

            MessageBox.Show(Me, ex.Message, "An error is occured", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub textBoxSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textBoxSearch.TextChanged

        Try

            UpdateSearchResult()
            UpdateDetails()

        Catch ex As Exception

            MessageBox.Show(Me, ex.Message, "An error is occured", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region


End Class
