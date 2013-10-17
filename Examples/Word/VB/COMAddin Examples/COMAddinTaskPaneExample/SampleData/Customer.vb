Public Class Customer

    Dim _id As Integer
    Dim _name As String
    Dim _company As String
    Dim _city As String
    Dim _postalCode As String
    Dim _country As String
    Dim _phone As String

    Public Sub New(ByVal id As Integer, ByVal name As String, ByVal company As String, ByVal city As String, ByVal postalCode As String, ByVal country As String, ByVal phone As String)

        _id = id
        _company = company
        _name = name
        _city = city
        _postalCode = postalCode
        _country = country
        _phone = phone

    End Sub

    Public ReadOnly Property ID() As Integer
        Get
            Return _id
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _name
        End Get
    End Property

    Public ReadOnly Property Company() As String
        Get
            Return _company
        End Get
    End Property

    Public ReadOnly Property City() As String
        Get
            Return _city
        End Get
    End Property

    Public ReadOnly Property PostalCode() As String
        Get
            Return _postalCode
        End Get
    End Property

    Public ReadOnly Property Country() As String
        Get
            Return _country
        End Get
    End Property

    Public ReadOnly Property Phone() As String
        Get
            Return _phone
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return String.Format("Name: {1}{7}ID: {0}{7}Company: {2}{7}City: {3}{7}PostalCode: {4}{7}Country: {5}{7}Phone: {6}{7}{7}", _id, _name, _company, _city, _postalCode, _country, _phone, Environment.NewLine)
    End Function

End Class
