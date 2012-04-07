Public Class Customer

    Dim _id As Integer
    Dim _name As String
    Dim _company As String
    Dim _mail As String
    Dim _city As String
    Dim _postalCode As String
    Dim _country As String
    Dim _phone As String

    Public Sub New(ByVal id As Integer, ByVal name As String, ByVal company As String, ByVal mail As String, ByVal city As String, ByVal postalCode As String, ByVal country As String, ByVal phone As String)

        _id = id
        _company = company
        _mail = mail
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

    Public ReadOnly Property Mail() As String
        Get
            Return _mail
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

End Class
