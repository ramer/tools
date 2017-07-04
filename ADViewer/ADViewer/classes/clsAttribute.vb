Public Class clsAttribute
    Private _name As String
    Private _label As String
    Private _isdefault As Boolean

    Sub New()

    End Sub

    Sub New(Name As String,
            Label As String,
            Optional IsDefault As Boolean = False)

        _name = Name
        _label = Label
        _isdefault = IsDefault
    End Sub

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Label() As String
        Get
            Return _label
        End Get
        Set(ByVal value As String)
            _label = value
        End Set
    End Property

    Public Property IsDefault() As Boolean
        Get
            Return _isdefault
        End Get
        Set(ByVal value As Boolean)
            _isdefault = value
        End Set
    End Property

End Class
