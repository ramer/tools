Public Class clsPingResult

    Private _status As Integer
    Private _roundtriptime As Integer
    Private _address As String

    Sub New()

    End Sub

    Sub New(Status As Integer, RoundTripTime As Integer, Address As String)
        _status = Status
        _roundtriptime = RoundTripTime
        _address = Address
    End Sub

    Public Property Status As Integer
        Get
            Return _status
        End Get
        Set(value As Integer)
            _status = value
        End Set
    End Property

    Public Property RoundTripTime As Integer
        Get
            Return _roundtriptime
        End Get
        Set(value As Integer)
            _roundtriptime = value
        End Set
    End Property

    Public Property Address As String
        Get
            Return _address
        End Get
        Set(value As String)
            _address = value
        End Set
    End Property
End Class
