Public Class clsSeriesItem

    Private _timestamp As Date
    Private _roundtriptime As Double
    Private _availability As Double

    Sub New(TimeStamp As Date, RoundTripTime As Double, Availability As Double)
        _timestamp = TimeStamp
        _roundtriptime = RoundTripTime
        _availability = Availability
    End Sub

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

    Public ReadOnly Property RoundTripTime() As Double
        Get
            Return _roundtriptime
        End Get
    End Property

    Public ReadOnly Property Availability() As Double
        Get
            Return _availability
        End Get
    End Property

End Class
