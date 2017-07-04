Imports System.Net.NetworkInformation
Imports System.Xml.Serialization
Imports ServiceStack.DataAnnotations

<Serializable>
<[Alias]("history")>
<XmlType("HistoryItem")>
Public Class clsHistoryItem

    Private _id As Integer
    Private _hostid As Integer
    Private _timestamp As Date
    Private _roundtriptime As Integer
    Private _status As Integer

    Sub New()

    End Sub

    Sub New(HostId As Integer, PingResult As clsPingResult)
        _hostid = HostId
        _timestamp = Now
        _status = PingResult.Status
        _roundtriptime = PingResult.RoundTripTime
    End Sub

    <AutoIncrement>
    <XmlIgnore>
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property

    <[Alias]("hostid")>
    <XmlIgnore>
    Public Property HostId() As Integer
        Get
            Return _hostid
        End Get
        Set(value As Integer)
            _hostid = value
        End Set
    End Property

    <[Alias]("timestamp")>
    <CustomField("DATETIME")>
    <XmlElement("TimeStamp")>
    Public Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
        Set(value As Date)
            _timestamp = value
        End Set
    End Property

    <[Alias]("roundtriptime")>
    <XmlElement("Time")>
    Public Property RoundTripTime() As Integer
        Get
            Return _roundtriptime
        End Get
        Set(value As Integer)
            _roundtriptime = value
        End Set
    End Property

    <[Alias]("status")>
    <XmlElement("Status")>
    Public Property Status() As Integer
        Get
            Return _status
        End Get
        Set(value As Integer)
            _status = value
        End Set
    End Property

End Class
