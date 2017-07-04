Imports System.ComponentModel
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Windows.Threading
Imports OxyPlot
Imports OxyPlot.Axes
Imports OxyPlot.Series
Imports ServiceStack.DataAnnotations
Imports ServiceStack.OrmLite

<Serializable>
<[Alias]("hosts")>
Public Class clsHost
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _id As Integer?
    Private _description As String
    Private _location As String
    Private _hostname As String
    Private _ipaddress As String
    Private _macaddress As String
    Private _status As clsPingResult
    Private _laststatus As clsPingResult
    Private _pingpriority As String = "ipaddress"
    Private _updateinterval As Integer = preferences.DefaultUpdateInterval
    Private _type As String = "computer"

    Private _realtimemodel As New PlotModel
    Private _realtimeseriessuccess As New LineSeries
    Private _realtimeseriesfail As New LineSeries
    Private _realtimeaxesx As New DateTimeAxis
    Private _realtimeaxesy As New LinearAxis
    Private _realtimelastpoint As DataPoint

    Private updatetimer As New DispatcherTimer()

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

#Region "Constructors"

    Sub New()
        Initialize()
        StartPinger()
    End Sub

    Sub New(IPAddress As String, Hostname As String, MACAddress As String)
        Me.IPAddress = IPAddress
        Me.Hostname = Hostname
        Me.MACAddress = MACAddress
        Initialize()
        StartPinger()
    End Sub

#End Region

#Region "Subs"

    Private Sub Initialize()
        CreateRealTimeModel()

        AddHandler updatetimer.Tick, AddressOf updatetimer_Tick
        updatetimer.Interval = TimeSpan.FromSeconds(_updateinterval)
    End Sub

    Public Sub StartPinger()
        updatetimer.Start()
    End Sub

    Public Sub StopPinger()
        updatetimer.Stop()
    End Sub

    Private Sub updatetimer_Tick(sender As Object, e As EventArgs)
        Ping()
    End Sub

    Public Async Sub UpdateHostname()
        Await Task.Run(Sub() Hostname = GetHostnameFromIPAddress(_ipaddress))
        NotifyPropertyChanged("Hostname")
    End Sub

    Public Async Sub UpdateIPAddress()
        Await Task.Run(Sub() IPAddress = GetIPAddressFromHostname(_hostname))
        NotifyPropertyChanged("IPAddress")
    End Sub

    Public Async Sub UpdateMACAddress()
        Await Task.Run(Sub() MACAddress = GetMACAddressFromIPAddress(_ipaddress))
        NotifyPropertyChanged("MACAddress")
    End Sub

    Public Async Sub Ping()
        If Id Is Nothing Then Exit Sub

        If _pingpriority = "ipaddress" Then
            If _ipaddress IsNot Nothing Then Await Task.Run(Sub() _status = GetPingReply(_ipaddress))
        ElseIf _pingpriority = "hostname" Then
            If _hostname IsNot Nothing Then Await Task.Run(Sub() _status = GetPingReply(_hostname))
        Else
            Exit Sub
        End If

        If _status Is Nothing Then Exit Sub

        Dim hi As New clsHistoryItem(_id, _status)
        HistoryAdd(hi)

        If _pingpriority = "hostname" And _status.Status = IPStatus.Success And Not _status.Address.Contains(":") Then IPAddress = _status.Address.ToString

        UpdateRealTimeModel()

        NotifyPropertyChanged("Status")
        NotifyPropertyChanged("Image")
    End Sub

    Public Sub HistoryClear()
        db.Delete(Of clsHistoryItem)(where:=Function(x) x.HostId = Id)
    End Sub

    Public Function HistoryGet() As List(Of clsHistoryItem)
        Return db.Select(Of clsHistoryItem)(Function(x) x.HostId = Id)
    End Function

    Private Sub HistoryAdd(hi As clsHistoryItem)
        db.Insert(hi)
    End Sub

#End Region

#Region "Models subs"

    Public Sub CreateRealTimeModel()
        _realtimemodel = New PlotModel
        _realtimeseriessuccess = New LineSeries
        _realtimeseriesfail = New LineSeries
        _realtimeaxesx = New DateTimeAxis
        _realtimeaxesy = New LinearAxis

        _realtimeseriessuccess.Color = OxyColor.Parse(preferences.ColorSuccessBackground.ToString)
        _realtimeseriesfail.Color = OxyColor.Parse(preferences.ColorFailBackground.ToString)
        _realtimeaxesx.Minimum = DateTimeAxis.ToDouble(Now.AddSeconds(-_updateinterval * 60))
        _realtimeaxesx.Maximum = DateTimeAxis.ToDouble(Now)
        _realtimeaxesx.IsPanEnabled = False : _realtimeaxesx.IsZoomEnabled = False
        _realtimeaxesx.IsAxisVisible = False
        _realtimeaxesy.Minimum = 0.0
        _realtimeaxesy.IsPanEnabled = False : _realtimeaxesy.IsZoomEnabled = False
        _realtimeaxesy.StringFormat = "0 ms"
        _realtimemodel.PlotAreaBackground = OxyColor.Parse(preferences.ColorElementBackground.ToString)
        _realtimemodel.Series.Add(_realtimeseriessuccess)
        _realtimemodel.Series.Add(_realtimeseriesfail)
        _realtimemodel.Axes.Add(_realtimeaxesx)
        _realtimemodel.Axes.Add(_realtimeaxesy)
    End Sub

    Private Sub UpdateRealTimeModel()
        Dim dt = Now

        If _status IsNot Nothing And _laststatus IsNot Nothing Then
            Dim _realtimenewpoint As New DataPoint(DateTimeAxis.ToDouble(dt), If(_status.Status = IPStatus.Success, _status.RoundtripTime, 0))

            If _status.Status = IPStatus.Success And _laststatus.Status = IPStatus.Success Then 'still success
                _realtimeseriessuccess.Points.Add(_realtimenewpoint)
            ElseIf _status.Status = IPStatus.Success And _laststatus.Status <> IPStatus.Success Then 'became success
                _realtimeseriessuccess.Points.Add(_realtimenewpoint)
                _realtimeseriesfail.Points.Add(New DataPoint(_realtimenewpoint.X, 0))
                _realtimeseriesfail.Points.Add(_realtimenewpoint)
                _realtimeseriesfail.Points.Add(DataPoint.Undefined)
            ElseIf _status.Status <> IPStatus.Success And _laststatus.Status = IPStatus.Success Then 'became failed
                _realtimeseriessuccess.Points.Add(DataPoint.Undefined)
                _realtimeseriesfail.Points.Add(_realtimelastpoint)
                _realtimeseriesfail.Points.Add(New DataPoint(_realtimelastpoint.X, 0))
                _realtimeseriesfail.Points.Add(_realtimenewpoint)
            ElseIf _status.Status <> IPStatus.Success And _laststatus.Status <> IPStatus.Success Then 'still failed
                _realtimeseriesfail.Points.Add(_realtimenewpoint)
            End If

            _realtimelastpoint = _realtimenewpoint
        End If

        _laststatus = _status

        For I = 0 To _realtimeseriessuccess.Points.Count - 1
            If _realtimeseriessuccess.Points(0).X < DateTimeAxis.ToDouble(Now.AddSeconds(-_updateinterval * 60)) Then _realtimeseriessuccess.Points.RemoveAt(0)
        Next
        For I = 0 To _realtimeseriesfail.Points.Count - 1
            If _realtimeseriesfail.Points(0).X < DateTimeAxis.ToDouble(Now.AddSeconds(-_updateinterval * 60)) Then _realtimeseriesfail.Points.RemoveAt(0)
        Next

        _realtimeaxesx.Minimum = DateTimeAxis.ToDouble(Now.AddSeconds(-_updateinterval * 60))
        _realtimeaxesx.Maximum = DateTimeAxis.ToDouble(Now)
        _realtimemodel.InvalidatePlot(True)
    End Sub

    Public Function UpdateHistoryModel(Optional datefrom As Date = Nothing, Optional dateto As Date = Nothing) As PlotModel
        Dim model As New PlotModel
        Dim seriessuccess As New AreaSeries
        Dim seriesfail As New LineSeries
        Dim axesx As New DateTimeAxis
        Dim axesy As New LinearAxis
        Dim laststatus As Integer
        Dim lastpoint As DataPoint

        Dim history As List(Of clsHistoryItem) = HistoryGet()

        seriessuccess.Color = OxyColor.Parse(preferences.ColorSuccessBackground.ToString)
        seriesfail.Color = OxyColor.Parse(preferences.ColorFailBackground.ToString)
        axesy.StringFormat = "0 ms"
        axesy.IsPanEnabled = False : axesy.IsZoomEnabled = False
        axesx.StringFormat = "HH:mm dd.MM"

        seriessuccess.Points.Clear()
        seriesfail.Points.Clear()

        For I As Long = 1 To history.Count - 1
            Dim historynewpoint As New DataPoint(DateTimeAxis.ToDouble(history(I).TimeStamp), If(history(I).Status = IPStatus.Success, history(I).RoundTripTime, 0))
            Dim newstatus = history(I).Status

            If (history(I).TimeStamp - history(I - 1).TimeStamp).TotalMinutes > 1 Then
                seriessuccess.Points.Add(DataPoint.Undefined)
                seriesfail.Points.Add(DataPoint.Undefined)
            End If

            If newstatus = IPStatus.Success And laststatus = IPStatus.Success Then 'still success
                seriessuccess.Points.Add(historynewpoint)
            ElseIf newstatus = IPStatus.Success And laststatus <> IPStatus.Success Then 'became success
                seriessuccess.Points.Add(historynewpoint)
                seriesfail.Points.Add(New DataPoint(historynewpoint.X, 0))
                seriesfail.Points.Add(historynewpoint)
                seriesfail.Points.Add(DataPoint.Undefined)
            ElseIf newstatus <> IPStatus.Success And laststatus = IPStatus.Success Then 'became failed
                seriessuccess.Points.Add(DataPoint.Undefined)
                seriesfail.Points.Add(lastpoint)
                seriesfail.Points.Add(New DataPoint(lastpoint.X, 0))
                seriesfail.Points.Add(historynewpoint)
            ElseIf newstatus <> IPStatus.Success And laststatus <> IPStatus.Success Then 'still failed
                seriesfail.Points.Add(historynewpoint)
            End If

            lastpoint = historynewpoint
            laststatus = newstatus
        Next

        model.Series.Add(seriessuccess)
        model.Series.Add(seriesfail)
        model.Axes.Add(axesx)
        model.Axes.Add(axesy)
        model.Title = Description
        model.Subtitle = Hostname & If(Not String.IsNullOrEmpty(Hostname) And Not String.IsNullOrEmpty(IPAddress), " - ", "") & IPAddress
        model.InvalidatePlot(True)

        Return model
    End Function

    Private Function UpdateHistoryModelPie() As PlotModel
        Dim model As New PlotModel
        Dim series As New PieSeries
        Dim success As Long = 0
        Dim fail As Long = 0

        Dim history As List(Of clsHistoryItem) = HistoryGet()

        For I As Long = 1 To history.Count - 1
            If history(I).Status = 0 Then 'success
                success += 1
            Else
                fail += 1
            End If
        Next

        series.Diameter = 0.6
        series.InnerDiameter = 0.3
        series.InsideLabelFormat = ""
        series.OutsideLabelFormat = "{1} - {2:0.0}%"
        series.AngleSpan = 360
        series.StartAngle = 0
        series.TickLabelDistance = 5
        series.TickRadialLength = 0
        series.TickHorizontalLength = 0

        series.Slices.Add(New PieSlice("Успех", success) With {.IsExploded = False, .Fill = OxyColor.Parse(preferences.ColorSuccessBackground.ToString)})
        series.Slices.Add(New PieSlice("Неудача", fail) With {.IsExploded = False, .Fill = OxyColor.Parse(preferences.ColorFailBackground.ToString)})

        model.Series.Add(series)
        model.Title = "Доступность"
        model.InvalidatePlot(True)

        Return model
    End Function

    Private Function UpdateHistoryModelBar() As PlotModel
        Dim model As New PlotModel
        Dim series As New BarSeries
        Dim axesy As New CategoryAxis
        Dim success As New SortedDictionary(Of Long, Long)
        Dim values As New List(Of BarItem)
        Dim labels As New List(Of String)

        Dim history As List(Of clsHistoryItem) = HistoryGet()

        For I As Long = 1 To history.Count - 1
            If history(I).Status = 0 Then 'success
                Dim category As Integer = If(history(I).RoundTripTime > 0, Math.Log10(history(I).RoundTripTime) / Math.Log10(2), 0)
                If Not success.ContainsKey(category) Then success.Add(category, 0)
                success(category) += 1
            End If
        Next

        For Each v In success.Values
            values.Add(New BarItem(v / history.Count * 100))
        Next

        For Each l In success.Keys
            labels.Add(2 ^ l & " мс.")
        Next

        series.LabelPlacement = LabelPlacement.Inside
        series.LabelFormatString = "{0:.00}%"
        series.ItemsSource = values
        series.FillColor = OxyColor.Parse(preferences.ColorSuccessBackground.ToString)

        axesy.Position = AxisPosition.Left
        axesy.ItemsSource = labels

        model.Series.Add(series)
        model.Axes.Add(axesy)
        model.Title = "Статистика"
        model.InvalidatePlot(True)

        Return model
    End Function

#End Region

#Region "Properties"

    <AutoIncrement>
    Public Property Id() As Integer?
        Get
            Return _id
        End Get
        Set(value As Integer?)
            _id = value

            NotifyPropertyChanged("Id")
        End Set
    End Property

    <Ignore>
    Public ReadOnly Property Image() As BitmapImage
        Get
            Dim status As String = "unknown"

            If _status IsNot Nothing Then
                If _status.Status = IPStatus.Success Then
                    status = "active"
                Else
                    status = "inactive"
                End If
            Else
                status = "unknown"
            End If

            Return New BitmapImage(New Uri("pack://application:,,,/" & "img/host/" & Type & "-" & status & ".png"))
        End Get
    End Property

    <[Alias]("description")>
    <CustomField("VARCHAR(1024)")>
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.Description, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("Description")
        End Set
    End Property

    <[Alias]("hostname")>
    <CustomField("VARCHAR(256)")>
    Public Property Hostname As String
        Get
            Return _hostname
        End Get
        Set(value As String)
            _hostname = value

            If String.IsNullOrEmpty(_description) Then
                Description = value
            End If

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.Hostname, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("Hostname")
        End Set
    End Property

    <[Alias]("ipaddress")>
    <CustomField("VARCHAR(46)")>
    Public Property IPAddress As String
        Get
            Return _ipaddress
        End Get
        Set(value As String)
            _ipaddress = value

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.IPAddress, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("IPAddress")
        End Set
    End Property

    <Ignore>
    Public ReadOnly Property Status As String
        Get
            If _status IsNot Nothing Then
                If _status.Status = IPStatus.Success Then
                    Return String.Format("{0} - {1} мс", _status.Address.ToString, _status.RoundTripTime)
                Else
                    Return CType([Enum].Parse(GetType(IPStatus), _status.Status), IPStatus).ToString
                End If
            Else
                Return "Статус неизвестен"
            End If
        End Get
    End Property

    <[Alias]("macaddress")>
    <CustomField("VARCHAR(20)")>
    Public Property MACAddress As String
        Get
            If _macaddress Is Nothing Then
                Return Nothing
            Else
                Return _macaddress
            End If
        End Get
        Set(value As String)
            _macaddress = value

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.MACAddress, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("MACAddress")
        End Set
    End Property

    <[Alias]("pingpriority")>
    <CustomField("VARCHAR(20)")>
    Public Property PingPriority As String
        Get
            If Not pingpriorities.Keys.Contains(_pingpriority) Then
                _pingpriority = "ipaddress"
            End If

            Return _pingpriority
        End Get
        Set(value As String)
            If pingpriorities.Keys.Contains(value) Then
                _pingpriority = value
            Else
                _pingpriority = "ipaddress"
            End If

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.PingPriority, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("PingPriority")
        End Set
    End Property

    <[Alias]("updateinterval")>
    <CustomField("INTEGER")>
    Public Property UpdateInterval As Integer
        Get
            Return _updateinterval
        End Get
        Set(value As Integer)
            If value < 1 Then value = 1
            _updateinterval = value
            updatetimer.Interval = TimeSpan.FromSeconds(_updateinterval)

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.UpdateInterval, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("UpdateInterval")
        End Set
    End Property

    <[Alias]("type")>
    <CustomField("VARCHAR(20)")>
    Public Property Type As String
        Get
            If Not hosttypes.Keys.Contains(_type) Then
                _type = "computer"
            End If

            Return _type
        End Get
        Set(value As String)
            If hosttypes.Keys.Contains(value) Then
                _type = value
            Else
                _type = "computer"
            End If

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.Type, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("Type")
            NotifyPropertyChanged("Image")
        End Set
    End Property

    <[Alias]("location")>
    <CustomField("VARCHAR(1024)")>
    Public Property Location As String
        Get
            Return _location
        End Get
        Set(value As String)
            _location = value

            If Id IsNot Nothing Then db.UpdateOnly(obj:=Me, onlyFields:=Function(p) p.Location, where:=Function(p) p.Id = Id)

            NotifyPropertyChanged("Location")
        End Set
    End Property

    <Ignore>
    Public ReadOnly Property RealTimeModel As PlotModel
        Get
            Return _realtimemodel
        End Get
    End Property

    <Ignore>
    Public ReadOnly Property HistoryModel As PlotModel
        Get
            Return UpdateHistoryModel()
        End Get
    End Property

    <Ignore>
    Public ReadOnly Property HistoryModelPie As PlotModel
        Get
            Return UpdateHistoryModelPie()
        End Get
    End Property

    <Ignore>
    Public ReadOnly Property HistoryModelBar As PlotModel
        Get
            Return UpdateHistoryModelBar()
        End Get
    End Property

#End Region

End Class
