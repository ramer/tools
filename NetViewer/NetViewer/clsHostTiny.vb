Imports System.ComponentModel
Imports System.Net
Imports System.Net.NetworkInformation

Public Class clsHostTiny
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _hostname As String
    Private _ipaddress As String
    Private _macaddress As String

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Status As clsPingResult)
        _ipaddress = Status.Address

        UpdateHostname()
        UpdateMACAddress()
    End Sub

    Public Async Sub UpdateHostname()
        Await Task.Run(Sub() _hostname = GetHostnameFromIPAddress(_ipaddress))
        NotifyPropertyChanged("Hostname")
    End Sub

    Public Async Sub UpdateIPAddress()
        Await Task.Run(Sub() _ipaddress = GetIPAddressFromHostname(_hostname))
        NotifyPropertyChanged("IPAddress")
    End Sub

    Public Async Sub UpdateMACAddress()
        Await Task.Run(Sub() _macaddress = GetMACAddressFromIPAddress(_ipaddress))
        NotifyPropertyChanged("MACAddress")
    End Sub

    Public ReadOnly Property Image() As BitmapImage
        Get
            Return New BitmapImage(New Uri("pack://application:,,,/" & "img/host/computer-active.png"))
        End Get
    End Property

    Public ReadOnly Property Hostname As String
        Get
            Return _hostname
        End Get
    End Property

    Public ReadOnly Property IPAddress As String
        Get
            Return _ipaddress
        End Get
    End Property

    Public ReadOnly Property MACAddress As String
        Get
            Return _macaddress
        End Get
    End Property

End Class
