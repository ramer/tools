Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices

Public Class clsDirectoryContainer
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _entry As DirectoryEntry
    Private _children As New ObservableCollection(Of clsDirectoryContainer)

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Entry As DirectoryEntry)
        _entry = Entry
    End Sub

    Public Property Entry() As DirectoryEntry
        Get
            Return _entry
        End Get
        Set(ByVal value As DirectoryEntry)
            _entry = value
        End Set
    End Property

    Public ReadOnly Property SchemaClassName() As String
        Get
            Return _entry.SchemaClassName
        End Get
    End Property

    Public ReadOnly Property Children() As ObservableCollection(Of clsDirectoryContainer)
        Get
            _children.Clear()

            Dim ds As New DirectorySearcher(_entry)
            ds.PropertiesToLoad.AddRange({"name", "objectClass"})
            ds.Filter = "(|(objectClass=organizationalUnit)(objectClass=container))"
            ds.SearchScope = SearchScope.OneLevel
            For Each sr As SearchResult In ds.FindAll()
                _children.Add(New clsDirectoryContainer(sr.GetDirectoryEntry()))
            Next

            Return _children
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return GetLDAPProperty(_entry.Properties, "name")
        End Get
    End Property

    Public ReadOnly Property Image() As String
        Get
            Select Case SchemaClassName
                Case "domainDNS"
                    Return "img/domain.ico"
                Case "organizationalUnit"
                    Return "img/organizationalunit.ico"
                Case "container"
                    Return "img/container.ico"
                Case Else
                    Return "img/organizationalunit.ico"
            End Select
        End Get
    End Property
End Class
