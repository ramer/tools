Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.Security.Principal

Public Class clsDeletedDirectoryObject
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _entry As SearchResult
    Private _domain As clsDomain

    Private _cn As String '
    Private _distinguishedName As String ' 
    Private _lastKnownParent As String '
    Private _lastKnownParentExist As Boolean? '
    Private _name As String '
    Private _objectClass As String
    Private _objectGUID As String '
    Private _objectSID As String '
    Private _sAMAccountName As String '
    Private _userAccountControl As Long = -1
    Private _whenChanged As Date '
    Private _whenCreated As Date '

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Entry As SearchResult, ByRef Domain As clsDomain)
        _entry = Entry
        _domain = Domain
    End Sub

    Public Sub Refresh()
        _cn = Nothing
        _distinguishedName = Nothing
        _lastKnownParent = Nothing
        _name = Nothing
        _objectClass = Nothing
        _objectGUID = Nothing
        _objectSID = Nothing
        _sAMAccountName = Nothing
        _userAccountControl = -1
        _whenChanged = Nothing
        _whenCreated = Nothing
    End Sub

    Public Function GetProperty(PropertyName As String) As Object
        Dim pvc As ResultPropertyValueCollection = _entry.Properties(PropertyName)

        If pvc.Count = 0 Then
            Return Nothing
        ElseIf pvc.Count = 1 Then
            If pvc.Item(0) Is Nothing Then Return Nothing

            Select Case pvc(0).GetType()
                Case GetType(String)
                    Return pvc.Item(0)
                Case GetType(Integer)
                    Return pvc.Item(0)
                Case GetType(Byte())
                    Return pvc.Item(0)
                Case GetType(Object())
                    Return pvc.Item(0)
                Case GetType(DateTime)
                    Return pvc.Item(0)
                Case GetType(Boolean)
                    Return pvc.Item(0)
                Case Else 'System.__ComObject
                    Return LongFromLargeInteger(pvc.Item(0))
            End Select
        Else
            Dim res As New List(Of Object)

            For I As Integer = 0 To pvc.Count - 1
                If pvc.Item(I) Is Nothing Then Continue For
                Select Case pvc(I).GetType()
                    Case GetType(String)
                        res.Add(pvc.Item(I))
                    Case GetType(Integer)
                        res.Add(pvc.Item(I))
                    Case GetType(Byte())
                        res.Add(pvc.Item(I))
                    Case GetType(Object())
                        res.Add(pvc.Item(I))
                    Case GetType(DateTime)
                        res.Add(pvc.Item(I))
                    Case GetType(Boolean)
                        res.Add(pvc.Item(I))
                    Case Else 'System.__ComObject
                        res.Add(LongFromLargeInteger(pvc.Item(I)))
                End Select
            Next

            Return res.ToArray
        End If
    End Function

    Public ReadOnly Property Entry() As SearchResult
        Get
            Return _entry
        End Get
    End Property

    Public ReadOnly Property Domain() As clsDomain
        Get
            Return _domain
        End Get
    End Property

    Public ReadOnly Property SchemaClassName() As String
        Get
            If objectClass.Contains("computer") Then
                Return "computer"
            ElseIf objectClass.Contains("user") Then
                Return "user"
            ElseIf objectClass.Contains("group") Then
                Return "group"
            ElseIf objectClass.Contains("organizationalUnit") Then
                Return "organizationalunit"
            ElseIf objectClass.Contains("contact") Then
                Return "contact"
            Else
                Return "unknown"
            End If
        End Get
    End Property

    Public ReadOnly Property Image() As String
        Get
            If objectClass.Contains("computer") Then
                Return "img/computer.ico"
            ElseIf objectClass.Contains("user") Then
                Return "img/user.ico"
            ElseIf objectClass.Contains("group") Then
                Return "img/group.ico"
            ElseIf objectClass.Contains("organizationalUnit") Then
                Return "img/organizationalunit.ico"
            ElseIf objectClass.Contains("contact") Then
                Return "img/contact.ico"
            Else
                Return "img/object_unknown.ico"
            End If
        End Get
    End Property

    'cached ldap properties

#Region "Shared attributes"

    Public ReadOnly Property cn() As String
        Get
            _cn = If(_cn, GetProperty("cn"))
            Return _cn
        End Get
    End Property

    Public ReadOnly Property distinguishedName() As String
        Get
            _distinguishedName = If(_distinguishedName, GetProperty("distinguishedName"))
            Return _distinguishedName
        End Get
    End Property

    Public ReadOnly Property lastKnownParent() As String
        Get
            _lastKnownParent = If(_lastKnownParent, GetProperty("lastKnownParent"))
            Return _lastKnownParent
        End Get
    End Property

    Public ReadOnly Property lastKnownParentExist() As Boolean
        Get
            If _lastKnownParentExist Is Nothing Then
                Try
                    Dim parent As New DirectoryEntry("LDAP://" & _domain.Name & "/" & lastKnownParent, _domain.Username, _domain.Password)
                    _lastKnownParentExist = parent.Properties.Count > 0
                Catch
                    _lastKnownParentExist = False
                End Try
            End If
            Return _lastKnownParentExist
        End Get
    End Property

    Public ReadOnly Property lastKnownParentImage() As String
        Get
            If lastKnownParentExist Then
                Return "img/ready.ico"
            Else
                Return "img/warning.ico"
            End If
        End Get
    End Property

    Public ReadOnly Property name() As String
        Get
            _name = If(_name, GetProperty("name"))
            Return _name
        End Get
    End Property

    Public ReadOnly Property nameFormated() As String
        Get
            Dim arr() As String = name.Split({Chr(10), vbCr, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            Return If(arr.Count >= 1, arr(0), "")
        End Get
    End Property

    Public ReadOnly Property objectClass() As String
        Get
            _objectClass = If(_objectClass, Join(CType(GetProperty("objectClass"), Object()), ", "))
            Return _objectClass
        End Get
    End Property

    Public ReadOnly Property objectGUID() As String
        Get
            _objectGUID = If(_objectGUID, New Guid(TryCast(GetProperty("objectGUID"), Byte())).ToString)
            Return _objectGUID
        End Get
    End Property

    Public ReadOnly Property objectSID() As String
        Get
            _objectSID = If(_objectSID, If(GetProperty("objectSID") IsNot Nothing, New SecurityIdentifier(GetProperty("objectSID"), 0).Value, ""))
            Return _objectSID
        End Get
    End Property

    Public ReadOnly Property sAMAccountName() As String
        Get
            _sAMAccountName = If(_sAMAccountName, GetProperty("sAMAccountName"))
            Return _sAMAccountName
        End Get
    End Property

    Public ReadOnly Property userAccountControl() As Long
        Get
            _userAccountControl = If(_userAccountControl >= 0, _userAccountControl, GetProperty("userAccountControl"))
            Return _userAccountControl
        End Get
    End Property

    Public ReadOnly Property whenCreated() As Date
        Get
            _whenCreated = If(_whenCreated = Nothing, GetProperty("whenCreated"), _whenCreated)
            Return _whenCreated
        End Get
    End Property

    Public ReadOnly Property whenCreatedFormated() As String
        Get
            Return If(whenCreated = Nothing, "неизвестно", whenCreated.ToString)
        End Get
    End Property

    Public ReadOnly Property whenChanged() As Date
        Get
            _whenChanged = If(_whenChanged = Nothing, GetProperty("whenChanged"), _whenChanged)
            Return _whenChanged
        End Get
    End Property

    Public ReadOnly Property whenChangedFormated() As String
        Get
            Return If(whenChanged = Nothing, "неизвестно", whenChanged.ToString)
        End Get
    End Property

#End Region

End Class
