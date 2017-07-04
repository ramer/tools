Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.Security.Principal

Public Class clsDirectoryObject
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _entry As DirectoryEntry
    Private _domain As clsDomain
    Private _tag As Object
    Private _children As New ObservableCollection(Of clsDirectoryObject)

    Private _sn As String
    Private _initials As String
    Private _givenName As String
    Private _distinguishedName As String
    Private _displayName As String
    Private _description As String
    Private _info As String
    Private _location As String
    Private _physicalDeliveryOfficeName As String
    Private _company As String
    Private _manager As clsDirectoryObject
    Private _employees As ObservableCollection(Of clsDirectoryObject)
    Private _name As String
    Private _telephoneNumber As String
    Private _homePhone As String
    Private _ipPhone As String
    Private _mobile As String
    Private _streetAddress As String
    Private _mail As String
    Private _title As String
    Private _department As String
    Private _sAMAccountName As String
    Private _userPrincipalName As String
    Private _groupType As Long = 0
    Private _userAccountControl As Long = -1
    Private _badPwdCount As Integer = -1
    Private _logonCount As Integer = -1
    Private _objectGUID As String
    Private _objectSID As String
    Private _memberOf As ObservableCollection(Of clsDirectoryObject)
    Private _member As ObservableCollection(Of clsDirectoryObject)
    'Private _publicDelegates() As Object
    Private _whenCreated As Date
    Private _whenChanged As Date
    Private _lastLogon As Long = -1
    Private _pwdLastSet As Long = -1
    Private _passwordExpiresDate As Date
    Private _accountExpires As Long = -1
    Private _dnshostname As String
    Private _operatingsystemversion As String
    Private _operatingsystem As String
    Private _thumbnailPhoto() As Byte
    Private _userWorkstations() As String

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Entry As DirectoryEntry, ByRef Domain As clsDomain)
        _entry = Entry
        _domain = Domain
    End Sub

    Public Sub Refresh()
        _sn = Nothing
        _initials = Nothing
        _givenName = Nothing
        _distinguishedName = Nothing
        _displayName = Nothing
        _description = Nothing
        _physicalDeliveryOfficeName = Nothing
        _company = Nothing
        _manager = Nothing
        _employees = Nothing
        _name = Nothing
        _telephoneNumber = Nothing
        _homePhone = Nothing
        _mobile = Nothing
        _ipPhone = Nothing
        _streetAddress = Nothing
        _mail = Nothing
        _title = Nothing
        _department = Nothing
        _sAMAccountName = Nothing
        _userPrincipalName = Nothing
        _groupType = 0
        _userAccountControl = -1
        _badPwdCount = -1
        _logonCount = -1
        _objectGUID = Nothing
        _objectSID = Nothing
        _memberOf = Nothing
        _member = Nothing
        _whenCreated = Nothing
        _whenChanged = Nothing
        _lastLogon = -1
        _pwdLastSet = -1
        _passwordExpiresDate = Nothing
        _accountExpires = -1
        _thumbnailPhoto = Nothing

        Entry.RefreshCache()
    End Sub

    Public Function GetProperty(PropertyName As String) As Object
        Dim pvc As PropertyValueCollection = _entry.Properties(PropertyName)

        If pvc.Value Is Nothing Then Return Nothing

        Select Case pvc.Value.GetType()
            Case GetType(String)
                Return pvc.Value
            Case GetType(Integer)
                Return pvc.Value
            Case GetType(Byte())
                Return pvc.Value
            Case GetType(Object())
                Return pvc.Value
            Case GetType(DateTime)
                Return pvc.Value
            Case GetType(Boolean)
                Return pvc.Value
            Case Else 'System.__ComObject
                Return LongFromLargeInteger(pvc.Value)
        End Select

    End Function

    Public Function SetProperty(PropertyName As String, Value As Object) As Boolean
        'Try
        If IsNumeric(Value) Or IsDate(Value) Or Len(Value) > 0 Then
            _entry.Properties(PropertyName).Value = Trim(Value)
        Else
            _entry.Properties(PropertyName).Clear()
        End If

        _entry.CommitChanges()

        Return True
        'Catch ex As Exception
        '    ThrowException(ex, "SetProperty")
        '    Return False
        'End Try
    End Function

    Public ReadOnly Property CustomProperty(propertyname As String) As Object
        Get
            Return GetProperty(propertyname)
        End Get
    End Property

    Public Property Entry() As DirectoryEntry
        Get
            Return _entry
        End Get
        Set(ByVal value As DirectoryEntry)
            _entry = value

            NotifyPropertyChanged("Entry")
        End Set
    End Property

    Public ReadOnly Property Domain() As clsDomain
        Get
            Return _domain
        End Get
    End Property

    Public ReadOnly Property SchemaClassName() As String
        Get
            Return _entry.SchemaClassName
        End Get
    End Property

    Public ReadOnly Property Children() As ObservableCollection(Of clsDirectoryObject)
        Get
            _children.Clear()

            Dim ds As New DirectorySearcher(_entry)
            ds.PropertiesToLoad.AddRange({"name", "objectClass"})
            ds.Filter = "(|(objectClass=organizationalUnit)(objectClass=container))"
            ds.SearchScope = SearchScope.OneLevel
            For Each sr As SearchResult In ds.FindAll()
                _children.Add(New clsDirectoryObject(sr.GetDirectoryEntry(), Domain))
            Next

            Return _children
        End Get
    End Property

    Public Sub ResetPassword()
        If Domain.DefaultPassword = "" Then Throw New Exception("Стандартный пароль в целевом домене не указан")

        _entry.Invoke("SetPassword", Domain.DefaultPassword)
        pwdLastSet = 0
        _entry.CommitChanges()
        description = String.Format("{0} {1} ({2})", "Пароль сброшен", Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString)
    End Sub

    Public Sub SetPassword(password As String)
        _entry.Invoke("SetPassword", password)
        pwdLastSet = -1
        _entry.CommitChanges()
        description = String.Format("{0} {1} ({2})", "Пароль сброшен", Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString)
    End Sub

    Public ReadOnly Property Status() As String
        Get
            Dim _statusFormated As String = ""

            Select Case SchemaClassName
                Case "user"
                    If passwordNeverExpires = False AndAlso passwordExpiresDate() <= Now Then
                        _statusFormated &= "Пароль истек" & vbCr
                    End If
                    If accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                        _statusFormated &= "Срок действия истек" & vbCr
                    End If
                    If disabled Then
                        _statusFormated &= "Учетная запись заблокирована" & vbCr
                    End If
                    If _statusFormated.Length > 1 Then _statusFormated = _statusFormated.Remove(_statusFormated.Length - 1)
                Case "computer"
                    If passwordNeverExpires = False AndAlso passwordExpiresDate() <= Now Then
                        _statusFormated &= "Пароль истек" & vbCr
                    End If
                    If accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                        _statusFormated &= "Срок действия истек" & vbCr
                    End If
                    If disabled Then
                        _statusFormated &= "Компьютер заблокирован" & vbCr
                    End If
                    If _statusFormated.Length > 1 Then _statusFormated = _statusFormated.Remove(_statusFormated.Length - 1)
            End Select
            Return _statusFormated
        End Get
    End Property

    Public Property Tag() As Object
        Get
            Return _tag
        End Get
        Set(ByVal value As Object)
            _tag = value
        End Set
    End Property

    Public ReadOnly Property Image() As BitmapImage
        Get
            Dim _image As String = ""

            Select Case SchemaClassName
                Case "user"
                    _image = "img/user.ico"
                    If passwordNeverExpires = False AndAlso passwordExpiresDate() <= Now Then
                        _image = "img/user_expired.ico"
                    End If
                    If accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                        _image = "img/user_expired.ico"
                    End If
                    If disabled Then
                        _image = "img/user_blocked.ico"
                    End If
                Case "computer"
                    _image = "img/computer.ico"
                    If passwordNeverExpires = False AndAlso passwordExpiresDate() <= Now Then
                        _image = "img/computer_expired.ico"
                    End If
                    If accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                        _image = "img/computer_expired.ico"
                    End If
                    If disabled Then
                        _image = "img/computer_blocked.ico"
                    End If
                Case "group"
                    If groupTypeSecurity Then
                        _image = "img/group.ico"
                    ElseIf groupTypeDistribution Then
                        _image = "img/group_distribution.ico"
                    Else
                        _image = "img/object_unknown.ico"
                    End If
                Case "contact"
                    _image = "img/contact.ico"
                Case "domainDNS"
                    _image = "img/domain.ico"
                Case "organizationalUnit"
                    _image = "img/organizationalunit.ico"
                Case "container"
                    _image = "img/container.ico"
                Case Else
                    _image = "img/object_unknown.ico"
            End Select

            Return New BitmapImage(New Uri("pack://application:,,,/" & _image))
        End Get
    End Property

    'cached ldap properties


#Region "User attributes"

    Public Property sn() As String
        Get
            _sn = If(_sn, If(GetProperty("sn"), ""))
            Return _sn
        End Get
        Set(ByVal value As String)
            If SetProperty("sn", value) Then _sn = value

            NotifyPropertyChanged("sn")
        End Set
    End Property

    Public Property initials() As String
        Get
            _initials = If(_initials, If(GetProperty("initials"), ""))
            Return _initials
        End Get
        Set(ByVal value As String)
            If SetProperty("initials", value) Then _initials = value

            NotifyPropertyChanged("initials")
        End Set
    End Property

    Public Property givenName() As String
        Get
            _givenName = If(_givenName, If(GetProperty("givenName"), ""))
            Return _givenName
        End Get
        Set(ByVal value As String)
            If SetProperty("givenName", value) Then _givenName = value

            NotifyPropertyChanged("givenName")
        End Set
    End Property

    Public Property displayName() As String
        Get
            _displayName = If(_displayName, If(GetProperty("displayName"), ""))
            Return _displayName
        End Get
        Set(ByVal value As String)
            If SetProperty("displayName", value) Then _displayName = value

            NotifyPropertyChanged("displayName")
        End Set
    End Property

    Public Property physicalDeliveryOfficeName() As String
        Get
            _physicalDeliveryOfficeName = If(_physicalDeliveryOfficeName, If(GetProperty("physicalDeliveryOfficeName"), ""))
            Return _physicalDeliveryOfficeName
        End Get
        Set(ByVal value As String)
            If SetProperty("physicalDeliveryOfficeName", value) Then _physicalDeliveryOfficeName = value

            NotifyPropertyChanged("physicalDeliveryOfficeName")
        End Set
    End Property

    Public Property company() As String
        Get
            _company = If(_company, If(GetProperty("company"), ""))
            Return _company
        End Get
        Set(ByVal value As String)
            If SetProperty("company", value) Then _company = value

            NotifyPropertyChanged("company")
        End Set
    End Property

    Public Property manager() As clsDirectoryObject
        Get
            If _manager Is Nothing Then
                Try
                    Dim managerDN As String = GetProperty("manager")
                    If managerDN Is Nothing Then
                        _manager = Nothing
                    Else
                        _manager = New clsDirectoryObject(New DirectoryEntry("LDAP://" & Domain.Name & "/" & managerDN, Domain.Username, Domain.Password), Domain)
                    End If
                    Return _manager
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                Return _manager
            End If
        End Get
        Set(value As clsDirectoryObject)
            Dim path() As String = value.Entry.Path.Split({"/"}, StringSplitOptions.RemoveEmptyEntries)
            If SetProperty("manager", path(UBound(path))) Then _manager = value

            NotifyPropertyChanged("manager")
        End Set
    End Property

    'Public Property Employees As ObservableCollection(Of clsDirectoryObject)
    '    Get
    '        If _employees Is Nothing Then
    '            Try
    '                Dim LDAPsearcher As New DirectorySearcher(Domain.DefaultNamingContext)
    '                LDAPsearcher.Filter = "(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(objectClass=contact))(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(manager=" & distinguishedName & "))"

    '                _employees = New ObservableCollection(Of clsDirectoryObject)
    '                For Each result As SearchResult In LDAPsearcher.FindAll()
    '                    _employees.Add(New clsDirectoryObject(result.GetDirectoryEntry, Domain))
    '                Next
    '                Return _employees
    '            Catch ex As Exception
    '                Return Nothing
    '            End Try
    '        Else
    '            Return _employees
    '        End If
    '    End Get
    '    Set(value As ObservableCollection(Of clsDirectoryObject))

    '    End Set
    'End Property

    Public Property telephoneNumber() As String
        Get
            _telephoneNumber = If(_telephoneNumber, If(GetProperty("telephoneNumber"), ""))
            Return _telephoneNumber
        End Get
        Set(ByVal value As String)
            If SetProperty("telephoneNumber", value) Then _telephoneNumber = value

            NotifyPropertyChanged("telephoneNumber")
        End Set
    End Property

    Public Property homePhone() As String
        Get
            _homePhone = If(_homePhone, If(GetProperty("homePhone"), ""))
            Return _homePhone
        End Get
        Set(ByVal value As String)
            If SetProperty("homePhone", value) Then _homePhone = value

            NotifyPropertyChanged("homePhone")
        End Set
    End Property

    Public Property ipPhone() As String
        Get
            _ipPhone = If(_ipPhone, If(GetProperty("ipPhone"), ""))
            Return _ipPhone
        End Get
        Set(ByVal value As String)
            If SetProperty("ipPhone", value) Then _ipPhone = value

            NotifyPropertyChanged("ipPhone")
        End Set
    End Property

    Public Property mobile() As String
        Get
            _mobile = If(_mobile, If(GetProperty("mobile"), ""))
            Return _mobile
        End Get
        Set(ByVal value As String)
            If SetProperty("mobile", value) Then _mobile = value

            NotifyPropertyChanged("mobile")
        End Set
    End Property

    Public Property streetAddress() As String
        Get
            _streetAddress = If(_streetAddress, If(GetProperty("streetAddress"), ""))
            Return _streetAddress
        End Get
        Set(ByVal value As String)
            If SetProperty("streetAddress", value) Then _streetAddress = value

            NotifyPropertyChanged("streetAddress")
        End Set
    End Property

    Public Property mail() As String
        Get
            _mail = If(_mail, If(GetProperty("mail"), ""))
            Return _mail
        End Get
        Set(ByVal value As String)
            If SetProperty("mail", value) Then _mail = value

            NotifyPropertyChanged("mail")
        End Set
    End Property

    Public Property title() As String
        Get
            _title = If(_title, If(GetProperty("title"), ""))
            Return _title
        End Get
        Set(ByVal value As String)
            If SetProperty("title", value) Then _title = value

            NotifyPropertyChanged("title")
        End Set
    End Property

    Public Property department() As String
        Get
            _department = If(_department, If(GetProperty("department"), ""))
            Return _department
        End Get
        Set(ByVal value As String)
            If SetProperty("department", value) Then _department = value

            NotifyPropertyChanged("department")
        End Set
    End Property

    Public Property userPrincipalName() As String
        Get
            _userPrincipalName = If(_userPrincipalName, If(GetProperty("userPrincipalName"), ""))
            Return _userPrincipalName
        End Get
        Set(ByVal value As String)
            If SetProperty("userPrincipalName", value) Then _userPrincipalName = value

            NotifyPropertyChanged("userPrincipalName")
        End Set
    End Property

    Public Property userPrincipalNameName() As String
        Get
            Dim s As String
            If userPrincipalName.Contains("@") Then
                s = Split(userPrincipalName, "@")(0)
            ElseIf Len(userPrincipalName) > 0 Then
                s = userPrincipalName
            Else
                s = ""
            End If
            Return s
        End Get
        Set(ByVal value As String)
            userPrincipalName = value & "@" & userPrincipalNameDomain

            NotifyPropertyChanged("userPrincipalNameName")
        End Set
    End Property

    Public Property userPrincipalNameDomain() As String
        Get
            Dim s As String
            If userPrincipalName.Contains("@") Then
                s = Split(userPrincipalName, "@")(1)
            Else
                s = ""
            End If
            Return s
        End Get
        Set(ByVal value As String)
            userPrincipalName = userPrincipalNameName & "@" & value

            NotifyPropertyChanged("userPrincipalNameDomain")
        End Set
    End Property

    Public Property memberOf() As ObservableCollection(Of clsDirectoryObject)
        Get
            If _memberOf Is Nothing Then
                Dim g As Object = GetProperty("memberOf")
                If IsArray(g) Then          'если групп несколько
                    _memberOf = New ObservableCollection(Of clsDirectoryObject)(CType(g, Object()).Select(Function(x As Object) New clsDirectoryObject(New DirectoryEntry("LDAP://" + Domain.Name + "/" + x.ToString, Domain.Username, Domain.Password), Domain)).ToArray)
                ElseIf g Is Nothing Then    'если групп нет
                    _memberOf = New ObservableCollection(Of clsDirectoryObject)
                Else                        'если группа одна
                    _memberOf = New ObservableCollection(Of clsDirectoryObject)({New clsDirectoryObject(New DirectoryEntry("LDAP://" + Domain.Name + "/" + g.ToString, Domain.Username, Domain.Password), Domain)})
                End If
            End If
            Return _memberOf
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _memberOf = value

            NotifyPropertyChanged("memberOf")
        End Set
    End Property

    'Public ReadOnly Property publicDelegates() As Object()
    'Get
    '        If _publicDelegates Is Nothing Then
    '            Dim Value As Object = GetProperty("publicDelegates")
    '            If IsArray(Value) Then
    '                _publicDelegates = Value
    '            ElseIf Value Is Nothing Then
    '                _publicDelegates = Nothing
    '            Else
    '                _publicDelegates = {Value}
    '            End If
    '        End If
    '        Return _publicDelegates
    '    End Get
    'End Property

    Public ReadOnly Property msExchHideFromAddressLists() As Boolean
        Get
            _entry.RefreshCache({"msExchHideFromAddressLists"})
            Return GetProperty("msExchHideFromAddressLists")
        End Get
    End Property

    Public Property thumbnailPhoto() As Byte()
        Get
            _thumbnailPhoto = If(_thumbnailPhoto, GetProperty("thumbnailPhoto"))
            Return _thumbnailPhoto
        End Get
        Set(value As Byte())
            If value Is Nothing Then
                Try
                    _entry.Properties("thumbnailPhoto").Clear()
                    _entry.CommitChanges()
                    _thumbnailPhoto = Nothing
                Catch ex As Exception
                    ThrowException(ex, "Clear thumbnailPhoto")
                End Try
            Else
                Try
                    _entry.Properties("thumbnailPhoto").Clear()
                    _entry.Properties("thumbnailPhoto").Add(value)
                    _entry.CommitChanges()
                    _thumbnailPhoto = value
                Catch ex As Exception
                    ThrowException(ex, "Set thumbnailPhoto")
                End Try
            End If

            NotifyPropertyChanged("thumbnailPhoto")
        End Set
    End Property

    Public Property userWorkstations() As String()
        Get
            If _userWorkstations IsNot Nothing Then
                Return _userWorkstations
            Else
                Dim str As String = GetProperty("userWorkstations")
                _userWorkstations = If(str IsNot Nothing, str.Split({","}, StringSplitOptions.RemoveEmptyEntries), {})
                Return _userWorkstations
            End If
        End Get
        Set(value As String())
            If SetProperty("userWorkstations", Join(value, ",")) Then _userWorkstations = value

            NotifyPropertyChanged("userWorkstations")
        End Set
    End Property

#End Region

#Region "Computer attributes"

    Public ReadOnly Property dNSHostName As String
        Get
            _dnshostname = If(_dnshostname, GetProperty("dNSHostName"))
            Return _dnshostname
        End Get
    End Property

    Public Property location() As String
        Get
            _location = If(_location, If(GetProperty("location"), ""))
            Return _location
        End Get
        Set(ByVal value As String)
            If SetProperty("location", value) Then _location = value

            NotifyPropertyChanged("location")
        End Set
    End Property

    Public ReadOnly Property operatingSystem() As String
        Get
            _operatingsystem = If(_operatingsystem, If(GetProperty("operatingSystem"), ""))
            Return _operatingsystem
        End Get
    End Property

    Public ReadOnly Property operatingSystemVersion() As String
        Get
            _operatingsystemversion = If(_operatingsystemversion, If(GetProperty("operatingSystemVersion"), ""))
            Return _operatingsystemversion
        End Get
    End Property

#End Region

#Region "Group attributes"

    Public Property groupType() As Long
        Get
            _groupType = If(_groupType <> 0, _groupType, GetProperty("groupType"))
            Return _groupType
        End Get
        Set(ByVal value As Long)
            If SetProperty("groupType", value) Then
                _groupType = value
            Else
                ShowWrongMemberMessage()
            End If
            NotifyPropertyChanged("groupType")
            NotifyPropertyChanged("groupTypeScopeDomainLocal")
            NotifyPropertyChanged("groupTypeScopeDomainGlobal")
            NotifyPropertyChanged("groupTypeScopeDomainUniversal")
            NotifyPropertyChanged("groupTypeSecurity")
            NotifyPropertyChanged("groupTypeDistribution")
            NotifyPropertyChanged("Image")
        End Set
    End Property

    Public Property groupTypeScopeDomainLocal() As Boolean
        Get
            Return groupType And ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If Not (gt And ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt + ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If (gt And ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If (gt And ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    Public Property groupTypeScopeDomainGlobal() As Boolean
        Get
            Return groupType And ADS_GROUP_TYPE_GLOBAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If (gt And ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If Not (gt And ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt + ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If (gt And ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    Public Property groupTypeScopeDomainUniversal() As Boolean
        Get
            Return groupType And ADS_GROUP_TYPE_UNIVERSAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If (gt And ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If (gt And ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt - ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If Not (gt And ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt + ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    Public Property groupTypeSecurity() As Boolean
        Get
            Return groupType And ADS_GROUP_TYPE_SECURITY_ENABLED
        End Get
        Set(ByVal value As Boolean)
            If value Then
                If Not (groupType And ADS_GROUP_TYPE_SECURITY_ENABLED) Then
                    groupType = groupType + ADS_GROUP_TYPE_SECURITY_ENABLED
                End If
            Else
                If (groupType And ADS_GROUP_TYPE_SECURITY_ENABLED) Then
                    groupType = groupType - ADS_GROUP_TYPE_SECURITY_ENABLED
                End If
            End If
        End Set
    End Property

    Public ReadOnly Property groupTypeDistribution() As Boolean
        Get
            Return Not groupTypeSecurity
        End Get
    End Property

    Public Property info() As String
        Get
            _info = If(_info, If(GetProperty("info"), ""))
            Return _info
        End Get
        Set(ByVal value As String)
            If SetProperty("info", value) Then _info = value

            NotifyPropertyChanged("info")
        End Set
    End Property

    Public Property member() As ObservableCollection(Of clsDirectoryObject)
        Get
            If _member Is Nothing Then
                Dim o As Object = GetProperty("member")
                If IsArray(o) Then          'если объектов несколько
                    _member = New ObservableCollection(Of clsDirectoryObject)(CType(o, Object()).Select(Function(x As Object) New clsDirectoryObject(New DirectoryEntry("LDAP://" + Domain.Name + "/" + x.ToString, Domain.Username, Domain.Password), Domain)).ToArray)
                ElseIf o Is Nothing Then    'если объектов нет
                    _member = New ObservableCollection(Of clsDirectoryObject)
                Else                        'если объект один
                    _member = New ObservableCollection(Of clsDirectoryObject)({New clsDirectoryObject(New DirectoryEntry("LDAP://" + Domain.Name + "/" + o.ToString, Domain.Username, Domain.Password), Domain)})
                End If
            End If
            Return _member
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _member = value

            NotifyPropertyChanged("member")
        End Set
    End Property

#End Region

#Region "Shared attributes"

    Public Property accountExpires() As Long
        Get
            _accountExpires = If(_accountExpires >= 0, _accountExpires, GetProperty("accountExpires"))
            Return _accountExpires
        End Get
        Set(ByVal value As Long)
            If SetProperty("accountExpires", value.ToString) Then _accountExpires = value

            NotifyPropertyChanged("accountExpires")
            NotifyPropertyChanged("accountExpiresDate")
            NotifyPropertyChanged("accountExpiresFormated")
            NotifyPropertyChanged("accountExpiresFlag")
            NotifyPropertyChanged("accountNeverExpires")
        End Set
    End Property

    Public Property accountExpiresDate() As Date
        Get
            If accountNeverExpires Then
                Return Now
            Else
                Return Date.FromFileTime(accountExpires)
            End If
        End Get
        Set(ByVal value As Date)
            accountExpires = value.ToFileTime

            NotifyPropertyChanged("accountExpiresDate")
        End Set
    End Property

    Public ReadOnly Property accountExpiresFormated() As String
        Get
            If accountNeverExpires Then
                Return "никогда"
            Else
                Return accountExpiresDate.ToString
            End If
        End Get
    End Property

    Public Property accountExpiresFlag() As Boolean
        Get
            Dim e As Long = accountExpires
            If e = 9223372036854775807 Then e = 0
            Return CBool(e)
        End Get
        Set(ByVal value As Boolean)
            accountExpires = If(value, Now.ToFileTime, 0)

            NotifyPropertyChanged("accountExpiresFlag")
        End Set
    End Property

    Public ReadOnly Property accountNeverExpires() As Boolean
        Get
            Dim e As Long = accountExpires
            If e = 9223372036854775807 Then e = 0
            Return Not CBool(e)
        End Get
    End Property

    Public ReadOnly Property badPwdCount() As Integer
        Get
            _badPwdCount = If(_badPwdCount >= 0, _badPwdCount, GetProperty("badPwdCount"))
            Return _badPwdCount
        End Get
    End Property

    Public Property description() As String
        Get
            _description = If(_description, If(GetProperty("description"), ""))
            Return _description
        End Get
        Set(ByVal value As String)
            If SetProperty("description", value) Then _description = value

            NotifyPropertyChanged("description")
        End Set
    End Property

    Public Sub NotifyMoved()
        _distinguishedName = Nothing
        NotifyPropertyChanged("distinguishedName")
        NotifyPropertyChanged("distinguishedNameFormated")
    End Sub

    Public ReadOnly Property distinguishedName() As String
        Get
            _distinguishedName = If(_distinguishedName, If(GetProperty("distinguishedName"), ""))
            Return _distinguishedName
        End Get
    End Property

    Public ReadOnly Property distinguishedNameFormated() As String
        Get
            Dim OU() As String = Split(distinguishedName, ",")
            Dim ResultString As String = ""
            For I As Integer = UBound(OU) To 0 Step -1
                If OU(I).StartsWith("OU") Then
                    ResultString = ResultString & " \ " & Mid(OU(I), 4)
                End If
            Next
            Return Mid(ResultString, 4)
        End Get
    End Property

    Public ReadOnly Property lastLogon() As Long
        Get
            _lastLogon = If(_lastLogon >= 0, _lastLogon, GetProperty("lastLogon"))
            Return _lastLogon
        End Get
    End Property

    Public ReadOnly Property lastLogonDate() As Date
        Get
            Return If(lastLogon <= 0, Nothing, Date.FromFileTime(lastLogon))
        End Get
    End Property

    Public ReadOnly Property lastLogonFormated() As String
        Get
            Return If(lastLogonDate = Nothing, "никогда", lastLogonDate.ToString)
        End Get
    End Property

    Public Property logonCount() As Integer
        Get
            _logonCount = If(_logonCount >= 0, _logonCount, GetProperty("logonCount"))
            Return _logonCount
        End Get
        Set(ByVal value As Integer)
            If SetProperty("logonCount", value) Then _logonCount = value

            NotifyPropertyChanged("logonCount")
        End Set
    End Property

    Public ReadOnly Property name() As String
        Get
            _name = If(_name, If(GetProperty("name"), ""))
            Return _name
        End Get
    End Property

    Public Sub NotifyRenamed()
        _name = Nothing
        NotifyPropertyChanged("name")
    End Sub

    Public ReadOnly Property objectGUID() As String
        Get
            _objectGUID = If(_objectGUID, New Guid(TryCast(GetProperty("objectGUID"), Byte())).ToString)
            Return _objectGUID
        End Get
    End Property

    Public ReadOnly Property objectSID() As String
        Get
            If _objectSID IsNot Nothing Then
                Return _objectSID
            Else
                Try
                    Dim sid As New SecurityIdentifier(TryCast(GetProperty("objectSid"), Byte()), 0)
                    If sid.IsAccountSid Then _objectSID = sid.ToString
                    Return _objectSID
                Catch
                    Return Nothing
                End Try
            End If
        End Get
    End Property

    Public Property pwdLastSet() As Long
        Get
            _pwdLastSet = If(_pwdLastSet >= 0, _pwdLastSet, GetProperty("pwdLastSet"))
            Return _pwdLastSet
        End Get
        Set(ByVal value As Long)
            If SetProperty("pwdLastSet", value.ToString) Then _pwdLastSet = value

            NotifyPropertyChanged("pwdLastSet")
            NotifyPropertyChanged("pwdLastSetDate")
            NotifyPropertyChanged("pwdLastSetFormated")
            NotifyPropertyChanged("passwordExpiresDate")
            NotifyPropertyChanged("passwordExpiresFormated")
            NotifyPropertyChanged("userMustChangePasswordNextLogon")
        End Set
    End Property

    Public ReadOnly Property pwdLastSetDate() As Date
        Get
            Return If(pwdLastSet = 0, Nothing, Date.FromFileTime(pwdLastSet))
        End Get
    End Property

    Public ReadOnly Property pwdLastSetFormated() As String
        Get
            Return If(pwdLastSetDate = Nothing, "никогда", pwdLastSetDate.ToString)
        End Get
    End Property

    Public ReadOnly Property passwordExpiresDate() As Date
        Get
            If passwordNeverExpires Then
                Return Nothing
            Else
                Return If(pwdLastSetDate = Nothing, Nothing, pwdLastSetDate.AddDays(Domain.MaxPwdAge))
            End If
        End Get
    End Property

    Public ReadOnly Property passwordExpiresFormated() As String
        Get
            If passwordNeverExpires Then
                Return "никогда"
            Else
                Return If(pwdLastSetDate = Nothing, "не задан", pwdLastSetDate.AddDays(Domain.MaxPwdAge).ToString)
            End If
        End Get
    End Property

    Public Property userMustChangePasswordNextLogon() As Boolean
        Get
            Return Not (CBool(pwdLastSet))
        End Get
        Set(ByVal value As Boolean)
            If value Then
                pwdLastSet = 0 'password expired
            Else
                pwdLastSet = -1 'password changed today
            End If
        End Set
    End Property

    Public Property sAMAccountName() As String
        Get
            _sAMAccountName = If(_sAMAccountName, If(GetProperty("sAMAccountName"), ""))
            Return _sAMAccountName
        End Get
        Set(ByVal value As String)
            If SetProperty("sAMAccountName", value) Then _sAMAccountName = value

            NotifyPropertyChanged("sAMAccountName")
        End Set
    End Property

    Public Property userAccountControl() As Long
        Get
            _userAccountControl = If(_userAccountControl >= 0, _userAccountControl, GetProperty("userAccountControl"))
            Return _userAccountControl
        End Get
        Set(ByVal value As Long)
            If SetProperty("userAccountControl", value) Then _userAccountControl = value

            NotifyPropertyChanged("userAccountControl")
            NotifyPropertyChanged("normalAccount")
            NotifyPropertyChanged("disabled")
            NotifyPropertyChanged("disabledFormated")
            NotifyPropertyChanged("passwordNeverExpires")
            NotifyPropertyChanged("passwordExpiresDate")
            NotifyPropertyChanged("passwordExpiresFormated")
            NotifyPropertyChanged("Image")
        End Set
    End Property

    Public Property normalAccount() As Boolean
        Get
            Return userAccountControl And ADS_UF_NORMAL_ACCOUNT
        End Get
        Set(ByVal value As Boolean)
            If value Then
                If Not (userAccountControl And ADS_UF_NORMAL_ACCOUNT) Then
                    userAccountControl = userAccountControl + ADS_UF_NORMAL_ACCOUNT
                End If
            Else
                If (userAccountControl And ADS_UF_NORMAL_ACCOUNT) Then
                    userAccountControl = userAccountControl - ADS_UF_NORMAL_ACCOUNT
                End If
            End If
        End Set
    End Property

    Public Property disabled() As Boolean
        Get
            Return userAccountControl And ADS_UF_ACCOUNTDISABLE
        End Get
        Set(ByVal value As Boolean)
            If value Then
                If Not (userAccountControl And ADS_UF_ACCOUNTDISABLE) Then
                    userAccountControl = userAccountControl + ADS_UF_ACCOUNTDISABLE
                End If
            Else
                If (userAccountControl And ADS_UF_ACCOUNTDISABLE) Then
                    userAccountControl = userAccountControl - ADS_UF_ACCOUNTDISABLE
                End If
            End If
            description = String.Format("{0} {1} ({2})", If(value, "Заблокирован", "Разблокирован"), Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString)
        End Set
    End Property

    Public ReadOnly Property disabledFormated() As String
        Get
            If userAccountControl And ADS_UF_ACCOUNTDISABLE Then
                Return "✓"
            Else
                Return ""
            End If
        End Get
    End Property

    Public Property passwordNeverExpires() As Boolean
        Get
            Return userAccountControl And ADS_UF_DONT_EXPIRE_PASSWD
        End Get
        Set(ByVal value As Boolean)
            If value Then
                If Not (userAccountControl And ADS_UF_DONT_EXPIRE_PASSWD) Then
                    userAccountControl = userAccountControl + ADS_UF_DONT_EXPIRE_PASSWD
                End If
            Else
                If (userAccountControl And ADS_UF_DONT_EXPIRE_PASSWD) Then
                    userAccountControl = userAccountControl - ADS_UF_DONT_EXPIRE_PASSWD
                End If
            End If
        End Set
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
            whenChanged = If(_whenChanged = Nothing, GetProperty("whenChanged"), _whenChanged)
            Return _whenChanged
        End Get
    End Property

#End Region

End Class
