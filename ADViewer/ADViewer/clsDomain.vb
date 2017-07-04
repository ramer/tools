Imports Microsoft.Win32
Imports System.DirectoryServices
Imports CredentialManagement
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices.ActiveDirectory
Imports System.Threading.Tasks

<AttributeUsage(AttributeTargets.Property)>
Public Class clsDomain
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _name As String
    Private _username As String
    Private _password As String

    Private _rootdse As DirectoryEntry
    Private _defaultnamingcontext As DirectoryEntry
    Private _configurationnamingcontext As DirectoryEntry
    Private _schemanamingcontext As DirectoryEntry
    Private _properties As New ObservableCollection(Of clsDomainProperty)
    Private _maxpwdage As Integer
    Private _suffixes As New ObservableCollection(Of String)

    Private _searchroot As DirectoryEntry

    Private _usernamepattern As String = ""
    Private _computerpattern As String = ""
    Private _telephonenumberpattern As New ObservableCollection(Of clsTelephoneNumberPattern)
    Private _telephonenumber As New ObservableCollection(Of clsTelephoneNumber)

    Private _defaultpassword As String = ""

    Private _defaultgroups As New ObservableCollection(Of clsDirectoryObject)

    Private _exchangeservers As New ObservableCollection(Of clsDirectoryObject)
    Private _useexchange As Boolean
    Private _exchangeserver As clsDirectoryObject

    Private _attributesmandatory As New ObservableCollection(Of clsAttribute)
    Private _attributesoptional As New ObservableCollection(Of clsAttribute)

    Private _validated As Boolean

    Private Sub NotifyPropertyChanged(propertyName As String)
        OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Sub New(domainname As String)
        _name = domainname

        LoadSettingsFromRegistry()

        Console.WriteLine("")
        For Each p As System.Reflection.PropertyInfo In Me.GetType().GetProperties()
            If p.CanRead Then
                Console.WriteLine("{0}: {1}", p.Name, p.GetValue(Me, Nothing))
            End If
        Next

    End Sub

    Private Async Function Validate() As Task(Of Boolean)
        Validated = False
        If Len(Name) = 0 Or Len(Username) = 0 Or Len(Password) = 0 Then Return False

        Dim connectionstring As String = "LDAP://" & Name & "/"
        Dim newDefaultNamingContext As String = ""
        Dim success As Boolean = False

        'RootDSE
        success = Await Task.Run(
            Function()
                Try
                    If _rootdse Is Nothing Then _rootdse = New DirectoryEntry(connectionstring & "RootDSE", Username, Password)
                    If _rootdse.Properties.Count = 0 Then Return False
                Catch
                    _rootdse = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False
        NotifyPropertyChanged("RootDSE")


        'defaultNamingContext
        success = Await Task.Run(
            Function()
                Try
                    If _defaultnamingcontext Is Nothing Then _defaultnamingcontext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "defaultNamingContext"), Username, Password)
                    If _defaultnamingcontext.Properties.Count = 0 Then Return False
                Catch
                    _defaultnamingcontext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False
        NotifyPropertyChanged("DefaultNamingContext")

        'configurationNamingContext
        success = Await Task.Run(
            Function()
                Try
                    If _configurationnamingcontext Is Nothing Then _configurationnamingcontext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "configurationNamingContext"), Username, Password)
                    If _configurationnamingcontext.Properties.Count = 0 Then Return False
                Catch
                    _configurationnamingcontext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False
        NotifyPropertyChanged("ConfigurationNamingContext")

        'schemaNamingContext
        success = Await Task.Run(
            Function()
                Try
                    If _schemanamingcontext Is Nothing Then _schemanamingcontext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "schemaNamingContext"), Username, Password)
                    If _schemanamingcontext.Properties.Count = 0 Then Return False
                Catch ex As Exception
                    _schemanamingcontext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False
        NotifyPropertyChanged("SchemaNamingContext")

        'properties
        _properties = Await Task.Run(
            Function()
                Dim p As New ObservableCollection(Of clsDomainProperty)
                Try
                    p.Clear()
                    p.Add(New clsDomainProperty("Пороговое значение блокировки", String.Format("{0} ошибок входа", GetLDAPProperty(_defaultnamingcontext.Properties, "lockoutThreshold"))))
                    p.Add(New clsDomainProperty("Время до сброса счетчика блокировки", String.Format("{0} минут", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("lockoutDuration")(0))).Minutes)))
                    p.Add(New clsDomainProperty("Продолжительность блокировки учетной записи", String.Format("{0} минут", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("lockOutObservationWindow")(0))).Minutes)))
                    p.Add(New clsDomainProperty("Максимальный срок действия пароля", String.Format("{0} дней", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("maxPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty("Минимальный срок действия пароля", String.Format("{0} дней", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("minPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty("Минимальная длина пароля", String.Format("{0} символов", GetLDAPProperty(_defaultnamingcontext.Properties, "minPwdLength") & " симв.")))
                    p.Add(New clsDomainProperty("Пароль должен отвечать требованиям сложности", String.Format("{0}", GetLDAPProperty(_defaultnamingcontext.Properties, "pwdProperties"))))
                    p.Add(New clsDomainProperty("Вести журнал паролей", String.Format("{0} сохраненных паролей", GetLDAPProperty(_defaultnamingcontext.Properties, "pwdHistoryLength"))))
                Catch
                    p.Clear()
                End Try
                Return p
            End Function)
        NotifyPropertyChanged("Properties")

        'maximum password age
        Await Task.Run(
            Sub()
                Try
                    If _maxpwdage = 0 Then _maxpwdage = -TimeSpan.FromTicks(LongFromLargeInteger(GetLDAPProperty(_defaultnamingcontext.Properties, "maxPwdAge"))).Days
                Catch ex As Exception
                    _maxpwdage = 0
                End Try
            End Sub)
        NotifyPropertyChanged("MaxPwdAge")

        'domain suffixes
        If _suffixes.Count = 0 Then _suffixes = Await Task.Run(
            Function()
                Dim s As New ObservableCollection(Of String)
                Try
                    Dim LDAPsearcher As New DirectorySearcher(ConfigurationNamingContext)
                    Dim LDAPresults As SearchResultCollection = Nothing

                    LDAPsearcher.Filter = "(&(objectClass=crossRef)(systemFlags=3))"
                    LDAPresults = LDAPsearcher.FindAll()
                    For Each LDAPresult As SearchResult In LDAPresults
                        s.Add(LCase(GetLDAPProperty(LDAPresult.Properties, "dnsRoot")))
                    Next LDAPresult
                Catch ex As Exception
                    s.Clear()
                End Try
                Return s
            End Function)
        NotifyPropertyChanged("Suffixes")

        'search root
        Await Task.Run(
            Sub()
                Try
                    If _searchroot Is Nothing AndAlso _defaultnamingcontext IsNot Nothing Then _searchroot = _defaultnamingcontext
                Catch ex As Exception
                    _searchroot = Nothing
                End Try
            End Sub)
        NotifyPropertyChanged("SearchRoot")
        NotifyPropertyChanged("SearchRootString")

        'exchange servers
        If _exchangeservers.Count = 0 Then _exchangeservers = Await Task.Run(
            Function()
                Dim e As New ObservableCollection(Of clsDirectoryObject)
                Try
                    Dim LDAPsearcher As New DirectorySearcher(_configurationnamingcontext)
                    Dim LDAPresults As SearchResultCollection = Nothing
                    Dim LDAPresult As SearchResult

                    LDAPsearcher.Filter = "(objectClass=msExchExchangeServer)"
                    LDAPresults = LDAPsearcher.FindAll()
                    For Each LDAPresult In LDAPresults
                        e.Add(New clsDirectoryObject(LDAPresult.GetDirectoryEntry, Me))
                    Next LDAPresult

                    _useexchange = False
                    _exchangeserver = Nothing
                Catch ex As Exception
                    e.Clear()
                End Try
                Return e
            End Function)
        NotifyPropertyChanged("ExchangeServers")
        NotifyPropertyChanged("UseExchange")
        NotifyPropertyChanged("ExchangeServer")

        Validated = True
        SaveSettingsToRegistry()
        Return True
    End Function

    Public Async Function Revalidate() As Task(Of Boolean)
        _rootdse = Nothing
        _defaultnamingcontext = Nothing
        _configurationnamingcontext = Nothing
        _schemanamingcontext = Nothing
        _properties.Clear()
        _maxpwdage = 0
        _suffixes.Clear()
        _searchroot = Nothing
        _exchangeservers.Clear()

        Return Await Validate()
    End Function

    Public Async Sub LoadSettingsFromRegistry()
        Try
            Dim regDomain As RegistryKey = regDomains.OpenSubKey(_name)
            With regDomain
                Dim cred As New Credential("", "", "ADViewer: " & _name, CredentialType.Generic)
                cred.PersistanceType = PersistanceType.Enterprise
                cred.Load()
                _username = cred.Username
                _password = cred.Password

                _rootdse = If(.GetValue("RootDSE", Nothing) IsNot Nothing, New DirectoryEntry(.GetValue("RootDSE"), _username, _password), Nothing)
                _defaultnamingcontext = If(.GetValue("DefaultNamingContext", Nothing) IsNot Nothing, New DirectoryEntry(.GetValue("DefaultNamingContext"), _username, _password), Nothing)
                _configurationnamingcontext = If(.GetValue("ConfigurationNamingContext", Nothing) IsNot Nothing, New DirectoryEntry(.GetValue("ConfigurationNamingContext"), _username, _password), Nothing)
                _schemanamingcontext = If(.GetValue("SchemaNamingContext", Nothing) IsNot Nothing, New DirectoryEntry(.GetValue("SchemaNamingContext"), _username, _password), Nothing)
                _maxpwdage = .GetValue("MaxPwdAge", 0)

                Dim ADViewerSuffixesRegPath As RegistryKey = regDomain.OpenSubKey("Suffixes")
                If ADViewerSuffixesRegPath IsNot Nothing Then
                    With ADViewerSuffixesRegPath
                        For Each s As String In .GetValueNames
                            _suffixes.Add(s)
                        Next
                    End With
                End If

                _searchroot = If(.GetValue("SearchRoot", Nothing) IsNot Nothing, New DirectoryEntry(.GetValue("SearchRoot"), _username, _password), Nothing)

                _usernamepattern = .GetValue("UserNamePattern", "")
                _computerpattern = .GetValue("ComputerPattern", "")

                Dim ADViewerTelephoneNumberPatternRegPath As RegistryKey = regDomain.OpenSubKey("TelephoneNumberPattern")
                If ADViewerTelephoneNumberPatternRegPath IsNot Nothing Then
                    Dim tnpl = New ObservableCollection(Of clsTelephoneNumberPattern)
                    For Each lbl As String In ADViewerTelephoneNumberPatternRegPath.GetSubKeyNames()
                        Dim tnp As New clsTelephoneNumberPattern

                        Dim reg As RegistryKey = ADViewerTelephoneNumberPatternRegPath.OpenSubKey(lbl)
                        tnp.Label = lbl
                        tnp.Pattern = reg.GetValue("Pattern", "")
                        tnp.Range = reg.GetValue("Range", "")

                        tnpl.Add(tnp)
                    Next
                    _telephonenumberpattern = tnpl
                End If

                _defaultpassword = .GetValue("DefaultPassword", "")

                Dim ADViewerDefaultGroupsRegPath As RegistryKey = regDomain.OpenSubKey("DefaultGroups")
                If ADViewerDefaultGroupsRegPath IsNot Nothing Then
                    With ADViewerDefaultGroupsRegPath
                        For Each g As String In .GetValueNames
                            If .GetValue(g, Nothing) IsNot Nothing Then _defaultgroups.Add(New clsDirectoryObject(New DirectoryEntry(.GetValue(g), _username, _password), Me))
                        Next
                    End With
                End If

                Dim ADViewerExchangeServersRegPath As RegistryKey = regDomain.OpenSubKey("ExchangeServers")
                If ADViewerExchangeServersRegPath IsNot Nothing Then
                    With ADViewerExchangeServersRegPath
                        For Each es As String In .GetValueNames
                            If .GetValue(es, Nothing) IsNot Nothing Then _exchangeservers.Add(New clsDirectoryObject(New DirectoryEntry(.GetValue(es), _username, _password), Me))
                        Next
                    End With
                End If
                _useexchange = .GetValue("UseExchange", False)

                For Each es In _exchangeservers
                    If .GetValue("ExchangeServer", Nothing) IsNot Nothing AndAlso es.Entry.Path = .GetValue("ExchangeServer") Then _exchangeserver = es
                Next

            End With

            Await Validate()
        Catch ex As Exception
            ThrowException(ex, "clsDomain.LoadSettingsFromRegistry")
        End Try
    End Sub

    Public Function SaveSettingsToRegistry() As Boolean
        Try
            If _name = "" Then Return False
            regDomains.DeleteSubKeyTree(UCase(_name), False)
            Dim regDomain As RegistryKey = regDomains.CreateSubKey(UCase(_name))
            With regDomain

                Dim cred As New Credential("", "", "ADViewer: " & _name, CredentialType.Generic)
                cred.PersistanceType = PersistanceType.Enterprise
                cred.Username = _username
                cred.Password = _password
                cred.Save()

                .SetValue("RootDSE", If(_rootdse IsNot Nothing, _rootdse.Path, ""))
                .SetValue("DefaultNamingContext", If(_defaultnamingcontext IsNot Nothing, _defaultnamingcontext.Path, ""))
                .SetValue("ConfigurationNamingContext", If(_configurationnamingcontext IsNot Nothing, _configurationnamingcontext.Path, ""))
                .SetValue("SchemaNamingContext", If(_schemanamingcontext IsNot Nothing, _schemanamingcontext.Path, ""))
                .SetValue("MaxPwdAge", _maxpwdage)

                Dim ADViewerSuffixesRegPath As RegistryKey = regDomain.CreateSubKey("Suffixes")
                With ADViewerSuffixesRegPath
                    For Each s As String In _suffixes
                        .SetValue(s, "")
                    Next
                End With

                .SetValue("SearchRoot", If(_searchroot IsNot Nothing, _searchroot.Path, ""))

                .SetValue("UserNamePattern", _usernamepattern)
                .SetValue("ComputerPattern", _computerpattern)

                Dim ADViewerTelephoneNumberPatternRegPath As RegistryKey = regDomain.CreateSubKey("TelephoneNumberPattern")
                For Each tn As clsTelephoneNumberPattern In _telephonenumberpattern
                    If tn.Label Is Nothing Then Continue For
                    Dim reg As RegistryKey = ADViewerTelephoneNumberPatternRegPath.CreateSubKey(tn.Label)
                    reg.SetValue("Pattern", If(tn.Pattern, ""), RegistryValueKind.String)
                    reg.SetValue("Range", If(tn.Range, ""), RegistryValueKind.String)
                Next

                .SetValue("DefaultPassword", _defaultpassword)

                Dim ADViewerDefaultGroupsRegPath As RegistryKey = regDomain.CreateSubKey("DefaultGroups")
                With ADViewerDefaultGroupsRegPath
                    For Each g As clsDirectoryObject In _defaultgroups
                        If g IsNot Nothing AndAlso g.Entry IsNot Nothing Then .SetValue(g.name, g.Entry.Path)
                    Next
                End With

                Dim ADViewerExchangeServersRegPath As RegistryKey = regDomain.CreateSubKey("ExchangeServers")
                With ADViewerExchangeServersRegPath
                    For Each es As clsDirectoryObject In _exchangeservers
                        If es IsNot Nothing AndAlso es.Entry IsNot Nothing Then .SetValue(es.name, es.Entry.Path)
                    Next
                End With

                .SetValue("UseExchange", _useexchange)

                If _exchangeserver IsNot Nothing AndAlso _exchangeserver.Entry IsNot Nothing Then .SetValue("ExchangeServer", _exchangeserver.Entry.Path)

            End With


            Return True
        Catch ex As Exception
            ThrowException(ex, "clsDomain.SaveSettingsToRegistry")
            Return False
        End Try
    End Function

    Public Function DeleteSettingsFromRegistry() As Boolean
        Try
            If _name = "" Then Return False
            regDomains.DeleteSubKeyTree(UCase(Name))
            Return True
        Catch ex As Exception
            ThrowException(ex, "clsDomain.DeleteSettingsFromRegistry")
            Return False
        End Try
    End Function

    Public Sub GetNextTelephoneNumber()
        _telephonenumber = GetNextDomainTelephoneNumbers(Me)
    End Sub

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
            NotifyPropertyChanged("Name")
            Validated = False
        End Set
    End Property

    Public Property Username() As String
        Get
            Return _username
        End Get
        Set(value As String)
            _username = value
            NotifyPropertyChanged("Username")
            Validated = False
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _password
        End Get
        Set(value As String)
            _password = value
            NotifyPropertyChanged("Password")
            Validated = False
        End Set
    End Property

    Public ReadOnly Property RootDSE() As DirectoryEntry
        Get
            Return _rootdse
        End Get
    End Property

    Public ReadOnly Property DefaultNamingContext() As DirectoryEntry
        Get
            Return _defaultnamingcontext
        End Get
    End Property

    Public ReadOnly Property ConfigurationNamingContext() As DirectoryEntry
        Get
            Return _configurationnamingcontext
        End Get
    End Property

    Public ReadOnly Property SchemaNamingContext() As DirectoryEntry
        Get
            Return _schemanamingcontext
        End Get
    End Property

    Public ReadOnly Property Properties() As ObservableCollection(Of clsDomainProperty)
        Get
            Return _properties
        End Get
    End Property

    Public ReadOnly Property MaxPwdAge() As Integer
        Get
            Return _maxpwdage
        End Get
    End Property

    Public ReadOnly Property Suffixes As ObservableCollection(Of String)
        Get
            Return _suffixes
        End Get
    End Property

    Public Property SearchRoot() As DirectoryEntry
        Get
            Return _searchroot
        End Get
        Set(value As DirectoryEntry)
            _searchroot = value
            NotifyPropertyChanged("SearchRoot")
            NotifyPropertyChanged("SearchRootString")
        End Set
    End Property

    Public ReadOnly Property SearchRootString() As String
        Get
            Return If(_searchroot IsNot Nothing, _searchroot.Path, "")
        End Get
    End Property

    Public Property UsernamePattern() As String
        Get
            Return _usernamepattern
        End Get
        Set(value As String)
            _usernamepattern = value
            NotifyPropertyChanged("UsernamePattern")
        End Set
    End Property

    Public Property ComputerPattern() As String
        Get
            Return _computerpattern
        End Get
        Set(value As String)
            _computerpattern = value
            NotifyPropertyChanged("ComputerPattern")
        End Set
    End Property

    Public Property TelephoneNumberPattern() As ObservableCollection(Of clsTelephoneNumberPattern)
        Get
            Return _telephonenumberpattern
        End Get
        Set(value As ObservableCollection(Of clsTelephoneNumberPattern))
            _telephonenumberpattern = value
            NotifyPropertyChanged("TelephoneNumberPattern")
        End Set
    End Property

    Public ReadOnly Property TelephoneNumber() As ObservableCollection(Of clsTelephoneNumber)
        Get
            Return _telephonenumber
        End Get
    End Property

    Public Property DefaultPassword() As String
        Get
            Return _defaultpassword
        End Get
        Set(value As String)
            _defaultpassword = value
            NotifyPropertyChanged("DefaultPassword")
        End Set
    End Property

    Public Property DefaultGroups() As ObservableCollection(Of clsDirectoryObject)
        Get
            Return _defaultgroups
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _defaultgroups = value
            NotifyPropertyChanged("DefaultGroups")
        End Set
    End Property

    Public ReadOnly Property ExchangeServers() As ObservableCollection(Of clsDirectoryObject)
        Get
            Return _exchangeservers
        End Get
    End Property

    Public Property UseExchange() As Boolean
        Get
            Return _useexchange
        End Get
        Set(value As Boolean)
            _useexchange = value
            NotifyPropertyChanged("UseExchange")
        End Set
    End Property

    Public Property ExchangeServer() As clsDirectoryObject
        Get
            Return _exchangeserver
        End Get
        Set(value As clsDirectoryObject)
            _exchangeserver = value
            NotifyPropertyChanged("ExchangeServer")
        End Set
    End Property

    Public Property Validated() As Boolean
        Get
            Return _validated
        End Get
        Set(value As Boolean)
            _validated = value
            NotifyPropertyChanged("Validated")
            NotifyPropertyChanged("ValidatedImage")
        End Set
    End Property

    Public ReadOnly Property ValidatedImage() As BitmapImage
        Get
            Return If(Not _validated, New BitmapImage(New Uri("pack://application:,,,/" & "img/warning.ico")), Nothing)
        End Get
    End Property

End Class

Public Class clsDomainProperty
    Dim _property As String
    Dim _value As String

    Sub New(prop As String, value As String)
        _property = prop
        _value = value
    End Sub

    Public ReadOnly Property Prop() As String
        Get
            Return _property
        End Get
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return _value
        End Get
    End Property
End Class