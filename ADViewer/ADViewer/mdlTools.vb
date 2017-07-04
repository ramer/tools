Imports System.Collections.ObjectModel
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Reflection
Imports System.Security
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Markup
Imports System.Xml
Imports Microsoft.Win32

Module mdlTools

    Public Const ADS_UF_SCRIPT = 1 '0x1
    Public Const ADS_UF_ACCOUNTDISABLE = 2 '0x2

    Public Const ADS_UF_HOMEDIR_REQUIRED = 8 '0x8
    Public Const ADS_UF_LOCKOUT = 16 '0x10
    Public Const ADS_UF_PASSWD_NOTREQD = 32 '0x20
    Public Const ADS_UF_PASSWD_CANT_CHANGE = 64 '0x40
    Public Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128 '0x80
    Public Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = 256 '0x100
    Public Const ADS_UF_NORMAL_ACCOUNT = 512 '0x200
    Public Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 2048 '0x800
    Public Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = 4096 '0x1000
    Public Const ADS_UF_SERVER_TRUST_ACCOUNT = 8192 '0x2000
    Public Const ADS_UF_DONT_EXPIRE_PASSWD = 65536 '0x10000
    Public Const ADS_UF_MNS_LOGON_ACCOUNT = 131072 '0x20000
    Public Const ADS_UF_SMARTCARD_REQUIRED = 262144 '0x40000
    Public Const ADS_UF_TRUSTED_FOR_DELEGATION = 524288 '0x80000
    Public Const ADS_UF_NOT_DELEGATED = 1048576 '0x100000
    Public Const ADS_UF_USE_DES_KEY_ONLY = 2097152 '0x200000
    Public Const ADS_UF_DONT_REQUIRE_PREAUTH = 4194304 '0x400000
    Public Const ADS_UF_PASSWORD_EXPIRED = 8388608 '0x800000
    Public Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216 '0x1000000

    Public Const ADS_GROUP_TYPE_GLOBAL_GROUP = 2 '0x00000002
    Public Const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 4 '0x00000004
    Public Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = 8 '0x00000008
    Public Const ADS_GROUP_TYPE_SECURITY_ENABLED = -2147483648 '0x80000000

    Public regDomains As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\ADViewerWPF\Domains")
    Public regInterface As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\ADViewerWPF\Interface")
    Public regSettings As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\ADViewerWPF\Settings")
    Public regExternalSoftware As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\ADViewerWPF\ExternalSoftware")

    Public preferences As clsPreferences
    Public domains As New ObservableCollection(Of clsDomain)
    Public sip As New clsSIP

    Public attributesExtended As New ObservableCollection(Of clsAttribute)
    Public attributesDefault As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("accountExpires", "Объект истекает", True)},
        {New clsAttribute("accountExpiresFormated", "Объект истекает (формат)", True)},
        {New clsAttribute("badPwdCount", "Ошибок ввода пароля", True)},
        {New clsAttribute("company", "Компания", True)},
        {New clsAttribute("department", "Подразделение", True)},
        {New clsAttribute("description", "Описание", True)},
        {New clsAttribute("disabled", "Заблокирован", True)},
        {New clsAttribute("disabledFormated", "Заблокирован (формат)", True)},
        {New clsAttribute("displayName", "Отображаемое имя", True)},
        {New clsAttribute("distinguishedName", "LDAP-путь", True)},
        {New clsAttribute("distinguishedNameFormated", "LDAP-путь (формат)", True)},
        {New clsAttribute("givenName", "Имя", True)},
        {New clsAttribute("Image", "⬕", True)},
        {New clsAttribute("initials", "Инициалы", True)},
        {New clsAttribute("lastLogonDate", "Последний вход", True)},
        {New clsAttribute("lastLogonFormated", "Последний вход (формат)", True)},
        {New clsAttribute("location", "Местонахождение", True)},
        {New clsAttribute("logonCount", "Входов", True)},
        {New clsAttribute("mail", "Основной адрес", True)},
        {New clsAttribute("manager", "Руководитель", True)},
        {New clsAttribute("name", "Имя объекта", True)},
        {New clsAttribute("objectGUID", "Уникальный идентификатор (GUID)", True)},
        {New clsAttribute("objectSID", "Уникальный идентификатор (SID)", True)},
        {New clsAttribute("passwordExpiresDate", "Пароль истекает", True)},
        {New clsAttribute("passwordExpiresFormated", "Пароль истекает (формат)", True)},
        {New clsAttribute("physicalDeliveryOfficeName", "Офис", True)},
        {New clsAttribute("pwdLastSetDate", "Дата смены пароля", True)},
        {New clsAttribute("pwdLastSetFormated", "Дата смены пароля (формат)", True)},
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)", True)},
        {New clsAttribute("SchemaClassName", "Класс", True)},
        {New clsAttribute("sn", "Фамилия", True)},
        {New clsAttribute("Status", "Статус", True)},
        {New clsAttribute("StatusFormated", "Статус (формат)", True)},
        {New clsAttribute("telephoneNumber", "Телефон", True)},
        {New clsAttribute("thumbnailPhoto", "Фото", True)},
        {New clsAttribute("title", "Должность", True)},
        {New clsAttribute("userPrincipalName", "Имя входа", True)},
        {New clsAttribute("whenCreated", "Создан", True)},
        {New clsAttribute("whenCreatedFormated", "Создан (формат)", True)}
    }
    Public attributesForSearchDefault As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("displayName", "Отображаемое имя", True)},
        {New clsAttribute("mail", "Основной адрес", True)},
        {New clsAttribute("name", "Имя объекта", True)},
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)", True)},
        {New clsAttribute("userPrincipalName", "Имя входа", True)}
    }

    Public portlistDefault As New Dictionary(Of Integer, String) From {{135, "RPC"},
                                                                       {139, "NETBIOS-SSN"},
                                                                       {445, "SMB over TCP"},
                                                                       {3389, "RDP"},
                                                                       {4899, "Radmin"},
                                                                       {5900, "VNC"},
                                                                       {6129, "DameWare RC"}}

    Public protocols As New Dictionary(Of LumiSoft.Net.BindInfoProtocol, String) From {{LumiSoft.Net.BindInfoProtocol.TCP, "TCP"},
                                                                                      {LumiSoft.Net.BindInfoProtocol.UDP, "UDP"}}

    Public Sub initializePreferences()
        preferences = New clsPreferences
    End Sub

    Public Sub initializeDomains()
        domains.Clear()

        For Each domainname As String In regDomains.GetSubKeyNames()
            domains.Add(New clsDomain(domainname))
        Next
    End Sub

    Public Sub initializeSIP()
        If preferences.SipUse Then sip.Register()
    End Sub

    Public Sub deinitializePreferences()
        preferences.SavePreferences()
    End Sub

    Public Sub deinitializeSIP()
        sip.Unregister()
    End Sub

    Public Function IMsgBox(Promt As String, Optional Title As String = "", Optional Buttons As MsgBoxStyle = vbOKOnly, Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly) As MessageBoxResult
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._content = Promt
        w._title = Title
        w._buttons = Buttons
        w._icon = Icon
        If w.ShowDialog() Then
            Return w._msgboxresult
        Else
            Return vbCancel
        End If
    End Function

    Public Function IInputBox(Promt As String, Optional Title As String = "", Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional DefaultAnswer As String = "") As String
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._inputbox = True
        w._content = Promt
        w._title = Title
        w._buttons = vbOKCancel
        w._icon = Icon
        w._defaultanswer = DefaultAnswer
        If w.ShowDialog() Then
            Return w.tbInput.Text
        Else
            Return Nothing
        End If
    End Function

    Public Function IPasswordBox(Promt As String, Optional Title As String = "", Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional DefaultAnswer As String = "") As String
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._passwordbox = True
        w._content = Promt
        w._title = Title
        w._buttons = vbOKCancel
        w._icon = Icon
        If w.ShowDialog() Then
            Return w.tbInput.Text
        Else
            Return Nothing
        End If
    End Function

    Public Function GetLDAPProperty(ByRef Properties As DirectoryServices.ResultPropertyCollection, ByVal Prop As String)
        Try
            If Properties(Prop).Count > 0 Then
                Return Properties(Prop)(0)
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function GetLDAPProperty(ByRef Properties As DirectoryServices.PropertyCollection, ByVal Prop As String)
        Try
            If Properties(Prop).Count > 0 Then
                Return Properties(Prop)(0)
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function LongFromLargeInteger(largeInteger As Object) As Long
        Dim valBytes(7) As Byte
        Dim result As Long
        Dim type As System.Type = largeInteger.[GetType]()
        Dim highPart As Integer = CInt(type.InvokeMember("HighPart", BindingFlags.GetProperty, Nothing, largeInteger, Nothing))
        Dim lowPart As Integer = CInt(type.InvokeMember("LowPart", BindingFlags.GetProperty, Nothing, largeInteger, Nothing))
        BitConverter.GetBytes(lowPart).CopyTo(valBytes, 0)
        BitConverter.GetBytes(highPart).CopyTo(valBytes, 4)

        result = BitConverter.ToInt64(valBytes, 0)
        If result = 9223372036854775807 Then result = 0

        Return result
    End Function

    Public Sub UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim ex As Exception = DirectCast(e.ExceptionObject, Exception)
        ThrowException(ex, "Необработанное исключение")
    End Sub

    Public Sub ThrowException(ByVal ex As Exception, ByVal Procedure As String)
        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog(Procedure,, ex))
    End Sub

    Public Sub ThrowCustomException(Message As String)
        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog(Message))
    End Sub

    Public Sub ThrowInformation(Message As String)
        With SingleInstanceApplication.nicon
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = ProductName()
            .BalloonTipText = Message
            .Tag = Nothing
            .Visible = False
            .Visible = True
            .ShowBalloonTip(5000)
        End With
    End Sub

    Public Sub ThrowSIPInformation(request As LumiSoft.Net.SIP.Stack.SIP_Request)
        Dim displayname As String = Encoding.UTF8.GetString(Encoding.GetEncoding(1251).GetBytes(request.From.Address.DisplayName))
        Dim uri() As String = request.From.Address.Uri.Value.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
        Dim message As String = ""

        message &= displayname
        If uri.Count > 0 Then message &= " (" & uri(0) & ")"
        message &= vbCrLf & vbCrLf & "показать подробнее?"

        With SingleInstanceApplication.nicon
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = "Входящий вызов"
            .BalloonTipText = message
            .Tag = request
            .Visible = False
            .Visible = True
            .ShowBalloonTip(10000)
        End With
    End Sub

    Public Function GetUserPartFromExchangeUsername(str) As String
        Dim arr As String() = str.Split({"\", "/"}, StringSplitOptions.RemoveEmptyEntries)
        Return If(arr.Length > 1, arr(arr.Length - 1), Nothing)
    End Function

    Public Function ProductName() As String
        If Windows.Application.ResourceAssembly Is Nothing Then
            Return Nothing
        End If

        Return Windows.Application.ResourceAssembly.GetName().Name
    End Function


    Public Function ExtractBytesFromString(str As String) As Long
        '2.011 GB (2,159,225,856 bytes)
        Try
            Return Long.Parse(str.Split({"(", ")"}, StringSplitOptions.RemoveEmptyEntries)(1).Replace(",", "").Split({" "}, StringSplitOptions.RemoveEmptyEntries)(0))
        Catch
            Return 0
        End Try
    End Function

    Public Function CountWords(str As String) As Integer
        Dim j As Integer = 0
        For i = 0 To str.Length - 1
            If str.Chars(i) = " " Then
                j += 1
            End If
        Next
        Return j
    End Function

    Public Function GetAttributesExtended() As clsAttribute()
        Dim attributes As New Dictionary(Of String, clsAttribute)

        For Each domain In domains
            Try
                Dim _domaincontrollers As DirectoryEntry = domain.DefaultNamingContext.Children.Find("OU=Domain Controllers")
                Dim _directorycontext As New DirectoryContext(DirectoryContextType.DirectoryServer, GetLDAPProperty(_domaincontrollers.Children(0).Properties, "dNSHostName"), domain.Username, domain.Password)
                Dim _schema As ActiveDirectorySchema = ActiveDirectorySchema.GetSchema(_directorycontext)
                Dim _userClass As ActiveDirectorySchemaClass = _schema.FindClass("user")

                For Each a As clsAttribute In _userClass.MandatoryProperties.Cast(Of ActiveDirectorySchemaProperty).Where(Function(attr As ActiveDirectorySchemaProperty) attr.IsSingleValued).Select(Function(attr As ActiveDirectorySchemaProperty) New clsAttribute(attr.Name, attr.CommonName)).ToArray
                    If Not attributes.ContainsKey(a.Name) Then attributes.Add(a.Name, a)
                Next
                For Each a As clsAttribute In _userClass.OptionalProperties.Cast(Of ActiveDirectorySchemaProperty).Where(Function(attr As ActiveDirectorySchemaProperty) attr.IsSingleValued).Select(Function(attr As ActiveDirectorySchemaProperty) New clsAttribute(attr.Name, attr.CommonName)).ToArray
                    If Not attributes.ContainsKey(a.Name) Then attributes.Add(a.Name, a)
                Next

            Catch ex As Exception

            End Try
        Next

        Return attributes.Values.ToArray.OrderBy(Function(x As clsAttribute) x.Label).ToArray
    End Function

    Public Function CreateColumn(columninfo As clsDataGridColumnInfo) As DataGridTemplateColumn
        Dim BasicProperties As PropertyInfo() = GetType(clsDirectoryObject).GetProperties()
        Dim BasicPropertiesNames As String() = BasicProperties.Select(Function(x As PropertyInfo) x.Name).ToArray

        Dim column As New DataGridTemplateColumn()
        column.Header = columninfo.Header
        column.SetValue(DataGridColumn.CanUserSortProperty, True)
        If columninfo.DisplayIndex > 0 Then column.DisplayIndex = columninfo.DisplayIndex
        If columninfo.Width > 0 Then column.Width = columninfo.Width
        Dim panel As New FrameworkElementFactory(GetType(VirtualizingStackPanel))
        panel.SetValue(VirtualizingStackPanel.VerticalAlignmentProperty, VerticalAlignment.Center)
        panel.SetValue(VirtualizingStackPanel.MarginProperty, New Thickness(5, 0, 5, 0))
        Dim first As Boolean = True
        For Each attr As clsAttribute In columninfo.Attributes
            Dim bind As System.Windows.Data.Binding

            If BasicPropertiesNames.Contains(attr.Name) Then
                bind = New System.Windows.Data.Binding(attr.Name)
            Else
                bind = New System.Windows.Data.Binding("CustomProperty[" & attr.Name & "]")
            End If
            bind.Mode = BindingMode.OneWay

            If attr.Name <> "Image" Then

                Dim text As New FrameworkElementFactory(GetType(TextBlock))
                If first Then
                    text.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold)
                    first = False
                    column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
                End If
                text.SetBinding(TextBlock.TextProperty, bind)
                text.SetValue(TextBlock.ToolTipProperty, attr.Label)
                panel.AppendChild(text)

            Else

                Dim ttbind As New System.Windows.Data.Binding("Status")
                ttbind.Mode = BindingMode.OneWay
                Dim img As New FrameworkElementFactory(GetType(Image))
                column.SetValue(clsSorter.PropertyNameProperty, "Image")
                img.SetBinding(Image.SourceProperty, bind)
                img.SetValue(Image.WidthProperty, 32.0)
                img.SetValue(Image.HeightProperty, 32.0)
                img.SetBinding(Image.ToolTipProperty, ttbind)
                panel.AppendChild(img)

            End If
            'Status
        Next

        Dim template As New DataTemplate()
        template.VisualTree = panel

        column.CellTemplate = template

        Return column
    End Function

    Public Function StringToLong(str) As Long
        Dim lng As Double
        If Not Double.TryParse(str, lng) Then lng = 0
        Return Int(lng)
    End Function

    Public Function BytesToString(byteCount As Long) As String
        Dim suf As String() = {" б", " Кб", " Мб", " Гб", " Тб", " Пб", " Эб"}
        If byteCount = 0 Then
            Return "0" + suf(0)
        End If
        Dim bytes As Long = Math.Abs(byteCount)
        Dim place As Integer = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)))
        Dim num As Double = Math.Round(bytes / Math.Pow(1024, place), 1)
        Return (Math.Sign(byteCount) * num).ToString() + suf(place)
    End Function

    Public Function Clone(obj As Object) As Object
        Return XamlReader.Load(XmlTextReader.Create(New StringReader(XamlWriter.Save(obj)), New XmlReaderSettings()))
    End Function

    Public Function GetNextDomainUsers(domain As clsDomain) As List(Of String)
        If domain Is Nothing Then Return Nothing
        Dim patterns() As String = domain.UsernamePattern.Split({",", vbCr, vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) Trim(x)).ToArray
        If patterns.Count = 0 Then Return Nothing
        Dim de As DirectoryEntry = domain.DefaultNamingContext
        If de Is Nothing Then Return Nothing

        Dim result As New List(Of String)

        Dim LDAPsearcher As New DirectorySearcher(de)
        Dim LDAPresults As SearchResultCollection = Nothing
        Dim LDAPresult As SearchResult

        LDAPsearcher.PropertiesToLoad.Add("objectCategory")
        LDAPsearcher.PropertiesToLoad.Add("objectClass")
        LDAPsearcher.PropertiesToLoad.Add("userPrincipalName")

        For Each pattern As String In patterns
            Dim LDAPPattern As String = Replace(Replace(Replace(pattern, """", ""), "0", "*"), "#", "*")

            LDAPsearcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson))((userPrincipalName=" & LDAPPattern & "@*)))" 'user@domain
            LDAPsearcher.PageSize = 1000
            LDAPresults = LDAPsearcher.FindAll()

            Dim dummy As New List(Of String)
            For Each LDAPresult In LDAPresults
                dummy.Add(LCase(Split(GetLDAPProperty(LDAPresult.Properties, "userPrincipalName"), "@")(0)))
            Next LDAPresult

            For I As Integer = 1 To dummy.Count + 1
                Dim u As String = LCase(Format(I, pattern))
                If Not dummy.Contains(u) Then
                    result.Add(u)
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetNextDomainUser(pattern As String, domain As clsDomain) As String
        If domain Is Nothing Then Return Nothing
        Dim de As DirectoryEntry = domain.DefaultNamingContext
        If de Is Nothing Then Return Nothing

        Dim LDAPsearcher As New DirectorySearcher(de)
        Dim LDAPresults As SearchResultCollection = Nothing
        Dim LDAPresult As SearchResult

        LDAPsearcher.PropertiesToLoad.Add("objectCategory")
        LDAPsearcher.PropertiesToLoad.Add("objectClass")
        LDAPsearcher.PropertiesToLoad.Add("userPrincipalName")

        Dim LDAPPattern As String = Replace(Replace(Replace(pattern, """", ""), "0", "*"), "#", "*")

        LDAPsearcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson))((userPrincipalName=" & LDAPPattern & "@*)))" 'user@domain
        LDAPsearcher.PageSize = 1000
        LDAPresults = LDAPsearcher.FindAll()

        Dim dummy As New List(Of String)
        For Each LDAPresult In LDAPresults
            dummy.Add(LCase(Split(GetLDAPProperty(LDAPresult.Properties, "userPrincipalName"), "@")(0)))
        Next LDAPresult

        For I As Integer = 1 To dummy.Count + 1
            Dim u As String = LCase(Format(I, pattern))
            If Not dummy.Contains(u) Then
                Return u
                Exit For
            End If
        Next

        Return Nothing
    End Function

    Public Function GetNextDomainComputers(domain As clsDomain) As List(Of String)
        If domain Is Nothing Then Return Nothing
        Dim patterns() As String = domain.ComputerPattern.Split({",", vbCr, vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) Trim(x)).ToArray
        If patterns.Count = 0 Then Return Nothing
        Dim de As DirectoryEntry = domain.DefaultNamingContext
        If de Is Nothing Then Return Nothing

        Dim result As New List(Of String)

        Dim LDAPsearcher As New DirectorySearcher(de)
        Dim LDAPresults As SearchResultCollection = Nothing
        Dim LDAPresult As SearchResult

        LDAPsearcher.PropertiesToLoad.Add("objectCategory")
        LDAPsearcher.PropertiesToLoad.Add("name")

        For Each pattern As String In patterns
            Dim LDAPPattern As String = Replace(Replace(Replace(pattern, """", ""), "0", "*"), "#", "*")

            LDAPsearcher.Filter = "(&(objectCategory=computer)(name=" & LDAPPattern & "))"
            LDAPsearcher.PageSize = 1000
            LDAPresults = LDAPsearcher.FindAll()

            Dim dummy As New List(Of String)
            For Each LDAPresult In LDAPresults
                dummy.Add(LCase(GetLDAPProperty(LDAPresult.Properties, "name")))
            Next LDAPresult

            For I As Integer = 1 To dummy.Count + 1
                Dim u As String = LCase(Format(I, pattern))
                If Not dummy.Contains(u) Then
                    result.Add(u)
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetNextDomainTelephoneNumbers(domain As clsDomain) As ObservableCollection(Of clsTelephoneNumber)
        If domain Is Nothing Then Return Nothing
        Dim patterns As ObservableCollection(Of clsTelephoneNumberPattern) = domain.TelephoneNumberPattern
        If patterns.Count = 0 Then Return Nothing
        Dim de As DirectoryEntry = domain.DefaultNamingContext
        If de Is Nothing Then Return Nothing

        Dim result As New ObservableCollection(Of clsTelephoneNumber)

        Dim LDAPsearcher As New DirectorySearcher(de)
        Dim LDAPresults As SearchResultCollection = Nothing
        Dim LDAPresult As SearchResult

        LDAPsearcher.PropertiesToLoad.Add("objectCategory")
        LDAPsearcher.PropertiesToLoad.Add("objectClass")
        LDAPsearcher.PropertiesToLoad.Add("userAccountControl")
        LDAPsearcher.PropertiesToLoad.Add("telephoneNumber")

        For Each pattern As clsTelephoneNumberPattern In patterns
            If Not pattern.Range.Contains("-") Then Continue For
            Dim numstart As Long = 0
            Dim numend As Long = 0
            If Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(0), numstart) Or
               Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(1), numend) Then Continue For

            LDAPsearcher.Filter = "(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(telephoneNumber=*))"
            LDAPsearcher.PageSize = 1000
            LDAPresults = LDAPsearcher.FindAll()

            Dim dummy As New List(Of String)
            For Each LDAPresult In LDAPresults
                dummy.Add(GetLDAPProperty(LDAPresult.Properties, "telephoneNumber"))
            Next LDAPresult

            For I As Long = numstart To numend
                Dim u As String = LCase(Format(I, pattern.Pattern))
                If Not dummy.Contains(u) Then
                    result.Add(New clsTelephoneNumber(pattern.Label, u))
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function Translit_RU_EN(ByVal text As String) As String
        Dim Russian() As String = {"а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я"}
        Dim English() As String = {"a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "", "y", "", "e", "yu", "ya"}

        For I As Integer = 0 To Russian.Count - 1
            text = text.Replace(Russian(I), English(I))
            text = text.Replace(UCase(Russian(I)), UCase(English(I)))
        Next

        Return LCase(text)
    End Function

    Public Sub ShowWrongMemberMessage()
        IMsgBox("Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена." & vbCrLf &
               "Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы." & vbCrLf &
               "Локальная группа домена может быть членом только другой локальной группы домена." & vbCrLf & vbCrLf &
               "Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы." & vbCrLf &
               "Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы." & vbCrLf &
               "Универсальная группа не может быть членом глобальной группы.", "Неверный тип группы", vbOKOnly, vbExclamation)
    End Sub

    Public Function ShowDirectoryObjectProperties(obj As clsDirectoryObject, Optional owner As Window = Nothing) As Window
        If obj.SchemaClassName = "user" Then
            Dim w As wndUser
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndUser) Is wnd.GetType AndAlso CType(wnd, wndUser).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndUser
            If owner IsNot Nothing Then
                w.Owner = owner
                'w.Left = owner.Left + owner.ActualWidth / 2 - w.Width / 2
                'w.Top = owner.Top + owner.ActualHeight / 2 - w.Height / 2
            End If

            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClassName = "computer" Then
            Dim w As wndComputer
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndComputer) Is wnd.GetType AndAlso CType(wnd, wndComputer).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndComputer
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClassName = "group" Then
            Dim w As wndGroup
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndGroup) Is wnd.GetType AndAlso CType(wnd, wndGroup).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndGroup
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClassName = "contact" Then
            Dim w As wndContact
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndContact) Is wnd.GetType AndAlso CType(wnd, wndContact).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndContact
            If owner IsNot Nothing Then
                w.Owner = owner
                'w.Left = owner.Left + owner.ActualWidth / 2 - w.Width / 2
                'w.Top = owner.Top + owner.ActualHeight / 2 - w.Height / 2
            End If

            w.currentobject = obj
            w.Show()
            Return w
        Else
            Return Nothing
        End If
    End Function

    Public Sub Log(message As String)
        SingleInstanceApplication.tsocLog.Add(New clsLog(message))
    End Sub

    Public Function GetApplicationIcon(fileName As String) As ImageSource
        Dim ai As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(ai.Handle, New Int32Rect(0, 0, ai.Width, ai.Height), BitmapSizeOptions.FromEmptyOptions())
    End Function

    Public Function ToSecure(current As String) As SecureString
        Dim s = New SecureString()
        For Each c As Char In current.ToCharArray()
            s.AppendChar(c)
        Next
        Return s
    End Function

    Public Function GetLDAPPath(path As String) As String()
        Dim dNp As String() = path.Split({"LDAP:", "/"}, StringSplitOptions.RemoveEmptyEntries)
        If dNp.Count <= 1 Then Return {}

        Dim p As New List(Of String)
        p.Add(dNp(0))
        p.AddRange(dNp(1).Split({","}, StringSplitOptions.RemoveEmptyEntries).Reverse.Select(Function(x As String) x.Replace("OU=", "").Replace("CN=", "")).Where(Function(x As String) Not x.StartsWith("DC")).ToList)

        Return p.ToArray
    End Function

    Public Sub DoThePrint(document As System.Windows.Documents.FlowDocument)
        ' Clone the source document's content into a new FlowDocument.
        ' This is because the pagination for the printer needs to be
        ' done differently than the pagination for the displayed page.
        ' We print the copy, rather that the original FlowDocument.
        Dim s As New System.IO.MemoryStream()
        Dim source As New TextRange(document.ContentStart, document.ContentEnd)
        source.Save(s, System.Windows.DataFormats.Xaml)
        Dim copy As New FlowDocument()
        Dim dest As New TextRange(copy.ContentStart, copy.ContentEnd)
        dest.Load(s, System.Windows.DataFormats.Xaml)

        ' Create a XpsDocumentWriter object, implicitly opening a Windows common print dialog,
        ' and allowing the user to select a printer.

        ' get information about the dimensions of the seleted printer+media.
        Dim ia As System.Printing.PrintDocumentImageableArea = Nothing
        Dim docWriter As System.Windows.Xps.XpsDocumentWriter = System.Printing.PrintQueue.CreateXpsDocumentWriter(ia)

        If docWriter IsNot Nothing AndAlso ia IsNot Nothing Then
            Dim paginator As DocumentPaginator = DirectCast(copy, IDocumentPaginatorSource).DocumentPaginator

            ' Change the PageSize and PagePadding for the document to match the CanvasSize for the printer device.
            paginator.PageSize = New Size(ia.MediaSizeWidth, ia.MediaSizeHeight)
            Dim t As New Thickness(72)
            ' copy.PagePadding;
            copy.PagePadding = New Thickness(Math.Max(ia.OriginWidth, t.Left), Math.Max(ia.OriginHeight, t.Top), Math.Max(ia.MediaSizeWidth - (ia.OriginWidth + ia.ExtentWidth), t.Right), Math.Max(ia.MediaSizeHeight - (ia.OriginHeight + ia.ExtentHeight), t.Bottom))

            copy.ColumnWidth = Double.PositiveInfinity
            'copy.PageWidth = 528; // allow the page to be the natural with of the output device

            ' Send content to the printer.
            docWriter.Write(paginator)
        End If
    End Sub

    Public Function Ping(hostname As String) As PingReply
        Dim pingsender As New Ping
        Dim pingoptions As New PingOptions
        Dim pingtimeout As Integer = 1000
        Dim pingreplytask As Task(Of PingReply) = Nothing
        Dim pingreply As PingReply = Nothing
        Dim pingbuffer() As Byte = Encoding.ASCII.GetBytes(Space(32))
        Dim addresses() As IPAddress

        Try
            addresses = Dns.GetHostAddresses(hostname)
        Catch ex As Exception
            Return Nothing
        End Try

        If addresses.Count = 0 Then Return Nothing

        pingoptions.Ttl = 128
        pingoptions.DontFragment = False
        pingreplytask = Task.Run(Function() pingsender.Send(addresses(0), pingtimeout, pingbuffer, pingoptions))
        pingreplytask.Wait()

        Return pingreplytask.Result
    End Function

    Public Function TraceRoute(hostname As String) As List(Of PingReply)
        Dim pingsender As New Ping
        Dim pingoptions As New PingOptions
        Dim pingtimeout As Integer = 1000
        Dim pingreplytask As Task(Of PingReply) = Nothing
        Dim pingreply As PingReply = Nothing
        Dim pingbuffer() As Byte = Encoding.ASCII.GetBytes(Space(32))
        Dim addresses() As IPAddress

        Try
            addresses = Dns.GetHostAddresses(hostname)
        Catch ex As Exception
            Return New List(Of PingReply)
        End Try

        If addresses.Count = 0 Then Return New List(Of PingReply)

        Dim resultlist As New List(Of PingReply)

        For ttl As Integer = 1 To 128
            pingoptions.Ttl = ttl
            pingoptions.DontFragment = False

            pingreplytask = Task.Run(Function() pingsender.Send(addresses(0), pingtimeout, pingbuffer, pingoptions))
            pingreplytask.Wait()

            resultlist.Add(pingreplytask.Result)
            If pingreplytask.Result.Status = IPStatus.Success Then Exit For
        Next

        Return resultlist
    End Function

    Public Function PortScan(hostname As String, portlist As Dictionary(Of Integer, String)) As Dictionary(Of Integer, Boolean)
        Dim resultlist As New Dictionary(Of Integer, Boolean)

        For Each port As Integer In portlist.Keys
            Dim tcpClient = New TcpClient()
            Dim connectionTask = tcpClient.ConnectAsync(hostname, port).ContinueWith(Function(tsk) If(tsk.IsFaulted, Nothing, tcpClient))
            Dim timeoutTask = Task.Delay(1000).ContinueWith(Of TcpClient)(Function(tsk) Nothing)
            Dim resultTask = Task.WhenAny(connectionTask, timeoutTask).Unwrap()

            resultTask.Wait()
            Dim resultTcpClient = resultTask.Result
            If resultTcpClient IsNot Nothing Then
                resultlist.Add(port, resultTcpClient.Connected)
                resultTcpClient.Close()
            Else
                resultlist.Add(port, False)
            End If
        Next

        Return resultlist
    End Function

    Public Function GetLocalIPAddress() As String
        Dim host = Dns.GetHostEntry(Dns.GetHostName())
        For Each ip As IPAddress In host.AddressList
            If ip.AddressFamily = AddressFamily.InterNetwork Then
                Return ip.ToString()
            End If
        Next
        ThrowCustomException("Local IP Address Not Found!")
        Return Nothing
    End Function

End Module

#Region "XAML Value Converters"

Public Class BooleanAndConverter
    Implements IMultiValueConverter

    Private Function IMultiValueConverter_Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
        For Each value As Object In values
            If CBool(value) = False Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Function IMultiValueConverter_ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
        Throw New NotSupportedException("BooleanAndConverter is a OneWay converter.")
    End Function

End Class

Public Class InverseBooleanConverter
    Implements IValueConverter

    Public Function IValueConverter_Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        If targetType <> GetType(Boolean) Then
            Throw New InvalidOperationException("The target must be a boolean")
        End If

        Return Not CBool(value)
    End Function

    Public Function IValueConverter_ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        If targetType <> GetType(Boolean) Then
            Throw New InvalidOperationException("The target must be a boolean")
        End If

        Return Not CBool(value)
    End Function
End Class

#End Region
