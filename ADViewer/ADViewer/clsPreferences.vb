Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports CredentialManagement
Imports LumiSoft.Net
Imports Microsoft.Win32

Public Class clsPreferences
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    ' basic
    Private _clipboardsource As Boolean
    Private _clipboardsourcelimit As Boolean
    Private _searchresultgrouping As Boolean
    Private _searchresultincludeusers As Boolean
    Private _searchresultincludecomputers As Boolean
    Private _searchresultincludegroups As Boolean
    Private _powershelldebug As Boolean

    ' layout
    Private _columns As New ObservableCollection(Of clsDataGridColumnInfo)

    ' search attributes
    Private _attributesforsearch As New ObservableCollection(Of clsAttribute)

    ' behavior
    Private _startwithwindows As Boolean
    Private _startwithwindowsminimized As Boolean
    Private _closeonxbutton As Boolean

    ' appearance
    Private _colortext As Color
    Private _colorwindowbackground As Color
    Private _colorelementbackground As Color
    Private _colormenubackground As Color
    Private _colorbuttonbackground As Color
    Private _colorbuttoninactivebackground As Color
    Private _colorlistviewrow As Color
    Private _colorlistviewalternationrow As Color

    ' externalsoftware
    Private _externalsoftware As New ObservableCollection(Of clsExternalSoftware)

    ' SIP
    Private _sipuse As Boolean
    Private _sipserver As String
    Private _sipregistrationname As String
    Private _sipusername As String
    Private _sippassword As String
    Private _sipdomain As String
    Private _sipprotocol As BindInfoProtocol = BindInfoProtocol.TCP

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()
        SetupPreferences()
    End Sub


    Public Sub SetupPreferences()
        Try
            ClipboardSource = regSettings.GetValue("ClipboardSource", False)
            ClipboardSourceLimit = regSettings.GetValue("ClipboardSourceLimit", True)
            SearchResultGrouping = regSettings.GetValue("SearchResultGrouping", False)
            SearchResultIncludeUsers = regSettings.GetValue("SearchResultIncludeUsers", True)
            SearchResultIncludeComputers = regSettings.GetValue("SearchResultIncludeComputers", True)
            SearchResultIncludeGroups = regSettings.GetValue("SearchResultIncludeGroups", True)
            PowershellDebug = regSettings.GetValue("PowershellDebug", False)

            Dim regInterfaceColumns As RegistryKey = regInterface.CreateSubKey("Columns")
            If regInterfaceColumns.GetSubKeyNames.Count > 0 Then

                Dim cil As New ObservableCollection(Of clsDataGridColumnInfo)
                For Each DisplayIndex As Integer In regInterfaceColumns.GetSubKeyNames()
                    Dim ci As New clsDataGridColumnInfo
                    Dim regInterfaceColumn As RegistryKey = regInterfaceColumns.OpenSubKey(DisplayIndex)
                    ci.DisplayIndex = DisplayIndex
                    ci.Width = regInterfaceColumn.GetValue("Width", 150.0)
                    ci.Header = regInterfaceColumn.GetValue("Header", "")
                    Dim cial As New List(Of clsAttribute)
                    For Each num As Integer In regInterfaceColumn.GetSubKeyNames()
                        Dim cia As New clsAttribute
                        Dim regInterfaceColumnAttribute As RegistryKey = regInterfaceColumn.OpenSubKey(num)
                        cia.Name = regInterfaceColumnAttribute.GetValue("Name", "")
                        cia.Label = regInterfaceColumnAttribute.GetValue("Label", "")
                        cia.IsDefault = regInterfaceColumnAttribute.GetValue("IsDefault", False)
                        cial.Add(cia)
                    Next
                    ci.Attributes = cial
                    cil.Add(ci)
                Next
                Columns = cil
            Else
                DefaultColumns()
            End If

            Dim regSettingsAtrributesForSearch As RegistryKey = regSettings.CreateSubKey("AtrributesForSearch")

            Dim afsl = New ObservableCollection(Of clsAttribute)
            For Each name As String In regSettingsAtrributesForSearch.GetSubKeyNames()
                Dim afs As New clsAttribute

                Dim reg As RegistryKey = regSettingsAtrributesForSearch.OpenSubKey(name)
                afs.Name = name
                afs.Label = reg.GetValue("Label", "")
                afs.IsDefault = reg.GetValue("IsDefault", True)

                afsl.Add(afs)
            Next
            AttributesForSearch = If(afsl.Count > 0, afsl, attributesForSearchDefault)

            StartWithWindows = regSettings.GetValue("StartWithWindows", False)
            StartWithWindowsMinimized = regSettings.GetValue("StartWithWindowsMinimized", False)
            CloseOnXButton = regSettings.GetValue("CloseOnXButton", True)

            ColorText = ColorConverter.ConvertFromString(regSettings.GetValue("ColorText", Colors.Black.ToString))
            ColorWindowBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorWindowBackground", Colors.WhiteSmoke.ToString))
            ColorElementBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorElementBackground", Colors.White.ToString))
            ColorMenuBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorMenuBackground", Colors.WhiteSmoke.ToString))
            ColorButtonBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorButtonBackground", Colors.LightSkyBlue.ToString))
            ColorButtonInactiveBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorButtonInactiveBackground", "#FFD2EBFB"))
            ColorListviewRow = ColorConverter.ConvertFromString(regSettings.GetValue("ColorListviewRow", Colors.White.ToString))
            ColorListviewAlternationRow = ColorConverter.ConvertFromString(regSettings.GetValue("ColorListviewAlternationRow", Colors.AliceBlue.ToString))

            Dim regExternalSoftwareSoftware As RegistryKey = regExternalSoftware.CreateSubKey("Software")

            Dim esl = New ObservableCollection(Of clsExternalSoftware)
            For Each lbl As String In regExternalSoftwareSoftware.GetSubKeyNames()
                Dim es As New clsExternalSoftware

                Dim reg As RegistryKey = regExternalSoftwareSoftware.OpenSubKey(lbl)
                es.Label = lbl
                es.Path = reg.GetValue("Path", "")
                es.Arguments = reg.GetValue("Arguments", "")
                es.CurrentCredentials = reg.GetValue("CurrentCredentials", True)

                esl.Add(es)
            Next
            ExternalSoftware = esl

            SipUse = regSettings.GetValue("SipUse", False)
            SipServer = regSettings.GetValue("SipServer", "")
            SipRegistrationName = regSettings.GetValue("SipRegistrationName", "")
            SipDomain = regSettings.GetValue("SipDomain", "")
            SipProtocol = If(regSettings.GetValue("SipProtocol", "TCP") = "TCP", BindInfoProtocol.TCP, BindInfoProtocol.UDP)
            Dim cred As New Credential("", "", "ADViewerSIP", CredentialType.Generic)
            cred.PersistanceType = PersistanceType.Enterprise
            cred.Load()
            SipUsername = cred.Username
            SipPassword = cred.Password

        Catch ex As Exception
            ThrowException(ex, "LoadSettingsFromRegistry")
        End Try
    End Sub

    Public Sub SavePreferences()
        Try
            regSettings.SetValue("ClipboardSource", ClipboardSource)
            regSettings.SetValue("ClipboardSourceLimit", ClipboardSourceLimit)
            regSettings.SetValue("SearchResultGrouping", SearchResultGrouping)
            regSettings.SetValue("SearchResultIncludeUsers", SearchResultIncludeUsers)
            regSettings.SetValue("SearchResultIncludeComputers", SearchResultIncludeComputers)
            regSettings.SetValue("SearchResultIncludeGroups", SearchResultIncludeGroups)
            regSettings.SetValue("PowershellDebug", PowershellDebug)

            regInterface.DeleteSubKeyTree("Columns", False)
            Dim regInterfaceColumns As RegistryKey = regInterface.CreateSubKey("Columns")

            For Each column As clsDataGridColumnInfo In Columns
                Dim regInterfaceColumn As RegistryKey = regInterfaceColumns.CreateSubKey(column.DisplayIndex)
                With regInterfaceColumn
                    .SetValue("Width", Int(column.Width))
                    .SetValue("Header", column.Header)
                End With
                For I As Integer = 0 To column.Attributes.Count - 1
                    Dim regInterfaceColumnAttribute As RegistryKey = regInterfaceColumn.CreateSubKey(I)
                    With regInterfaceColumnAttribute
                        .SetValue("Label", column.Attributes(I).Label)
                        .SetValue("Name", column.Attributes(I).Name)
                        .SetValue("IsDefault", column.Attributes(I).IsDefault)
                    End With
                Next
            Next

            regSettings.DeleteSubKeyTree("AtrributesForSearch", False)
            Dim regSettingsAtrributesForSearch As RegistryKey = regSettings.CreateSubKey("AtrributesForSearch")

            For Each afs As clsAttribute In AttributesForSearch
                Dim regSettingsAtrributesForSearchAttribute As RegistryKey = regSettingsAtrributesForSearch.CreateSubKey(afs.Name)
                With regSettingsAtrributesForSearchAttribute
                    .SetValue("Label", afs.Label)
                    .SetValue("IsDefault", afs.IsDefault)
                End With
            Next

            regSettings.SetValue("StartWithWindows", StartWithWindows)
            regSettings.SetValue("StartWithWindowsMinimized", StartWithWindowsMinimized)
            regSettings.SetValue("CloseOnXButton", CloseOnXButton)

            regSettings.SetValue("ColorText", ColorText)
            regSettings.SetValue("ColorWindowBackground", ColorWindowBackground)
            regSettings.SetValue("ColorElementBackground", ColorElementBackground)
            regSettings.SetValue("ColorMenuBackground", ColorMenuBackground)
            regSettings.SetValue("ColorButtonBackground", ColorButtonBackground)
            regSettings.SetValue("ColorButtonInactiveBackground", ColorButtonInactiveBackground)
            regSettings.SetValue("ColorListviewRow", ColorListviewRow)
            regSettings.SetValue("ColorListviewAlternationRow", ColorListviewAlternationRow)

            regExternalSoftware.DeleteSubKeyTree("Software", False)
            Dim ADViewerExternalSoftwareSettingsSoftwareRegPath As RegistryKey = regExternalSoftware.CreateSubKey("Software")

            For Each es As clsExternalSoftware In ExternalSoftware
                If es.Label Is Nothing Then Continue For
                Dim reg As RegistryKey = ADViewerExternalSoftwareSettingsSoftwareRegPath.CreateSubKey(es.Label)
                reg.SetValue("Path", If(es.Path, ""), RegistryValueKind.String)
                reg.SetValue("Arguments", If(es.Arguments, ""), RegistryValueKind.String)
                reg.SetValue("CurrentCredentials", es.CurrentCredentials, RegistryValueKind.String)
            Next

            regSettings.SetValue("SipUse", SipUse)
            regSettings.SetValue("SipServer", SipServer)
            regSettings.SetValue("SipRegistrationName", SipRegistrationName)
            regSettings.SetValue("SipDomain", SipDomain)
            regSettings.SetValue("SipProtocol", SipProtocol.ToString)
            If Not String.IsNullOrEmpty(SipUsername) And Not String.IsNullOrEmpty(SipPassword) Then
                Dim cred As New Credential("", "", "ADViewerSIP", CredentialType.Generic)
                cred.PersistanceType = PersistanceType.Enterprise
                cred.Username = SipUsername
                cred.Password = SipPassword
                cred.Save()
            End If

        Catch ex As Exception
            ThrowException(ex, "SaveSettingsToRegistry")
        End Try
    End Sub

    Public Sub DefaultColumns()
        Dim layout As New ObservableCollection(Of clsDataGridColumnInfo)
        layout.Add(New clsDataGridColumnInfo("⬕", New List(Of clsAttribute) From {New clsAttribute("Image", "⬕")}, 0, 50))
        layout.Add(New clsDataGridColumnInfo("Имя", New List(Of clsAttribute) From {New clsAttribute("name", "Имя объекта"), New clsAttribute("description", "Описание")}, 1, 250))
        layout.Add(New clsDataGridColumnInfo("Имя входа", New List(Of clsAttribute) From {New clsAttribute("userPrincipalName", "Имя входа"), New clsAttribute("distinguishedName", "LDAP-путь")}, 2, 500))
        layout.Add(New clsDataGridColumnInfo("Телефон", New List(Of clsAttribute) From {New clsAttribute("telephoneNumber", "Телефон"), New clsAttribute("physicalDeliveryOfficeName", "Офис")}, 3, 150))
        layout.Add(New clsDataGridColumnInfo("Место работы", New List(Of clsAttribute) From {New clsAttribute("title", "Должность"), New clsAttribute("department", "Подразделение"), New clsAttribute("company", "Компания")}, 4, 400))
        layout.Add(New clsDataGridColumnInfo("Основной адрес", New List(Of clsAttribute) From {New clsAttribute("mail", "Основной адрес")}, 5, 200))
        layout.Add(New clsDataGridColumnInfo("Объект", New List(Of clsAttribute) From {New clsAttribute("whenCreatedFormated", "Создан"), New clsAttribute("lastLogonFormated", "Последний вход"), New clsAttribute("accountExpiresFormated", "Объект истекает")}, 6, 150))
        layout.Add(New clsDataGridColumnInfo("Пароль", New List(Of clsAttribute) From {New clsAttribute("pwdLastSetFormated", "Пароль изменен"), New clsAttribute("passwordExpiresFormated", "Пароль истекает")}, 7, 150))
        Columns = layout
    End Sub

    Public Sub LoadColumnsFromRegistry(ByRef dg As DataGrid)
        Dim ADViewerColumnSettingsRegPath As RegistryKey = regInterface.CreateSubKey("Columns")

        For Each Column As DataGridColumn In dg.Columns
            If Column.Header Is Nothing Then Continue For
            Dim ADViewerCurrentColumnRegPath As RegistryKey = ADViewerColumnSettingsRegPath.OpenSubKey(Column.Header.ToString)
            If ADViewerCurrentColumnRegPath IsNot Nothing Then
                With ADViewerCurrentColumnRegPath
                    Column.Width = Double.Parse(.GetValue("Width", Column.Width))
                    Column.DisplayIndex = Integer.Parse(.GetValue("DisplayIndex", Column.DisplayIndex))
                End With
            End If
        Next
    End Sub

    Public Sub SaveColumnsToRegistry(ByRef dg As DataGrid)
        regInterface.DeleteSubKeyTree("Columns", False)
        Dim ADViewerColumnSettingsRegPath As RegistryKey = regInterface.CreateSubKey("Columns")

        For Each Column As DataGridColumn In dg.Columns
            If Column.Header Is Nothing Then Continue For
            Dim ADViewerCurrentColumnRegPath As RegistryKey = ADViewerColumnSettingsRegPath.CreateSubKey(Column.Header.ToString)
            With ADViewerCurrentColumnRegPath
                .SetValue("Width", Int(Column.ActualWidth))
                .SetValue("DisplayIndex", Column.DisplayIndex)
            End With
        Next
    End Sub

    Public Property ClipboardSource As Boolean
        Get
            Return _clipboardsource
        End Get
        Set(value As Boolean)
            _clipboardsource = value
            NotifyPropertyChanged("ClipboardSource")
        End Set
    End Property

    Public Property ClipboardSourceLimit As Boolean
        Get
            Return _clipboardsourcelimit
        End Get
        Set(value As Boolean)
            _clipboardsourcelimit = value
            NotifyPropertyChanged("ClipboardSourceLimit")
        End Set
    End Property

    Public Property SearchResultGrouping As Boolean
        Get
            Return _searchresultgrouping
        End Get
        Set(value As Boolean)
            _searchresultgrouping = value
            NotifyPropertyChanged("SearchResultGrouping")
        End Set
    End Property

    Public Property SearchResultIncludeUsers As Boolean
        Get
            Return _searchresultincludeusers
        End Get
        Set(value As Boolean)
            _searchresultincludeusers = value
            NotifyPropertyChanged("SearchResultIncludeUsers")
        End Set
    End Property

    Public Property SearchResultIncludeComputers As Boolean
        Get
            Return _searchresultincludecomputers
        End Get
        Set(value As Boolean)
            _searchresultincludecomputers = value
            NotifyPropertyChanged("SearchResultIncludeComputers")
        End Set
    End Property

    Public Property SearchResultIncludeGroups As Boolean
        Get
            Return _searchresultincludegroups
        End Get
        Set(value As Boolean)
            _searchresultincludegroups = value
            NotifyPropertyChanged("SearchResultIncludeGroups")
        End Set
    End Property

    Public Property PowershellDebug As Boolean
        Get
            Return _powershelldebug
        End Get
        Set(value As Boolean)
            _powershelldebug = value
            NotifyPropertyChanged("PowershellDebug")
        End Set
    End Property

    Public Property Columns() As ObservableCollection(Of clsDataGridColumnInfo)
        Get
            Return _columns
        End Get
        Set(value As ObservableCollection(Of clsDataGridColumnInfo))
            _columns = value

            For Each w As Window In Application.Current.Windows
                If w.GetType Is GetType(wndMain) Then
                    CType(w, wndMain).RebuildColumns()
                End If
            Next
        End Set
    End Property

    Public Property AttributesForSearch() As ObservableCollection(Of clsAttribute)
        Get
            Return _attributesforsearch
        End Get
        Set(value As ObservableCollection(Of clsAttribute))
            _attributesforsearch = value
            NotifyPropertyChanged("AttributesForSearch")
        End Set
    End Property

    Public Property StartWithWindows As Boolean
        Get
            Return _startwithwindows
        End Get
        Set(value As Boolean)
            _startwithwindows = value
            NotifyPropertyChanged("StartWithWindows")
        End Set
    End Property

    Public Property StartWithWindowsMinimized As Boolean
        Get
            Return _startwithwindowsminimized
        End Get
        Set(value As Boolean)
            _startwithwindowsminimized = value
            NotifyPropertyChanged("StartWithWindowsMinimized")
        End Set
    End Property

    Public Property CloseOnXButton As Boolean
        Get
            Return _closeonxbutton
        End Get
        Set(value As Boolean)
            _closeonxbutton = value
            NotifyPropertyChanged("CloseOnXButton")
        End Set
    End Property

    Public Property ColorText As Color
        Get
            Return _colortext
        End Get
        Set(value As Color)
            _colortext = value
            Application.Current.Resources("ColorText") = New SolidColorBrush(_colortext)
            NotifyPropertyChanged("ColorText")
        End Set
    End Property

    Public Property ColorWindowBackground As Color
        Get
            Return _colorwindowbackground
        End Get
        Set(value As Color)
            _colorwindowbackground = value
            Application.Current.Resources("ColorWindowBackground") = New SolidColorBrush(_colorwindowbackground)
            NotifyPropertyChanged("ColorWindowBackground")
        End Set
    End Property

    Public Property ColorElementBackground As Color
        Get
            Return _colorelementbackground
        End Get
        Set(value As Color)
            _colorelementbackground = value
            Application.Current.Resources("ColorElementBackground") = New SolidColorBrush(_colorelementbackground)
            NotifyPropertyChanged("ColorElementBackground")
        End Set
    End Property

    Public Property ColorMenuBackground As Color
        Get
            Return _colormenubackground
        End Get
        Set(value As Color)
            _colormenubackground = value
            Application.Current.Resources("ColorMenuBackground") = New SolidColorBrush(_colormenubackground)
            NotifyPropertyChanged("ColorMenuBackground")
        End Set
    End Property

    Public Property ColorButtonBackground As Color
        Get
            Return _colorbuttonbackground
        End Get
        Set(value As Color)
            _colorbuttonbackground = value
            Application.Current.Resources("ColorButtonBackground") = New SolidColorBrush(_colorbuttonbackground)
            NotifyPropertyChanged("ColorButtonBackground")
        End Set
    End Property

    Public Property ColorButtonInactiveBackground As Color
        Get
            Return _colorbuttoninactivebackground
        End Get
        Set(value As Color)
            _colorbuttoninactivebackground = value
            Application.Current.Resources("ColorButtonInactiveBackground") = New SolidColorBrush(_colorbuttoninactivebackground)
            NotifyPropertyChanged("ColorButtonInactiveBackground")
        End Set
    End Property

    Public Property ColorListviewRow As Color
        Get
            Return _colorlistviewrow
        End Get
        Set(value As Color)
            _colorlistviewrow = value
            Application.Current.Resources("ColorListviewRow") = New SolidColorBrush(_colorlistviewrow)
            NotifyPropertyChanged("ColorListviewRow")
        End Set
    End Property

    Public Property ColorListviewAlternationRow As Color
        Get
            Return _colorlistviewalternationrow
        End Get
        Set(value As Color)
            _colorlistviewalternationrow = value
            Application.Current.Resources("ColorListviewAlternationRow") = New SolidColorBrush(_colorlistviewalternationrow)
            NotifyPropertyChanged("ColorListviewAlternationRow")
        End Set
    End Property

    Public Property ExternalSoftware As ObservableCollection(Of clsExternalSoftware)
        Get
            Return _externalsoftware
        End Get
        Set(value As ObservableCollection(Of clsExternalSoftware))
            _externalsoftware = value
            NotifyPropertyChanged("ExternalSoftware")
        End Set
    End Property

    Public Property SipUse As Boolean
        Get
            Return _sipuse
        End Get
        Set(value As Boolean)
            _sipuse = value
            NotifyPropertyChanged("SipUse")
        End Set
    End Property

    Public Property SipServer As String
        Get
            Return _sipserver
        End Get
        Set(value As String)
            _sipserver = value
            NotifyPropertyChanged("SipServer")
        End Set
    End Property

    Public Property SipRegistrationName As String
        Get
            Return _sipregistrationname
        End Get
        Set(value As String)
            _sipregistrationname = value
            NotifyPropertyChanged("SipRegistrationName")
        End Set
    End Property

    Public Property SipUsername As String
        Get
            Return _sipusername
        End Get
        Set(value As String)
            _sipusername = value
            NotifyPropertyChanged("SipUsername")
        End Set
    End Property

    Public Property SipPassword As String
        Get
            Return _sippassword
        End Get
        Set(value As String)
            _sippassword = value
            NotifyPropertyChanged("SipPassword")
        End Set
    End Property

    Public Property SipDomain As String
        Get
            Return _sipdomain
        End Get
        Set(value As String)
            _sipdomain = value
            NotifyPropertyChanged("SipDomain")
        End Set
    End Property

    Public Property SipProtocol As BindInfoProtocol
        Get
            Return _sipprotocol
        End Get
        Set(value As BindInfoProtocol)
            _sipprotocol = value
            NotifyPropertyChanged("SipProtocol")
        End Set
    End Property
End Class
