Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Microsoft.Win32

Public Class clsPreferences
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    ' behavior
    Private _startwithwindows As Boolean
    Private _startwithwindowsminimized As Boolean
    Private _closeonxbutton As Boolean

    Private _defaultupdateinterval As Integer

    ' appearance
    Private _colortext As Color
    Private _colorwindowbackground As Color
    Private _colorelementbackground As Color
    Private _colormenubackground As Color
    Private _colorbuttonbackground As Color
    Private _colorsuccessbackground As Color
    Private _colorfailbackground As Color
    Private _colorbuttoninactivebackground As Color
    Private _colorlistviewrow As Color
    Private _colorlistviewalternationrow As Color

    ' externalsoftware
    Private _externalsoftware As New ObservableCollection(Of clsExternalSoftware)

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

            StartWithWindows = regSettings.GetValue("StartWithWindows", False)
            StartWithWindowsMinimized = regSettings.GetValue("StartWithWindowsMinimized", False)
            CloseOnXButton = regSettings.GetValue("CloseOnXButton", True)

            DefaultUpdateInterval = regSettings.GetValue("DefaultUpdateInterval", 30)

            ColorText = ColorConverter.ConvertFromString(regSettings.GetValue("ColorText", Colors.Black.ToString))
            ColorWindowBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorWindowBackground", Colors.WhiteSmoke.ToString))
            ColorElementBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorElementBackground", Colors.White.ToString))
            ColorMenuBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorMenuBackground", Colors.WhiteSmoke.ToString))
            ColorButtonBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorButtonBackground", Colors.LightSkyBlue.ToString))
            ColorSuccessBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorSuccessBackground", Colors.LightSkyBlue.ToString))
            ColorFailBackground = ColorConverter.ConvertFromString(regSettings.GetValue("ColorFailBackground", "#FFFA6464"))
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

                esl.Add(es)
            Next
            ExternalSoftware = esl

        Catch ex As Exception
            ThrowException(ex, "LoadSettingsFromRegistry")
        End Try
    End Sub

    Public Sub SavePreferences()
        Try

            regSettings.SetValue("StartWithWindows", StartWithWindows)
            regSettings.SetValue("StartWithWindowsMinimized", StartWithWindowsMinimized)
            regSettings.SetValue("CloseOnXButton", CloseOnXButton)

            regSettings.SetValue("DefaultUpdateInterval", DefaultUpdateInterval)

            regSettings.SetValue("ColorText", ColorText)
            regSettings.SetValue("ColorWindowBackground", ColorWindowBackground)
            regSettings.SetValue("ColorElementBackground", ColorElementBackground)
            regSettings.SetValue("ColorMenuBackground", ColorMenuBackground)
            regSettings.SetValue("ColorButtonBackground", ColorButtonBackground)
            regSettings.SetValue("ColorSuccessBackground", ColorSuccessBackground)
            regSettings.SetValue("ColorFailBackground", ColorFailBackground)
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
            Next

        Catch ex As Exception
            ThrowException(ex, "SaveSettingsToRegistry")
        End Try
    End Sub

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

    Public Property DefaultUpdateInterval As Integer
        Get
            Return _defaultupdateinterval
        End Get
        Set(value As Integer)
            If value < 1 Then value = 1
            _defaultupdateinterval = value
            NotifyPropertyChanged("DefaultUpdateInterval")
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

    Public Property ColorFailBackground As Color
        Get
            Return _colorfailbackground
        End Get
        Set(value As Color)
            _colorfailbackground = value
            Application.Current.Resources("ColorFailBackground") = New SolidColorBrush(_colorfailbackground)
            NotifyPropertyChanged("ColorFailBackground")
        End Set
    End Property

    Public Property ColorSuccessBackground As Color
        Get
            Return _colorsuccessbackground
        End Get
        Set(value As Color)
            _colorsuccessbackground = value
            Application.Current.Resources("ColorSuccessBackground") = New SolidColorBrush(_colorsuccessbackground)
            NotifyPropertyChanged("ColorSuccessBackground")
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

End Class
