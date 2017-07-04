Imports System.ComponentModel

Public Class wndPreferences

    Private sourceobject As Object
    Private allowdrag As Boolean

    Private Sub wndPreferences_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataContext = preferences
    End Sub

    Private Sub tbDefaultUpdateInterval_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles tbDefaultUpdateInterval.PreviewTextInput
        If Not Char.IsDigit(e.Text, e.Text.Length - 1) Then e.Handled = True
    End Sub

    Private Sub cmboTheme_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboTheme.SelectionChanged
        Select Case cmboTheme.SelectedIndex
            Case 0 ' светло-серая
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightGray
                    .ColorSuccessBackground = Colors.LightGray
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFE8E6E6")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.WhiteSmoke
                End With
            Case 1 ' светло-синяя
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightSkyBlue
                    .ColorSuccessBackground = Colors.LightSkyBlue
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFD2EBFB")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.AliceBlue
                End With
            Case 2 ' светло-зеленая
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightGreen
                    .ColorSuccessBackground = Colors.LightGreen
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFC8EEC8")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.Honeydew
                End With
            Case 3 ' темно-серая
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.DarkGray
                    .ColorSuccessBackground = Colors.DarkGray
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF515151")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 4 ' темно-синяя
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.RoyalBlue
                    .ColorSuccessBackground = Colors.RoyalBlue
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF1D2E63")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 5 ' темно-зеленая
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.Green
                    .ColorSuccessBackground = Colors.Green
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF003000")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 6 ' темно-оранжевая
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.Chocolate
                    .ColorSuccessBackground = Colors.Chocolate
                    .ColorFailBackground = ColorConverter.ConvertFromString("#FFFA6464")
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF61310F")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
        End Select
    End Sub

    Private Sub chbStartWithWindows_Checked(sender As Object, e As RoutedEventArgs) Handles chbStartWithWindows.Checked, chbStartWithWindows.Unchecked, chbStartWithWindowsMinimized.Checked, chbStartWithWindowsMinimized.Unchecked
        If chbStartWithWindows.IsChecked Then
            If chbStartWithWindowsMinimized.IsChecked Then
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue("NetViewer", """" & System.Reflection.Assembly.GetExecutingAssembly().Location & """ -minimized") 'My.Application.Info.AssemblyName)
            Else
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue("NetViewer", System.Reflection.Assembly.GetExecutingAssembly().Location) 'My.Application.Info.AssemblyName)
            End If
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue("NetViewer")
        End If
    End Sub

    Private Sub wndPreferences_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub btnExternalSoftwareBrowse_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim es As clsExternalSoftware = CType(sender, FrameworkElement).DataContext
            Dim dlg As New Forms.OpenFileDialog
            dlg.Filter = "Приложение|*.exe"
            If dlg.ShowDialog = Forms.DialogResult.OK Then
                es.Path = dlg.FileName
            End If
        Catch
        End Try
    End Sub

    Private Sub btnDefaultUpdateIntervalDecrease_Click(sender As Object, e As RoutedEventArgs) Handles btnDefaultUpdateIntervalDecrease.Click
        If preferences.DefaultUpdateInterval <= 1 Then Exit Sub
        preferences.DefaultUpdateInterval -= 1
    End Sub

    Private Sub btnDefaultUpdateIntervalIncrease_Click(sender As Object, e As RoutedEventArgs) Handles btnDefaultUpdateIntervalIncrease.Click
        preferences.DefaultUpdateInterval += 1
    End Sub
End Class
