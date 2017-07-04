
Public Class wndPowerShellDebug

    Private Sub wndPowerShellDebug_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgLog.ItemsSource = SingleInstanceApplication.powershellLog
    End Sub

End Class
