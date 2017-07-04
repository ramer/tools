Imports System.ComponentModel

Public Class wndHost

    Public Property currentobject As clsHost

    Private Sub wndHost_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataContext = currentobject
        cmboType.ItemsSource = hosttypes
        cmboPingPriority.ItemsSource = pingpriorities
    End Sub

    Private Sub wndHost_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub tbUpdateInterval_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles tbUpdateInterval.PreviewTextInput
        If Not Char.IsDigit(e.Text, e.Text.Length - 1) Then e.Handled = True
    End Sub

    Private Sub btnUpdateIntervalDecrease_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdateIntervalDecrease.Click
        If currentobject.UpdateInterval <= 1 Then Exit Sub
        currentobject.UpdateInterval -= 1
    End Sub

    Private Sub btnUpdateIntervalIncrease_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdateIntervalIncrease.Click
        currentobject.UpdateInterval += 1
    End Sub

    Private Sub btnUpdateHostname_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdateHostname.Click
        currentobject.UpdateHostname()
    End Sub

    Private Sub btnUpdateIPAddress_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdateIPAddress.Click
        currentobject.UpdateIPAddress()
    End Sub

    Private Sub btnUpdateMACAddress_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdateMACAddress.Click
        currentobject.UpdateMACAddress()
    End Sub
End Class
