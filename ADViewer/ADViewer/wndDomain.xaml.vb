Imports System.ComponentModel
Imports System.Windows.Threading

Public Class wndDomain
    'Private _currentdomain As clsDomain

    'Public Property CurrentDomain() As clsDomain
    '    Get
    '        Return _currentdomain
    '    End Get
    '    Set(value As clsDomain)
    '        _currentdomain = value
    '    End Set
    'End Property

    Private Sub wndDomain_Loaded(sender As Object, e As RoutedEventArgs) Handles wndDomain.Loaded
        Dispatcher.BeginInvoke(New Action(AddressOf wndDomain_LoadedDelayed), DispatcherPriority.ContextIdle, Nothing)
    End Sub

    Private Sub wndDomain_LoadedDelayed()
        lvDomains.ItemsSource = domains
    End Sub

    Private Sub wndDomain_Closing(sender As Object, e As CancelEventArgs) Handles wndDomain.Closing
        For Each dn In domains
            If dn.Validated Then dn.SaveSettingsToRegistry()
        Next
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub btnDomainsAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsAdd.Click
        Dim newdomain = New clsDomain()
        domains.Add(newdomain)
        lvDomains.SelectedItem = newdomain
        tabctlDomain.SelectedIndex = 0
        tbPassword.Password = ""
        tbDomainName.Focus()
    End Sub

    Private Sub btnDomainsRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsRemove.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        CType(lvDomains.SelectedItem, clsDomain).DeleteSettingsFromRegistry()
        domains.Remove(lvDomains.SelectedItem)
    End Sub

    Private Sub tbPassword_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbPassword.LostFocus
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        CType(lvDomains.SelectedItem, clsDomain).Password = CType(sender, PasswordBox).Password
    End Sub

    Private Sub btnSearchRootBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnSearchRootBrowse.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        Dim w As New wndDomainBrowser With {.Owner = Me}
        w.currentdomain = CType(lvDomains.SelectedItem, clsDomain)
        w.ShowDialog()
        If w.currentcontainer IsNot Nothing Then
            CType(lvDomains.SelectedItem, clsDomain).SearchRoot = w.currentcontainer.Entry
        End If
    End Sub

    Private Async Sub btnValidate_Click(sender As Object, e As RoutedEventArgs) Handles btnValidate.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        cap.Visibility = Visibility.Visible
        Await CType(lvDomains.SelectedItem, clsDomain).Revalidate()
        cap.Visibility = Visibility.Hidden
    End Sub

End Class
