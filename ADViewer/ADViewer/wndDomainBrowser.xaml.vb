Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.Threading.Tasks
Imports System.Windows.Threading

Public Class wndDomainBrowser

    Public currentdomain As clsDomain
    Public currentcontainer As clsDirectoryObject

    Private Async Sub wndDomainBrowser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If currentdomain Is Nothing Then Exit Sub
        cap.Visibility = Visibility.Visible

        tvDomainBrowser.ItemsSource = Await Task.Run(Function() {New clsDirectoryObject(currentdomain.DefaultNamingContext, currentdomain)})

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub tvDomainBrowser_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles tvDomainBrowser.SelectedItemChanged
        currentcontainer = CType(e.NewValue, clsDirectoryObject)
    End Sub

    Private Sub wndDomainBrowser_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub
End Class
