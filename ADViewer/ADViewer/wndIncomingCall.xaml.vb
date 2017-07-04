Imports System.Collections.ObjectModel

Public Class wndIncomingCall

    Public WithEvents searcher As New clsSearcher

    Public Property objects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private Sub wndIncomingCall_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tblckTimeStamp.Text = Now.ToString
    End Sub

    Private attributesForSearchSIP As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("telephoneNumber", "", True)},
        {New clsAttribute("mobile", "", True)},
        {New clsAttribute("homePhone", "", True)},
        {New clsAttribute("ipPhone", "", True)}
    }

    Public Async Sub Search(pattern As String)
        cap.Visibility = Visibility.Visible

        Await searcher.BasicSearchAsync(objects, pattern,, attributesForSearchSIP, True, False, False)

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub Searcher_BasicSearchAsyncDataRecieved() Handles searcher.BasicSearchAsyncDataRecieved
        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub lvIncomingCall_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvIncomingCall.MouseDoubleClick
        If lvIncomingCall.SelectedItem Is Nothing Then Exit Sub

        ShowDirectoryObjectProperties(lvIncomingCall.SelectedItem, Me)
    End Sub

    Private Sub hlDisplayName_Click(sender As Object, e As RoutedEventArgs) Handles hlDisplayName.Click
        wndMainActivate(rnDisplayName.Text)
    End Sub

    Private Sub hlURI_Click(sender As Object, e As RoutedEventArgs) Handles hlURI.Click
        wndMainActivate(rnURI.Text)
    End Sub

    Private Sub wndMainActivate(pattern As String)
        Dim w As New wndMain
        w.Search(pattern)
        w.Show()
        w.Activate()
        w.Topmost = True
        w.Topmost = False
    End Sub

End Class
