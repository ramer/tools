Public Class wndErrorDebug

    Private Sub wndErrorDebug_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgLog.ItemsSource = SingleInstanceApplication.errorLog
    End Sub

    Private Sub ctxmnuErrorCopy_Click() Handles ctxmnuErrorCopy.Click
        If dgLog.SelectedItem Is Nothing Then Exit Sub

        Dim ex As clsErrorLog = CType(dgLog.SelectedItem, clsErrorLog)
        Clipboard.SetText(ex.TimeStamp & vbCrLf & ex.Command & vbCrLf & ex.Obj & vbCrLf & ex.Err)
    End Sub

End Class
