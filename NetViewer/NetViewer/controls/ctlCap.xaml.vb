Public Class ctlCap

    Private Sub grdBack_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles grdBack.MouseDown
        If e.ChangedButton = MouseButton.Left And e.ClickCount = 3 Then Me.Visibility = Visibility.Hidden
    End Sub

End Class
