
Public Class wndQuestion

    Public _inputbox As Boolean
    Public _passwordbox As Boolean
    Public _content As String
    Public _title As String
    Public _buttons As MsgBoxStyle
    Public _icon As MsgBoxStyle
    Public _defaultanswer As String
    Public _msgboxresult As MessageBoxResult
    Public _inputboxresult As String

    Private Sub wndQuestion_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        btnOK.Visibility = Visibility.Collapsed
        btnYes.Visibility = Visibility.Collapsed
        btnNo.Visibility = Visibility.Collapsed
        btnCancel.Visibility = Visibility.Collapsed

        tbContent.Text = If(_content, "")
        Title = If(_title, "")

        If _buttons = vbOK Or _buttons = vbOKOnly Or _buttons = vbOKCancel Then btnOK.Visibility = Visibility.Visible
        If _buttons = vbYesNo Or _buttons = vbYesNoCancel Then btnYes.Visibility = Visibility.Visible
        If _buttons = vbYesNo Or _buttons = vbYesNoCancel Then btnNo.Visibility = Visibility.Visible
        If _buttons = vbOKCancel Or _buttons = vbYesNoCancel Then btnCancel.Visibility = Visibility.Visible

        Select Case _icon
            Case vbInformation
                imgIcon.Source = New BitmapImage(New Uri("img/information.png", UriKind.Relative))
            Case vbQuestion
                imgIcon.Source = New BitmapImage(New Uri("img/question.png", UriKind.Relative))
            Case vbExclamation
                imgIcon.Source = New BitmapImage(New Uri("img/exclamation.png", UriKind.Relative))
            Case vbCritical
                imgIcon.Source = New BitmapImage(New Uri("img/critical.png", UriKind.Relative))
        End Select

        If _inputbox Or _passwordbox Then
            tbInput.Visibility = Visibility.Visible
        Else
            tbInput.Visibility = Visibility.Collapsed
        End If

        If _inputbox Then
            tbInput.Text = _defaultanswer
            tbInput.SelectAll()
            tbInput.Focus()
        End If

        If _passwordbox Then
            btnGenerate.Visibility = Visibility.Visible
            tbInput.Text = clsPassGenerator.Generate(20)
            tbInput.SelectAll()
            tbInput.Focus()
        Else
            btnGenerate.Visibility = Visibility.Collapsed
        End If

    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        _msgboxresult = MessageBoxResult.OK
        DialogResult = True
    End Sub

    Private Sub btnYes_Click(sender As Object, e As RoutedEventArgs) Handles btnYes.Click
        _msgboxresult = MessageBoxResult.Yes
        DialogResult = True
    End Sub

    Private Sub btnNo_Click(sender As Object, e As RoutedEventArgs) Handles btnNo.Click
        _msgboxresult = MessageBoxResult.No
        DialogResult = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        _msgboxresult = MessageBoxResult.Cancel
        DialogResult = False
    End Sub

    Private Sub wndQuestion_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key = Key.Escape Then
            If _buttons = vbOKCancel Or _buttons = vbYesNoCancel Then
                _msgboxresult = MessageBoxResult.Cancel
                DialogResult = False
            ElseIf _buttons = vbOKOnly Then
                _msgboxresult = MessageBoxResult.OK
                DialogResult = True
            End If
        End If
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As RoutedEventArgs) Handles btnGenerate.Click
        tbInput.Text = clsPassGenerator.Generate(20)
        tbInput.Focus()
    End Sub

End Class
