Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.DirectoryServices
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions

Class wndMain
    Public WithEvents searcher As New clsSearcher
    Public WithEvents clipboardTimer As New Threading.DispatcherTimer()

    Public Shared hkF5 As New RoutedCommand

    Public Property objects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private clipboardtext As String

#Region "Events"

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        RebuildColumns()
        clipboardTimer.Interval = New TimeSpan(0, 0, 1)
        clipboardTimer.Start()

        If domains.Count = 0 Then
            Dim w As New wndDomain With {.Owner = Me}
            w.ShowDialog()
        End If

        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf HotKey_F5))

        imgSipStatus.DataContext = sip
        grdSipStatus.DataContext = preferences
        tbSearchPattern.Focus()
    End Sub

    Private Sub wndMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim count As Integer = 0

        For Each wnd As Window In SingleInstanceApplication.Current.Windows
            If GetType(wndMain) Is wnd.GetType Then count += 1
        Next

        If preferences.CloseOnXButton AndAlso count <= 1 Then Application.Current.Shutdown()
    End Sub

    Private Sub clipboardTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles clipboardTimer.Tick
        If preferences IsNot Nothing AndAlso preferences.ClipboardSource Then

            Dim newclipboarddata As String = Clipboard.GetText

            If preferences.ClipboardSourceLimit Then
                If CountWords(newclipboarddata) <= 2 Then 'не больше трех слов
                    If clipboardtext <> newclipboarddata Then
                        clipboardtext = newclipboarddata
                        Search(clipboardtext)
                    End If
                End If
            Else
                If clipboardtext <> newclipboarddata Then
                    clipboardtext = newclipboarddata
                    Search(clipboardtext)
                End If
            End If
        End If
    End Sub

    Private Sub tbSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles tbSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            Search(CType(sender, TextBox).Text)
        End If
    End Sub

    Private Sub dgMain_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgMain.MouseDoubleClick
        If dgMain.SelectedItem Is Nothing Then Exit Sub

        ShowDirectoryObjectProperties(dgMain.SelectedItem, Me)
    End Sub

    Private Sub dgMain_ColumnReordered(sender As Object, e As DataGridColumnEventArgs) Handles dgMain.ColumnReordered
        UpdateColumns()
    End Sub

    Private Sub dgMain_LayoutUpdated(sender As Object, e As EventArgs) Handles dgMain.LayoutUpdated
        UpdateColumns()
    End Sub

    Private Sub dgMain_LoadingRow(sender As Object, e As DataGridRowEventArgs) Handles dgMain.LoadingRow
        e.Row.Header = (e.Row.GetIndex + 1).ToString
    End Sub

    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        Dim w As New wndMain
        w.Show()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click
        Dim str As String = ""
        For I = 0 To 40
            str &= clsPassGenerator.Generate(20) & vbCrLf
        Next
        MsgBox(str)
        Clipboard.SetText(str)
    End Sub


#End Region

#Region "Subs"

    Public Sub RebuildColumns()
        If objects.Count > 10 Then objects.Clear()

        dgMain.Columns.Clear()
        For Each columninfo As clsDataGridColumnInfo In preferences.Columns
            dgMain.Columns.Add(CreateColumn(columninfo))
        Next
    End Sub

    Public Sub UpdateColumns()
        For Each dgcolumn As DataGridColumn In dgMain.Columns
            For Each pcolumn As clsDataGridColumnInfo In preferences.Columns
                If dgcolumn.Header.ToString = pcolumn.Header Then
                    pcolumn.DisplayIndex = dgcolumn.DisplayIndex
                    pcolumn.Width = dgcolumn.ActualWidth
                End If
            Next
        Next
    End Sub

    Private Sub HotKey_F5()
        SearchRepeat()
    End Sub

    Private Sub SearchRepeat()
        Search(tbSearchPattern.Text)
    End Sub

    Public Async Sub Search(pattern As String)
        If String.IsNullOrEmpty(pattern) Then Exit Sub

        Try
            CollectionViewSource.GetDefaultView(dgMain.ItemsSource).GroupDescriptions.Clear()
        Catch
        End Try

        tbSearchPattern.Text = pattern
        tbSearchPattern.SelectAll()

        cap.Visibility = Visibility.Visible
        pbSearch.Visibility = Visibility.Visible

        Await searcher.BasicSearchAsync(objects, pattern,, preferences.AttributesForSearch, preferences.SearchResultIncludeUsers, preferences.SearchResultIncludeComputers, preferences.SearchResultIncludeGroups)

        If preferences.SearchResultGrouping Then
            Try
                CollectionViewSource.GetDefaultView(dgMain.ItemsSource).GroupDescriptions.Add(New PropertyGroupDescription("Domain.Name"))
            Catch
            End Try
        End If

        cap.Visibility = Visibility.Hidden
        pbSearch.Visibility = Visibility.Hidden
    End Sub

    Private Sub Searcher_BasicSearchAsyncDataRecieved() Handles searcher.BasicSearchAsyncDataRecieved
        cap.Visibility = Visibility.Hidden
    End Sub

#End Region

#Region "Menu"

    Private Sub mnuFileExit_Click(sender As Object, e As RoutedEventArgs) Handles mnuFileExit.Click
        Application.Current.Shutdown()
    End Sub

    Private Sub mnuEditCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles mnuEditCreateObject.Click
        Dim w As New wndCreateObject With {.Owner = Me}
        w.Show()
    End Sub

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click
        Dim w As New wndDomain With {.Owner = Me}
        w.ShowDialog()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        Dim w As New wndPreferences With {.Owner = Me}
        w.ShowDialog()
    End Sub

    Private Sub mnuServiceDeletedObjects_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDeletedObjects.Click
        Dim w As wndDeletedObjects
        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndDeletedObjects) Is wnd.GetType Then
                w = wnd
                w.Show() : w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                w.Topmost = True : w.Topmost = False
                Exit Sub
            End If
        Next

        w = New wndDeletedObjects
        w.Owner = Me
        w.Show()
    End Sub

    Private Sub mnuServiceADCommander_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceADCommander.Click
        Dim w As wndCommander
        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndCommander) Is wnd.GetType Then
                w = wnd
                w.Show() : w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                w.Topmost = True : w.Topmost = False
                Exit Sub
            End If
        Next

        w = New wndCommander
        w.Owner = Me
        w.Show()
    End Sub

    Private Sub mnuServiceLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceLog.Click
        Dim w As wndLog

        For Each wnd As Window In Application.Current.Windows
            If GetType(wndLog) Is wnd.GetType Then
                w = wnd
                w.Show()
                w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                Exit Sub
            End If
        Next

        w = New wndLog
        w.Show()
    End Sub

    Private Sub mnuServiceError_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceError.Click
        Dim w As wndError

        For Each wnd As Window In Application.Current.Windows
            If GetType(wndError) Is wnd.GetType Then
                w = wnd
                w.Show()
                w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                Exit Sub
            End If
        Next

        w = New wndError
        w.Show()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim w As New wndAbout With {.Owner = Me}
        w.ShowDialog()
    End Sub

    Private Sub ctxmnuMain_Opened(sender As Object, e As RoutedEventArgs) Handles ctxmnuMain.Opened
        ctxmnuMainExternalSoftware.Items.Clear()
        For Each es As clsExternalSoftware In preferences.ExternalSoftware
            Dim esmnu As New MenuItem
            esmnu.Header = es.Label
            esmnu.Icon = New Image With {.Source = es.Image}
            esmnu.Tag = es
            AddHandler esmnu.PreviewMouseDown, AddressOf ctxmnuMainExternalSoftwareItem_PreviewMouseDown
            ctxmnuMainExternalSoftware.Items.Add(esmnu)
        Next

        Dim singleselect As Boolean = False
        Dim user As Boolean = False
        Dim computer As Boolean = False

        If dgMain.SelectedItems.Count = 1 Then
            singleselect = True
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj IsNot Nothing Then
                If obj.SchemaClassName = "user" Then
                    user = True
                ElseIf obj.SchemaClassName = "computer" Then
                    computer = True
                End If
            End If
        End If

        ctxmnuMainExternalSoftware.Visibility = If(singleselect, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainCopy.Visibility = If(user Or singleselect = False, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainMove.Visibility = If(singleselect, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainRename.Visibility = If(singleselect, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainRemove.Visibility = If(singleselect, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainResetPassword.Visibility = If(user, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainDisableEnable.Visibility = If(user Or computer, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainExpirationDate.Visibility = If(user, Visibility.Visible, Visibility.Collapsed)
        ctxmnuMainShowInCommander.Visibility = If(singleselect, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub ctxmnuMainExternalSoftwareItem_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        If dgMain.SelectedItem Is Nothing Then Exit Sub
        Dim obj As clsDirectoryObject = dgMain.SelectedItem
        Dim esmnu As MenuItem = CType(sender, MenuItem)
        If esmnu Is Nothing Then Exit Sub
        Dim es As clsExternalSoftware = CType(esmnu.Tag, clsExternalSoftware)
        If es Is Nothing Then Exit Sub

        Dim args As String = es.Arguments
        If args Is Nothing Then args = ""

        Dim patterns As MatchCollection = Regex.Matches(args, "%(.*?)%")

        For Each pattern As Match In patterns
            Dim val As String = If(obj.CustomProperty(Replace(pattern.Value, "%", "")), pattern.Value).ToString
            args = Replace(args, pattern.Value, val)
        Next

        args = Replace(args, "%myusername%", obj.Domain.Username)
        args = Replace(args, "%mypassword%", obj.Domain.Password)
        args = Replace(args, "%mydomain%", obj.Domain.Name)

        Dim psi As New ProcessStartInfo(es.Path, args)

        If es.CurrentCredentials = True Then
            If e.RightButton = MouseButtonState.Pressed AndAlso IMsgBox(es.Path & vbCrLf & "от текущего пользователя," & vbCrLf & "с аргументами:" & vbCrLf & vbCrLf & Replace(args, obj.Domain.Password, "%mypassword%") & vbCrLf & vbCrLf & "продолжить выполнение?", "Запуск приложения", vbYesNo, vbQuestion) <> vbYes Then Exit Sub

            psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
            psi.UseShellExecute = False
            Process.Start(psi)
        Else
            If e.RightButton = MouseButtonState.Pressed AndAlso IMsgBox(es.Path & vbCrLf & "от пользователя, указанного в настройках домена," & vbCrLf & "с аргументами:" & vbCrLf & vbCrLf & Replace(args, obj.Domain.Password, "%mypassword%") & vbCrLf & vbCrLf & "продолжить выполнение?", "Запуск приложения", vbYesNo, vbQuestion) <> vbYes Then Exit Sub

            psi.Domain = obj.Domain.Name
            psi.UserName = obj.Domain.Username
            psi.Password = ToSecure(obj.Domain.Password)
            psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
            psi.UseShellExecute = False
            Process.Start(psi)
        End If
    End Sub

    Private Sub ctxmnuMainCopyDisplayName_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainCopyDisplayName.Click
        If dgMain.SelectedItems Is Nothing Then Exit Sub

        Try
            Clipboard.SetText(Join(dgMain.SelectedItems.Cast(Of clsDirectoryObject).ToArray.Select(Function(x) x.displayName).ToArray, vbCrLf))
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainCopyDisplayName_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainCopyBasicAttributes_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainCopyBasicAttributes.Click
        If dgMain.SelectedItems Is Nothing Then Exit Sub

        Try
            Clipboard.SetText(Join(dgMain.SelectedItems.Cast(Of clsDirectoryObject).ToArray.Select(Function(x) x.displayName & vbTab & x.userPrincipalNameName & vbTab & x.mail & vbTab & x.telephoneNumber).ToArray, vbCrLf))
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainCopyBasicAttributes_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainCopyAllAttributes_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainCopyAllAttributes.Click
        If dgMain.SelectedItems Is Nothing Then Exit Sub
        Try
            Dim clip As String = ""
            For Each ci As clsDataGridColumnInfo In preferences.Columns
                For Each cia As clsAttribute In ci.Attributes
                    clip &= cia.Label
                    clip &= vbTab
                Next
            Next
            clip &= vbCrLf
            For Each obj As clsDirectoryObject In dgMain.SelectedItems
                For Each ci As clsDataGridColumnInfo In preferences.Columns
                    If ci.Header = "⬕" Then Continue For
                    For Each cia As clsAttribute In ci.Attributes
                        For Each prop As PropertyInfo In GetType(clsDirectoryObject).GetProperties()
                            If prop.Name = cia.Name Then
                                If prop.PropertyType.IsArray Then
                                    Dim objArray() As Object = prop.GetValue(obj, Nothing)
                                    clip &= If(objArray IsNot Nothing, Join(objArray, " / "), "")
                                Else
                                    clip &= If(prop.GetValue(obj, Nothing), "")
                                End If
                            End If
                        Next
                        clip &= vbTab
                    Next
                Next
                clip &= vbCrLf
            Next
            Clipboard.SetText(clip)
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainCopyAllAttributes_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainSelectAll_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainSelectAll.Click
        dgMain.SelectAll()
    End Sub

    Private Sub ctxmnuMainCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainCreateObject.Click
        Dim w As New wndCreateObject With {.Owner = Me}
        w.Show()
    End Sub

    Private Sub ctxmnuMainMove_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainMove.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            Dim w As New wndDomainBrowser With {.Owner = Me}
            w.currentdomain = obj.Domain
            w.ShowDialog()
            If w.currentcontainer IsNot Nothing Then
                obj.Entry.MoveTo(w.currentcontainer.Entry)
            End If
            obj.NotifyMoved()
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainMove_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainRename_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainRename.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            Dim name As String = IInputBox("Введите новое имя объекта", "Переименование объекта", vbQuestion, obj.name)
            If Len(name) > 0 Then
                obj.Entry.Rename("cn=" & name)
                obj.Entry.CommitChanges()
                obj.NotifyRenamed()
                obj.NotifyMoved()
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainRename_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainRemove_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainRemove.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            If IMsgBox("Вы уверены?", "Удаление объекта", vbYesNo, vbQuestion) = vbYes Then
                Dim parent As DirectoryEntry = obj.Entry.Parent
                parent.Children.Remove(obj.Entry)
                parent.CommitChanges()
                objects.Remove(obj)
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainRemove_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainResetPassword_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainResetPassword.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            If IMsgBox("Вы уверены?", "Сброс пароля", vbYesNo, vbQuestion) = vbYes Then
                obj.ResetPassword()
                obj.passwordNeverExpires = False
                IMsgBox("Пароль сброшен.", "Сброс пароля", vbOKOnly, vbInformation)
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainResetPassword_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainDisableEnable_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainDisableEnable.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            If IMsgBox("Вы уверены?", If(obj.disabled, "Разблокирование объекта", "Блокирование объекта"), vbYesNo, vbQuestion) = vbYes Then
                obj.disabled = Not obj.disabled
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainDisableEnable_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainExpirationDate_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainExpirationDate.Click
        Try
            Dim obj As clsDirectoryObject = dgMain.SelectedItem
            If obj Is Nothing Then Exit Sub
            obj.accountExpiresDate = Today.AddDays(1)
            Dim w As Window = ShowDirectoryObjectProperties(obj, Me)
            If GetType(wndUser) Is w.GetType Then CType(w, wndUser).tabctlUser.SelectedIndex = 1
        Catch ex As Exception
            ThrowException(ex, "ctxmnuMainExpirationDate_Click")
        End Try
    End Sub

    Private Sub ctxmnuMainShowInCommanderLeftPanel_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainShowInCommanderLeftPanel.Click
        If dgMain.SelectedItem Is Nothing Then Exit Sub

        Dim w As wndCommander
        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndCommander) Is wnd.GetType Then
                w = wnd
                w.Show() : w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                w.Topmost = True : w.Topmost = False
                w.pnlLeft.ShowObject(dgMain.SelectedItem)
                Exit Sub
            End If
        Next

        w = New wndCommander
        w.Owner = Me

        w.Show()
        w.pnlLeft.ShowObject(dgMain.SelectedItem)
    End Sub

    Private Sub ctxmnuMainShowInCommanderRightPanel_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuMainShowInCommanderRightPanel.Click
        If dgMain.SelectedItem Is Nothing Then Exit Sub

        Dim w As wndCommander
        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndCommander) Is wnd.GetType Then
                w = wnd
                w.Show() : w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                w.Topmost = True : w.Topmost = False
                w.pnlRight.ShowObject(dgMain.SelectedItem)
                Exit Sub
            End If
        Next

        w = New wndCommander
        w.Owner = Me

        w.Show()
        w.pnlRight.ShowObject(dgMain.SelectedItem)
    End Sub

    Private Sub ctxmnuProperties_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuProperties.Click
        If dgMain.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(dgMain.SelectedItem, Me)
    End Sub

    Private Sub mnuFilePrint_Click(sender As Object, e As RoutedEventArgs) Handles mnuFilePrint.Click
        Dim wnd As New wndPrintPreview With {.Owner = Me}
        wnd.objects = objects
        wnd.ShowDialog()
    End Sub

#End Region

End Class
