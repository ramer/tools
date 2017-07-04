Imports System.ComponentModel
Imports System.Threading.Tasks

Public Class wndCreateObject
    Private Property _objectissharedmailbox As Boolean
    Private Property _objectdomain As clsDomain
    Private Property _objectcontainer As clsDirectoryObject
    Private Property _objectdisplayname As String
    Private Property _objectuserprincipalname As String
    Private Property _objectuserprincipalnamedomain As String
    Private Property _objectname As String
    Private Property _objectsamaccountname As String

    Private Sub wndCreateObject_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataContext = Me
        cmboUserDomain.ItemsSource = domains
        cmboComputerDomain.ItemsSource = domains
        cmboGroupDomain.ItemsSource = domains
        cmboContactDomain.ItemsSource = domains
    End Sub

    Private Sub cmboUserDomain_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboUserDomain.SelectionChanged
        cmboUserUserPrincipalName.ItemsSource = GetNextDomainUsers(CType(cmboUserDomain.SelectedValue, clsDomain))
    End Sub

    Private Sub cmboComputerDomain_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboComputerDomain.SelectionChanged
        cmboComputerObjectName.ItemsSource = GetNextDomainComputers(CType(cmboComputerDomain.SelectedValue, clsDomain))
    End Sub

    Private Sub btnUserContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnUserContainerBrowse.Click
        If cmboUserDomain.SelectedItem Is Nothing Then Exit Sub
        Dim w As New wndDomainBrowser With {.Owner = Me}
        w.currentdomain = CType(cmboUserDomain.SelectedItem, clsDomain)
        w.ShowDialog()
        If w.currentcontainer IsNot Nothing Then
            tbUserContainer.Tag = w.currentcontainer
            tbUserContainer.Text = w.currentcontainer.Entry.Path
        End If
    End Sub

    Private Sub btnComputerContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnComputerContainerBrowse.Click
        If cmboComputerDomain.SelectedItem Is Nothing Then Exit Sub
        Dim w As New wndDomainBrowser With {.Owner = Me}
        w.currentdomain = CType(cmboComputerDomain.SelectedItem, clsDomain)
        w.ShowDialog()
        If w.currentcontainer IsNot Nothing Then
            tbComputerContainer.Tag = w.currentcontainer
            tbComputerContainer.Text = w.currentcontainer.Entry.Path
        End If
    End Sub

    Private Sub btnGroupContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnGroupContainerBrowse.Click
        If cmboGroupDomain.SelectedItem Is Nothing Then Exit Sub
        Dim w As New wndDomainBrowser With {.Owner = Me}
        w.currentdomain = CType(cmboGroupDomain.SelectedItem, clsDomain)
        w.ShowDialog()
        If w.currentcontainer IsNot Nothing Then
            tbGroupContainer.Tag = w.currentcontainer
            tbGroupContainer.Text = w.currentcontainer.Entry.Path
        End If
    End Sub

    Private Sub btnContactContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnContactContainerBrowse.Click
        If cmboContactDomain.SelectedItem Is Nothing Then Exit Sub
        Dim w As New wndDomainBrowser With {.Owner = Me}
        w.currentdomain = CType(cmboContactDomain.SelectedItem, clsDomain)
        w.ShowDialog()
        If w.currentcontainer IsNot Nothing Then
            tbContactContainer.Tag = w.currentcontainer
            tbContactContainer.Text = w.currentcontainer.Entry.Path
        End If
    End Sub

    Private Sub tbUserDisplayname_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tbUserDisplayname.TextChanged
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Sub cmboUserUserPrincipalName_TextChanged(sender As Object, e As TextChangedEventArgs)
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Sub chbUserSharedMailbox_CheckedUnchecked(sender As Object, e As RoutedEventArgs) Handles chbUserSharedMailbox.Checked, chbUserSharedMailbox.Unchecked
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Async Sub btnCreate_Click(sender As Object, e As RoutedEventArgs) Handles btnCreate.Click
        Dim obj As clsDirectoryObject = Nothing

        cap.Visibility = Visibility.Visible

        If tabctlObject.SelectedIndex = 0 Then
            obj = Await CreateUser()
        ElseIf tabctlObject.SelectedIndex = 1 Then
            obj = Await CreateComputer()
        ElseIf tabctlObject.SelectedIndex = 2 Then
            obj = Await CreateGroup()
        ElseIf tabctlObject.SelectedIndex = 3 Then
            obj = Await CreateContact()
        End If

        cap.Visibility = Visibility.Hidden

        If obj IsNot Nothing Then
            ShowDirectoryObjectProperties(obj, Me.Owner)
            Me.Close()
        End If
    End Sub

    Private Sub wndCreateObject_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Async Function CreateUser() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdomain = cmboUserDomain.SelectedItem
        _objectcontainer = tbUserContainer.Tag
        _objectdisplayname = tbUserDisplayname.Text
        _objectuserprincipalname = cmboUserUserPrincipalName.Text
        _objectuserprincipalnamedomain = cmboUserUserPrincipalNameDomain.Text
        _objectname = tbUserObjectName.Text
        _objectsamaccountname = tbUserSamAccountName.Text

        If _objectdomain Is Nothing Or
           _objectcontainer Is Nothing Or
           _objectdisplayname = "" Or
           _objectuserprincipalname = "" Or
           _objectuserprincipalnamedomain = "" Or
           _objectname = "" Or
        _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        _objectissharedmailbox = chbUserSharedMailbox.IsChecked

        If _objectdomain.DefaultPassword = "" Then ThrowCustomException("Стандартный пароль не указан") : Return Nothing

        Await Task.Run(
            Sub()

                Try
                    _currentobject = New clsDirectoryObject(_objectcontainer.Entry.Children.Add("cn=" & _objectname, "user"), _objectdomain)
                    _currentobject.sAMAccountName = _objectsamaccountname
                    _currentobject.userPrincipalName = _objectuserprincipalname & "@" & _objectuserprincipalnamedomain
                Catch ex As Exception
                    ThrowException(ex, "Create User")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set User Password")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.displayName = _objectdisplayname

                    If Not _objectissharedmailbox Then ' user
                        _currentobject.givenName = If(Split(_objectdisplayname, " ").Count >= 2, Split(_objectdisplayname, " ")(1), "")
                        _currentobject.sn = If(Split(_objectdisplayname, " ").Count >= 1, Split(_objectdisplayname, " ")(0), "")
                        _currentobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT
                        _currentobject.userMustChangePasswordNextLogon = True
                    Else                                       ' sharedmailbox
                        _currentobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT + ADS_UF_ACCOUNTDISABLE
                        _currentobject.userMustChangePasswordNextLogon = True
                    End If

                Catch ex As Exception
                    ThrowException(ex, "Set User attributes")
                End Try

                Threading.Thread.Sleep(500)

                If Not _objectissharedmailbox Then ' user
                    Try
                        For Each group As clsDirectoryObject In _objectdomain.DefaultGroups
                            group.Entry.Invoke("Add", _currentobject.Entry.Path)
                            group.Entry.CommitChanges()
                            Threading.Thread.Sleep(500)
                        Next
                    Catch ex As Exception
                        ThrowException(ex, "Set User memberof attributes")
                    End Try
                End If

                Threading.Thread.Sleep(500)

            End Sub)

        _currentobject.Refresh()

        Return _currentobject
    End Function

    Private Async Function CreateComputer() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdomain = cmboComputerDomain.SelectedItem
        _objectcontainer = tbComputerContainer.Tag
        _objectname = cmboComputerObjectName.Text
        _objectsamaccountname = tbComputerSamAccountName.Text

        If _objectdomain Is Nothing Or
           _objectcontainer Is Nothing Or
           _objectname = "" Or
           _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    _currentobject = New clsDirectoryObject(_objectcontainer.Entry.Children.Add("cn=" & _objectname, "computer"), _objectdomain)
                    _currentobject.sAMAccountName = _objectsamaccountname
                    _currentobject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Create Computer")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set Computer Password")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.userAccountControl = ADS_UF_WORKSTATION_TRUST_ACCOUNT + ADS_UF_PASSWD_NOTREQD
                    _currentobject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Set Computer attributes")
                End Try

                Threading.Thread.Sleep(500)

            End Sub)

        Return _currentobject
    End Function

    Private Async Function CreateGroup() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdomain = cmboGroupDomain.SelectedItem
        _objectcontainer = tbGroupContainer.Tag
        _objectname = tbGroupObjectName.Text
        _objectsamaccountname = tbGroupSamAccountName.Text

        If _objectdomain Is Nothing Or
           _objectcontainer Is Nothing Or
           _objectname = "" Or
           _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    _currentobject = New clsDirectoryObject(_objectcontainer.Entry.Children.Add("cn=" & _objectname, "Group"), _objectdomain)
                    _currentobject.sAMAccountName = _objectsamaccountname
                    _currentobject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Create Group")
                End Try

                Dim grouptype As Long = 0
                grouptype += If(rbGroupScopeDomainLocal.IsChecked, ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP, 0)
                grouptype += If(rbGroupScopeGlobal.IsChecked, ADS_GROUP_TYPE_GLOBAL_GROUP, 0)
                grouptype += If(rbGroupScopeUniversal.IsChecked, ADS_GROUP_TYPE_UNIVERSAL_GROUP, 0)
                grouptype += If(rbGroupTypeSecurity.IsChecked, ADS_GROUP_TYPE_SECURITY_ENABLED, 0)

                If rbGroupScopeDomainLocal.IsChecked Then ' domain local group, but unversal first
                    Try
                        _currentobject.groupType = ADS_GROUP_TYPE_UNIVERSAL_GROUP
                        _currentobject.Entry.CommitChanges()

                        _currentobject.groupType = grouptype
                        _currentobject.Entry.CommitChanges()
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                Else
                    Try
                        _currentobject.groupType = grouptype
                        _currentobject.Entry.CommitChanges()
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                End If

                Threading.Thread.Sleep(500)

            End Sub)

        Return _currentobject
    End Function

    Private Async Function CreateContact() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdomain = cmboContactDomain.SelectedItem
        _objectcontainer = tbContactContainer.Tag
        _objectdisplayname = tbContactDisplayname.Text
        _objectname = tbContactObjectName.Text

        If _objectdomain Is Nothing Or
           _objectcontainer Is Nothing Or
           _objectdisplayname = "" Or
           _objectname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    _currentobject = New clsDirectoryObject(_objectcontainer.Entry.Children.Add("cn=" & _objectname, "contact"), _objectdomain)
                    _currentobject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Create Contact")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.displayName = _objectdisplayname
                    _currentobject.givenName = If(Split(_objectdisplayname, " ").Count >= 2, Split(_objectdisplayname, " ")(1), "")
                    _currentobject.sn = If(Split(_objectdisplayname, " ").Count >= 1, Split(_objectdisplayname, " ")(0), "")

                    _currentobject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Set Contact attributes")
                End Try

                Threading.Thread.Sleep(500)

            End Sub)

        _currentobject.Refresh()

        Return _currentobject
    End Function


End Class
