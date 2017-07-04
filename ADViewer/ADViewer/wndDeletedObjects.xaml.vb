Imports System.Collections.ObjectModel
Imports System.DirectoryServices
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Threading.Tasks

Public Class wndDeletedObjects
    Public WithEvents searcher As New clsSearcher

    Public Property foundobjects As New clsThreadSafeObservableCollection(Of clsDeletedDirectoryObject)

    Private Sub wndDeletedObjects_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        dgDeletedObjects.ItemsSource = foundobjects
        cmboDomains.ItemsSource = domains
    End Sub

    Private Sub dgDeletedObjects_LoadingRow(sender As Object, e As DataGridRowEventArgs) Handles dgDeletedObjects.LoadingRow
        e.Row.Header = (e.Row.GetIndex + 1).ToString
    End Sub

    Private Async Sub Search(Optional pattern As String = """*""", Optional domain As clsDomain = Nothing)
        tbSearchPattern.Text = pattern
        tbSearchPattern.SelectAll()

        cap.Visibility = Visibility.Visible

        Await searcher.TombstoneSearchAsync(foundobjects, pattern, domain)

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub RestoreObject(currentDeletedObject As clsDeletedDirectoryObject, primarydata As Boolean, enableobject As Boolean, defaultgroups As Boolean)
        If currentDeletedObject.lastKnownParentExist = False Then
            ThrowCustomException("Родительский объект не существует")
            Exit Sub
        End If

        Try

            Dim dam As New DirectoryAttributeModification()
            dam.Name = "isDeleted"
            dam.Operation = DirectoryAttributeOperation.Delete

            Dim dn As String = String.Format("{0}={1},{2}", If(currentDeletedObject.SchemaClassName = "organizationalunit", "OU", "CN"), currentDeletedObject.nameFormated, currentDeletedObject.lastKnownParent)

            Dim dam2 As New DirectoryAttributeModification()
            dam2.Name = "distinguishedName"
            dam2.Operation = DirectoryAttributeOperation.Replace
            dam2.Add(dn)

            Dim mr As New ModifyRequest(currentDeletedObject.distinguishedName, New DirectoryAttributeModification() {dam, dam2})
            mr.Controls.Add(New ShowDeletedControl())

            Dim conn As New LdapConnection(New LdapDirectoryIdentifier(currentDeletedObject.Domain.Name), New NetworkCredential(currentDeletedObject.Domain.Username, currentDeletedObject.Domain.Password), AuthType.Negotiate)
            Using conn
                conn.Bind()
                conn.SessionOptions.ProtocolVersion = 3
                Dim resp As ModifyResponse = DirectCast(conn.SendRequest(mr), ModifyResponse)

                If resp.ResultCode <> ResultCode.Success Then
                    ThrowCustomException(resp.ErrorMessage)
                    Exit Sub
                End If
            End Using

            Dim currentObject As New clsDirectoryObject(New DirectoryEntry("LDAP://" & currentDeletedObject.Domain.Name & "/" & dn, currentDeletedObject.Domain.Username, currentDeletedObject.Domain.Password), currentDeletedObject.Domain)

            If currentObject.SchemaClassName = "user" Then
                Try
                    currentObject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set User Password")
                End Try

                Try
                    If primarydata Then
                        currentObject.displayName = currentObject.name
                        currentObject.givenName = If(Split(currentObject.name, " ").Count >= 2, Split(currentObject.name, " ")(1), "")
                        currentObject.sn = If(Split(currentObject.name, " ").Count >= 1, Split(currentObject.name, " ")(0), "")
                    End If

                    If enableobject Then
                        currentObject.disabled = False
                    End If

                    currentObject.userPrincipalName = currentObject.sAMAccountName & "@" & LCase(currentObject.Domain.Name)
                    currentObject.userMustChangePasswordNextLogon = True
                    currentObject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Set User attributes")
                End Try

                If defaultgroups Then
                    Try
                        For Each group As clsDirectoryObject In currentObject.Domain.DefaultGroups
                            group.Entry.Invoke("Add", currentObject.Entry.Path)
                            group.Entry.CommitChanges()
                            currentObject.memberOf.Add(group)
                            Threading.Thread.Sleep(500)
                        Next
                    Catch ex As Exception
                        ThrowException(ex, "Set User memberof attributes")
                    End Try
                End If
            ElseIf currentObject.SchemaClassName = "computer" Then
                Try
                    If enableobject Then
                        currentObject.disabled = False
                    End If

                    currentObject.Entry.CommitChanges()
                Catch ex As Exception
                    ThrowException(ex, "Set Computer attributes")
                End Try
            End If

        Catch ex As Exception
            ThrowException(ex, "RestoreObject")
        End Try
    End Sub

    Private Async Sub btnRestore_Click(sender As Object, e As RoutedEventArgs) Handles btnRestore.Click
        If dgDeletedObjects.SelectedItem Is Nothing Then Exit Sub

        Dim deletedobjects As New clsThreadSafeObservableCollection(Of clsDeletedDirectoryObject)
        Dim primarydata As Boolean = chbRestorePrimaryData.IsChecked
        Dim enableobject As Boolean = chbRestoreEnableObject.IsChecked
        Dim defaultgroups As Boolean = chbRestoreDefaultGroups.IsChecked

        Dim str As String = ""
        For Each obj As clsDeletedDirectoryObject In dgDeletedObjects.SelectedItems
            deletedobjects.Add(obj)
            Str &= obj.name.Split(New Char() {Chr(10)})(0) & vbTab & obj.sAMAccountName & vbCrLf
        Next
        Clipboard.SetText(str)

        cap.Visibility = Visibility.Visible

        Await Task.Run(
            Sub()
                For Each obj As clsDeletedDirectoryObject In deletedobjects
                    RestoreObject(obj, primarydata, enableobject, defaultgroups)
                Next
            End Sub)

        cap.Visibility = Visibility.Hidden

        If cmboDomains.SelectedItem IsNot Nothing Then
            Search(tbSearchPattern.Text, CType(cmboDomains.SelectedItem, clsDomain))
        End If
    End Sub

    Private Sub tbSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles tbSearchPattern.KeyDown
        If e.Key = Key.Enter And cmboDomains.SelectedItem IsNot Nothing Then
            Search(tbSearchPattern.Text, CType(cmboDomains.SelectedItem, clsDomain))
        End If
    End Sub

End Class
