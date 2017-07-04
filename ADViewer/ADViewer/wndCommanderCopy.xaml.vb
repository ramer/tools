Imports System.Collections.ObjectModel
Imports System.DirectoryServices
Imports System.Threading.Tasks

Public Class wndCommanderCopy

    Public Property sourceobjects As clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property destination As clsDirectoryObject

    Private UserPrincipalNamePattern As String
    Private UserPrincipalNameDomain As String
    Private AddDefaultGroups As Boolean

    Private Sub wndCommanderCopy_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDestination.Text = destination.Entry.Path
        cmboUserPrincipalNamePattern.ItemsSource = destination.Domain.UsernamePattern.Split({",", vbCr, vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) Trim(x)).ToArray
        cmboUserPrincipalNameDomain.ItemsSource = destination.Domain.Suffixes
    End Sub

    Private Async Sub btnCopy_Click(sender As Object, e As RoutedEventArgs) Handles btnCopy.Click
        cap.Visibility = Visibility.Visible

        UserPrincipalNamePattern = cmboUserPrincipalNamePattern.Text
        UserPrincipalNameDomain = cmboUserPrincipalNameDomain.Text
        AddDefaultGroups = chbAddDefaultGroups.IsChecked

        If rbCopyObjectsAndContainers.IsChecked Then
            If destination.Domain.DefaultPassword = "" Then IMsgBox("Стандартный пароль в целевом домене не указан", "Ошибка", vbOKOnly, vbExclamation) : cap.Visibility = Visibility.Hidden : Exit Sub
            If destination.Domain Is sourceobjects(0).Domain AndAlso (UserPrincipalNamePattern = "" Or UserPrincipalNameDomain = "") Then IMsgBox("Одинаковые имена входа в одном домене недопустимы - укажите шаблон именования", "Ошибка", vbOKOnly, vbExclamation) : cap.Visibility = Visibility.Hidden : Exit Sub
            Await Task.Run(Sub() Copy(sourceobjects.ToArray, destination, True, True))
        ElseIf rbCopyObjects.IsChecked Then
            If destination.Domain.DefaultPassword = "" Then IMsgBox("Стандартный пароль в целевом домене не указан", "Ошибка", vbOKOnly, vbExclamation) : cap.Visibility = Visibility.Hidden : Exit Sub
            If destination.Domain Is sourceobjects(0).Domain AndAlso (UserPrincipalNamePattern = "" Or UserPrincipalNameDomain = "") Then IMsgBox("Одинаковые имена входа в одном домене недопустимы - укажите шаблон именования", "Ошибка", vbOKOnly, vbExclamation) : cap.Visibility = Visibility.Hidden : Exit Sub
            Await Task.Run(Sub() Copy(sourceobjects.ToArray, destination, True, False))
        ElseIf rbCopyContainers.IsChecked Then
            Await Task.Run(Sub() Copy(sourceobjects.ToArray, destination, False, True))
        End If

        cap.Visibility = Visibility.Hidden
        Me.Close()
    End Sub

    Private Sub Copy(sourceobjects As clsDirectoryObject(), destination As clsDirectoryObject, objectsflag As Boolean, containersflag As Boolean)
        For Each currentobject As clsDirectoryObject In sourceobjects
            If objectsflag AndAlso currentobject.SchemaClassName = "user" Then
                Dim newuser As clsDirectoryObject = Nothing

                Try
                    newuser = New clsDirectoryObject(destination.Entry.Children.Add(currentobject.Entry.Name, "user"), destination.Domain)

                    If UserPrincipalNamePattern = "" Or UserPrincipalNameDomain = "" Then
                        newuser.sAMAccountName = currentobject.sAMAccountName
                        newuser.userPrincipalName = currentobject.userPrincipalName
                    Else
                        Dim name As String = GetNextDomainUser(UserPrincipalNamePattern, destination.Domain)
                        If name Is Nothing Then Continue For
                        newuser.sAMAccountName = name
                        newuser.userPrincipalName = name & "@" & UserPrincipalNameDomain
                    End If

                    newuser.givenName = currentobject.givenName
                    newuser.initials = currentobject.initials
                    newuser.sn = currentobject.sn
                    newuser.displayName = currentobject.displayName
                    newuser.description = currentobject.description
                    newuser.physicalDeliveryOfficeName = currentobject.physicalDeliveryOfficeName
                    newuser.telephoneNumber = currentobject.telephoneNumber
                    newuser.homePhone = currentobject.homePhone
                    newuser.ipPhone = currentobject.ipPhone
                    newuser.mobile = currentobject.mobile
                    newuser.mail = currentobject.mail
                    newuser.title = currentobject.title
                    newuser.department = currentobject.department
                    newuser.company = currentobject.company
                    newuser.streetAddress = newuser.streetAddress

                    newuser.Entry.CommitChanges()

                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Create User", currentobject.Entry.Name, ex))
                    Continue For
                End Try

                Try
                    newuser.ResetPassword()
                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Set User Password", currentobject.Entry.Name, ex))
                End Try

                Try
                    newuser.userAccountControl = ADS_UF_NORMAL_ACCOUNT
                    newuser.userMustChangePasswordNextLogon = True

                    newuser.Entry.CommitChanges()
                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Set User Account control", currentobject.Entry.Name, ex))
                End Try

                If AddDefaultGroups Then
                    Try
                        For Each group As clsDirectoryObject In destination.Domain.DefaultGroups
                            group.Entry.Invoke("Add", newuser.Entry.Path)
                            group.Entry.CommitChanges()
                            Threading.Thread.Sleep(500)
                        Next
                    Catch ex As Exception
                        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Set User memberof attributes", currentobject.Entry.Name, ex))
                    End Try
                End If

            ElseIf objectsflag AndAlso currentobject.SchemaClassName = "group" Then

                Dim newgroup As clsDirectoryObject = Nothing

                Try
                    newgroup = New clsDirectoryObject(destination.Entry.Children.Add(currentobject.Entry.Name, "group"), destination.Domain)
                    newgroup.sAMAccountName = currentobject.sAMAccountName
                    newgroup.description = currentobject.description
                    newgroup.mail = currentobject.mail
                    newgroup.Entry.CommitChanges()
                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Create Group", currentobject.Entry.Name, ex))
                End Try

                If currentobject.groupTypeScopeDomainLocal Then ' domain local group, but unversal first
                    Try
                        newgroup.groupType = ADS_GROUP_TYPE_UNIVERSAL_GROUP
                        newgroup.Entry.CommitChanges()

                        newgroup.groupType = currentobject.groupType
                        newgroup.Entry.CommitChanges()
                    Catch ex As Exception
                        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Set Group type", currentobject.Entry.Name, ex))
                    End Try
                Else
                    Try
                        newgroup.groupType = currentobject.groupType
                        newgroup.Entry.CommitChanges()
                    Catch ex As Exception
                        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Set Group type", currentobject.Entry.Name, ex))
                    End Try
                End If

            ElseIf objectsflag AndAlso currentobject.SchemaClassName = "contact" Then

                Dim newcontact As clsDirectoryObject = Nothing

                Try
                    newcontact = New clsDirectoryObject(destination.Entry.Children.Add(currentobject.Entry.Name, "contact"), destination.Domain)

                    newcontact.givenName = currentobject.givenName
                    newcontact.initials = currentobject.initials
                    newcontact.sn = currentobject.sn
                    newcontact.displayName = currentobject.displayName
                    newcontact.description = currentobject.description
                    newcontact.physicalDeliveryOfficeName = currentobject.physicalDeliveryOfficeName
                    newcontact.telephoneNumber = currentobject.telephoneNumber
                    newcontact.homePhone = currentobject.homePhone
                    newcontact.ipPhone = currentobject.ipPhone
                    newcontact.mobile = currentobject.mobile
                    newcontact.mail = currentobject.mail
                    newcontact.title = currentobject.title
                    newcontact.department = currentobject.department
                    newcontact.company = currentobject.company
                    newcontact.streetAddress = currentobject.streetAddress

                    newcontact.Entry.CommitChanges()
                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Create Contact", currentobject.Entry.Name, ex))
                End Try

            ElseIf containersflag AndAlso (currentobject.SchemaClassName = "organizationalUnit" Or currentobject.SchemaClassName = "container") Then

                Dim newou As DirectoryEntry = Nothing
                Try
                    newou = destination.Entry.Children.Add(currentobject.Entry.Name, "organizationalUnit")
                    newou.CommitChanges()
                Catch ex As Exception
                    SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog("Create Container", newou.Name, ex))
                End Try

                If currentobject.Entry.Children.Cast(Of DirectoryEntry).Count > 0 Then
                    Copy(currentobject.Entry.Children.Cast(Of DirectoryEntry).Select(Function(x As DirectoryEntry) New clsDirectoryObject(x, currentobject.Domain)).ToArray,
                         New clsDirectoryObject(newou, destination.Domain), objectsflag, containersflag)
                End If
            End If
        Next
    End Sub


End Class
