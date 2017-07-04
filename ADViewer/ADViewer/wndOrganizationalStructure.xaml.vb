
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.Threading.Tasks

Public Class wndOrganizationalStructure

    Private sosurceitem As TreeViewItem
    Private sourceobject As Object
    Private allowdrag As Boolean

    Private Async Sub wndOrganizationalStructure_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        cap.Visibility = Visibility.Visible

        Dim managers As ObservableCollection(Of clsDirectoryObject)

        Await Task.Run(Sub() managers = Search(domains(1)))

        tvEmployees.Items.Clear()

        For Each manager In Search(domains(1))
            tvEmployees.Items.Add(manager)
        Next

        cap.Visibility = Visibility.Hidden
    End Sub


    Private Sub tv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles tvEmployees.PreviewMouseLeftButtonDown
        Dim treeView As TreeView = TryCast(sender, TreeView)
        allowdrag = e.GetPosition(sender).X < treeView.ActualWidth - SystemParameters.VerticalScrollBarWidth And e.GetPosition(sender).Y < treeView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    End Sub

    Private Sub tv_MouseMove(sender As Object, e As MouseEventArgs) Handles tvEmployees.MouseMove
        Dim treeView As TreeView = TryCast(sender, TreeView)

        If e.LeftButton = MouseButtonState.Pressed And treeView.SelectedItem IsNot Nothing And allowdrag Then
            sourceobject = treeView

            Dim obj As clsDirectoryObject = CType(treeView.SelectedItem, clsDirectoryObject)
            Dim dragData As New DataObject("clsDirectoryObject", obj)

            DragDrop.DoDragDrop(treeView, dragData, DragDropEffects.Move)
        End If
    End Sub

    Private Sub tv_DragEnter(sender As Object, e As DragEventArgs) Handles tvEmployees.DragEnter
        If Not e.Data.GetDataPresent("clsDirectoryObject") Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub tv_Drop(sender As Object, e As DragEventArgs) Handles tvEmployees.Drop
        If e.Data.GetDataPresent("clsDirectoryObject") Then

            cap.Visibility = Visibility.Visible

            Dim obj As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            tvEmployees.Items.Remove(obj)

            Dim targetItem As TreeViewItem = GetNearestContainer(TryCast(e.OriginalSource, UIElement))
            Dim draggedItem = FindVisualParent(Of TreeViewItem)(TryCast(tvEmployees.SelectedItem, UIElement))

            If targetItem IsNot Nothing AndAlso obj IsNot Nothing Then

                tvEmployees.Items.Remove(draggedItem)
                targetItem.Items.Add(obj)

                e.Effects = DragDropEffects.Move
            End If

            cap.Visibility = Visibility.Hidden

        End If
    End Sub


    Private Function GetNearestContainer(element As UIElement) As TreeViewItem
        ' Walk up the element tree to the nearest tree view item.
        Dim container As TreeViewItem = TryCast(element, TreeViewItem)
        While (container Is Nothing) AndAlso (element IsNot Nothing)
            element = TryCast(VisualTreeHelper.GetParent(element), UIElement)
            container = TryCast(element, TreeViewItem)
        End While
        Return container
    End Function

    Private Function FindVisualParent(Of TObject As UIElement)(child As UIElement) As TObject
        If child Is Nothing Then
            Return Nothing
        End If

        Dim parent As UIElement = TryCast(VisualTreeHelper.GetParent(child), UIElement)

        While parent IsNot Nothing
            Dim found As TObject = TryCast(parent, TObject)
            If found IsNot Nothing Then
                Return found
            Else
                parent = TryCast(VisualTreeHelper.GetParent(parent), UIElement)
            End If
        End While

        Return Nothing
    End Function










    Private Function Search(dmn As clsDomain) As ObservableCollection(Of clsDirectoryObject)
        Dim resultlist As New ObservableCollection(Of clsDirectoryObject)

        Dim properties As String() = {"objectGuid",
                                        "userAccountControl",
                                        "accountExpires",
                                        "name",
                                        "distinguishedName",
                                        "pwdLastSet",
                                        "manager"}
        Try
            Dim LDAPsearcher As New DirectorySearcher(dmn.DefaultNamingContext)
            Dim LDAPresults As SearchResultCollection = Nothing
            Dim Filter As String

            LDAPsearcher.PropertiesToLoad.Clear()
            LDAPsearcher.PropertiesToLoad.AddRange(properties)

            Filter = "(&" +
                            "(|" +
                                "(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(objectClass=contact))(!(UserAccountControl:1.2.840.113556.1.4.803:=2)))" +
                            ")" +
                            "(|" +
                                "(!manager=*)" +
                            ")" +
                         ")"

            LDAPsearcher.Filter = Filter
            LDAPsearcher.PageSize = 1000

            Dim searchResultCollection As SearchResultCollection = LDAPsearcher.FindAll()


            For Each result As SearchResult In searchResultCollection
                Dim obj As New clsDirectoryObject(result.GetDirectoryEntry, dmn)
                If obj.Employees.Count > 0 Then resultlist.Add(obj)
            Next

        Catch ex As Exception
            ThrowException(ex, dmn.Name)
        End Try

        Return resultlist
    End Function

End Class
