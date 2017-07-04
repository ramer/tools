
Imports System.Collections.ObjectModel

Public Class ctlGroupMemberOf


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlGroupMemberOf),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))
    Public Shared ReadOnly CurrentDomainProperty As DependencyProperty = DependencyProperty.Register("CurrentDomain",
                                                    GetType(clsDomain),
                                                    GetType(ctlGroupMemberOf),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf CurrentDomainPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomain As clsDomain
    Private Property _currentdomaingroups As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    WithEvents searcher As New clsSearcher

    Private sourceobject As Object
    Private allowdrag As Boolean

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Public Property CurrentDomain() As clsDirectoryObject
        Get
            Return GetValue(CurrentDomainProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentDomainProperty, value)
        End Set
    End Property

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMemberOf = CType(d, ctlGroupMemberOf)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomain = CType(e.NewValue, clsDirectoryObject).Domain
            ._currentdomaingroups.Clear()
            .lvSelectedGroups.ItemsSource = If(._currentobject IsNot Nothing, ._currentobject.memberOf, Nothing)
            .lvDomainGroups.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomaingroups, Nothing)
        End With
    End Sub

    Private Shared Sub CurrentDomainPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMemberOf = CType(d, ctlGroupMemberOf)
        With instance
            ._currentobject = Nothing
            ._currentdomain = CType(e.NewValue, clsDomain)
            ._currentdomaingroups.Clear()
            .lvSelectedGroups.ItemsSource = If(._currentdomain IsNot Nothing, ._currentdomain.DefaultGroups, Nothing)
            .lvDomainGroups.ItemsSource = If(._currentdomain IsNot Nothing, ._currentdomaingroups, Nothing)
        End With
    End Sub

    Private Function mode() As Integer
        If _currentobject IsNot Nothing Then
            Return 0
        ElseIf _currentdomain IsNot Nothing Then
            Return 1
        Else
            Return -1
        End If
    End Function

    Private Async Sub tbDomainGroupsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainGroupsFilter.KeyDown
        If e.Key = Key.Enter Then
            Await searcher.BasicSearchAsync(_currentdomaingroups, tbDomainGroupsFilter.Text, _currentdomain,, False, False, True, True)
        End If
    End Sub

    Private Sub lvDomainGroups_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvDomainGroups.MouseDoubleClick
        If lvDomainGroups.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvDomainGroups.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSelectedGroups_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvSelectedGroups.MouseDoubleClick
        If lvSelectedGroups.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSelectedGroups.SelectedItem, Window.GetWindow(Me))
    End Sub


    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvSelectedGroups.PreviewMouseLeftButtonDown,
                                                                                                   lvDomainGroups.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        allowdrag = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth 'And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedGroups.MouseMove,
                                                                            lvDomainGroups.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And listView.SelectedItem IsNot Nothing And allowdrag Then
            sourceobject = listView

            Dim obj As clsDirectoryObject = CType(listView.SelectedItem, clsDirectoryObject)
            Dim dragData As New DataObject("clsDirectoryObject", obj)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Move)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.DragEnter,
                                                                            lvDomainGroups.DragEnter

        If Not e.Data.GetDataPresent("clsDirectoryObject") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.Drop,
                                                                        lvDomainGroups.Drop

        If e.Data.GetDataPresent("clsDirectoryObject") And sender IsNot sourceobject Then
            Dim draggedObject As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            If sender Is lvSelectedGroups Then ' adding member
                If draggedObject Is Nothing Then Exit Sub
                AddMember(draggedObject)
            Else
                If draggedObject Is Nothing Then Exit Sub
                RemoveMember(draggedObject)
            End If

        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try

            If mode() = 0 Then
                For Each group As clsDirectoryObject In _currentobject.memberOf
                    If group.name = [object].name Then Exit Sub
                Next
                [object].Entry.Invoke("Add", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()

                _currentobject.memberOf.Add([object])
                lvSelectedGroups.ItemsSource = Nothing 'ultra dirty hack, without it LV doesn't updating
                lvSelectedGroups.ItemsSource = _currentobject.memberOf
            ElseIf mode() = 1 Then
                For Each group As clsDirectoryObject In _currentdomain.DefaultGroups
                    If group.name = [object].name Then Exit Sub
                Next
                _currentdomain.DefaultGroups.Add([object])
            End If

        Catch ex As Exception
            ThrowException(ex, "AddMember")
            If [object].SchemaClassName = "group" Then ShowWrongMemberMessage()
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try

            If mode() = 0 Then
                [object].Entry.Invoke("Remove", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()

                _currentobject.memberOf.Remove([object])

                lvSelectedGroups.ItemsSource = Nothing 'ultra dirty hack, without it LV doesn't updating
                lvSelectedGroups.ItemsSource = _currentobject.memberOf
            ElseIf mode() = 1 Then
                _currentdomain.DefaultGroups.Remove([object])
            End If

        Catch ex As Exception
            ThrowException(ex, "RemoveMember")
        End Try
    End Sub

    Private Sub ctlGroupMemberOf_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainGroupsFilter.Focus()
        btnDefaultGroups.Visibility = If(mode() = 0, Visibility.Visible, Visibility.Hidden)
    End Sub

    Private Sub btnDefaultGroups_Click(sender As Object, e As RoutedEventArgs) Handles btnDefaultGroups.Click
        If mode() <> 0 Then Exit Sub

        For Each group In _currentobject.memberOf
            group.Entry.Invoke("Remove", _currentobject.Entry.Path)
            group.Entry.CommitChanges()
        Next
        _currentobject.memberOf.Clear()
        For Each group In _currentobject.Domain.DefaultGroups
            group.Entry.Invoke("Add", _currentobject.Entry.Path)
            group.Entry.CommitChanges()
            _currentobject.memberOf.Add(group)
        Next

        lvSelectedGroups.ItemsSource = Nothing 'ultra dirty hack, without it LV doesn't updating
        lvSelectedGroups.ItemsSource = _currentobject.memberOf
    End Sub
End Class
