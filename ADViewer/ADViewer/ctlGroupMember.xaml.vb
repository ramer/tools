
Imports System.Collections.ObjectModel

Public Class ctlGroupMember


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlGroupMember),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

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

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMember = CType(d, ctlGroupMember)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomainobjects.Clear()
            .lvSelectedObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentobject.member, Nothing)
            .lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)
        End With
    End Sub

    Private Async Sub tbDomainObjectsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainObjectsFilter.KeyDown
        If e.Key = Key.Enter Then
            Await searcher.BasicSearchAsync(_currentdomainobjects, tbDomainObjectsFilter.Text, _currentobject.Domain,, True, True, True, False)
        End If
    End Sub

    Private Sub lvDomainObjects_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvDomainObjects.MouseDoubleClick
        If lvDomainObjects.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvDomainObjects.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSelectedObjects_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseDoubleClick
        If lvSelectedObjects.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSelectedObjects.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvSelectedObjects.PreviewMouseLeftButtonDown,
                                                                                                   lvDomainObjects.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        allowdrag = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth 'And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseMove,
                                                                            lvDomainObjects.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And listView.SelectedItem IsNot Nothing And allowdrag Then
            sourceobject = listView

            Dim obj As clsDirectoryObject = CType(listView.SelectedItem, clsDirectoryObject)
            Dim dragData As New DataObject("clsDirectoryObject", obj)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Move)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.DragEnter,
                                                                            lvDomainObjects.DragEnter

        If Not e.Data.GetDataPresent("clsDirectoryObject") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.Drop,
                                                                        lvDomainObjects.Drop

        If e.Data.GetDataPresent("clsDirectoryObject") And sender IsNot sourceobject Then
            Dim draggedObject As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            If sender Is lvSelectedObjects Then ' adding member
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
            For Each obj As clsDirectoryObject In _currentobject.member
                If obj.name = [object].name Then Exit Sub
            Next
            _currentobject.Entry.Invoke("Add", [object].Entry.Path)
            _currentobject.Entry.CommitChanges()
            _currentobject.member.Add([object])
        Catch ex As Exception
            ThrowException(ex, "AddMember")
            If [object].SchemaClassName = "group" Then ShowWrongMemberMessage()
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try
            _currentobject.Entry.Invoke("Remove", [object].Entry.Path)
            _currentobject.Entry.CommitChanges()
            _currentobject.member.Remove([object])

        Catch ex As Exception
            ThrowException(ex, "RemoveMember")
        End Try
    End Sub

    Private Sub ctlGroupMember_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainObjectsFilter.Focus()
    End Sub
End Class
