Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices

Public Class wndCommander
    Public Property lobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property robjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Shared hkAltF1 As New RoutedCommand
    Public Shared hkAltF2 As New RoutedCommand
    Public Shared hkF4 As New RoutedCommand
    Public Shared hkF5 As New RoutedCommand
    Public Shared hkF6 As New RoutedCommand
    Public Shared hkF7 As New RoutedCommand
    Public Shared hkF8 As New RoutedCommand
    Public Shared hkF9 As New RoutedCommand

    Private activepanel As ctlCommanderPanel
    Private inactivepanel As ctlCommanderPanel
    Private activelitem As Integer = 0
    Private activeritem As Integer = 0

    Private lcontainer As clsDirectoryObject
    Private rcontainer As clsDirectoryObject

#Region "Events"

    Private Sub wndCommander_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        hkAltF1.InputGestures.Add(New KeyGesture(Key.F1, ModifierKeys.Alt))
        Me.CommandBindings.Add(New CommandBinding(hkAltF1, AddressOf HotKey_AltF1))
        hkAltF2.InputGestures.Add(New KeyGesture(Key.F2, ModifierKeys.Alt))
        Me.CommandBindings.Add(New CommandBinding(hkAltF2, AddressOf HotKey_AltF2))
        hkF4.InputGestures.Add(New KeyGesture(Key.F4))
        Me.CommandBindings.Add(New CommandBinding(hkF4, AddressOf HotKey_F4))
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf HotKey_F5))
        hkF6.InputGestures.Add(New KeyGesture(Key.F6))
        Me.CommandBindings.Add(New CommandBinding(hkF6, AddressOf HotKey_F6))
        hkF7.InputGestures.Add(New KeyGesture(Key.F7))
        Me.CommandBindings.Add(New CommandBinding(hkF7, AddressOf HotKey_F7))
        hkF8.InputGestures.Add(New KeyGesture(Key.F8))
        Me.CommandBindings.Add(New CommandBinding(hkF8, AddressOf HotKey_F8))
        hkF9.InputGestures.Add(New KeyGesture(Key.F9))
        Me.CommandBindings.Add(New CommandBinding(hkF9, AddressOf HotKey_F9))
    End Sub

    Private Sub wndCommander_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub pnlLeft_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles pnlLeft.PreviewKeyDown
        Select Case e.Key
            Case Key.Tab
                pnlRight.dgPanel.Focus()
                If pnlRight.dgPanel.SelectedItem Is Nothing Then Exit Sub
                Dim row As DataGridRow = pnlRight.dgPanel.ItemContainerGenerator.ContainerFromItem(pnlRight.dgPanel.SelectedItem)
                row.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
                e.Handled = True
        End Select
    End Sub

    Private Sub pnlRight_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles pnlRight.PreviewKeyDown
        Select Case e.Key
            Case Key.Tab
                pnlLeft.dgPanel.Focus()
                If pnlLeft.dgPanel.SelectedItem Is Nothing Then Exit Sub
                Dim row As DataGridRow = pnlLeft.dgPanel.ItemContainerGenerator.ContainerFromItem(pnlLeft.dgPanel.SelectedItem)
                row.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
                e.Handled = True
        End Select
    End Sub

    Private Sub pnlLeft_GotFocus(sender As Object, e As RoutedEventArgs) Handles pnlLeft.GotFocus
        activepanel = pnlLeft
        inactivepanel = pnlRight
    End Sub

    Private Sub pnlRight_GotFocus(sender As Object, e As RoutedEventArgs) Handles pnlRight.GotFocus
        activepanel = pnlRight
        inactivepanel = pnlLeft
    End Sub

    Private Sub btnCommandEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandEdit.Click
        cmdEdit()
    End Sub

    Private Sub btnCommandCopy_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandCopy.Click
        cmdCopy()
    End Sub

    Private Sub btnCommandMove_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandMove.Click
        cmdMove()
    End Sub

    Private Sub btnCommandCreateContainer_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandCreateContainer.Click
        cmdCreateContainer()
    End Sub

    Private Sub btnCommandRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandRemove.Click
        cmdRemove()
    End Sub

    Private Sub btnCommandCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles btnCommandCreateObject.Click
        cmdCreateObject()
    End Sub

#End Region

#Region "Subs"

#End Region

#Region "Hotkeys"

    Private Sub HotKey_AltF1()
        pnlLeft.cmboDomains.IsDropDownOpen = True
        pnlLeft.cmboDomains.Focus()
    End Sub

    Private Sub HotKey_AltF2()
        pnlRight.cmboDomains.IsDropDownOpen = True
        pnlRight.cmboDomains.Focus()
    End Sub

    Private Sub HotKey_F4()
        cmdEdit()
    End Sub

    Private Sub HotKey_F5()
        cmdCopy()
    End Sub

    Private Sub HotKey_F6()
        cmdMove()
    End Sub

    Private Sub HotKey_F7()
        cmdCreateContainer()
    End Sub

    Private Sub HotKey_F8()
        cmdRemove()
    End Sub

    Private Sub HotKey_F9()
        cmdCreateObject()
    End Sub

#End Region

#Region "Commands"

    Private Sub cmdEdit()
        If activepanel Is Nothing OrElse activepanel.dgPanel.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(activepanel.dgPanel.SelectedItem, Me)
    End Sub

    Private Sub cmdCopy()
        If activepanel Is Nothing OrElse activepanel.dgPanel.SelectedItems Is Nothing Then Exit Sub
        If inactivepanel Is Nothing OrElse inactivepanel.dgPanel.Tag Is Nothing Then Exit Sub

        Dim w As New wndCommanderCopy With {.Owner = Me}

        Dim sourceobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
        For Each obj As clsDirectoryObject In activepanel.dgPanel.SelectedItems
            sourceobjects.Add(obj)
        Next
        w.sourceobjects = sourceobjects
        w.destination = inactivepanel.dgPanel.Tag
        w.ShowDialog()

        inactivepanel.OpenNode(inactivepanel.dgPanel.Tag)
    End Sub

    Private Sub cmdMove()
        If activepanel Is Nothing OrElse activepanel.dgPanel.SelectedItems Is Nothing Then Exit Sub
        If inactivepanel Is Nothing OrElse inactivepanel.dgPanel.Tag Is Nothing Then Exit Sub

        Dim sourceobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
        Dim destination As clsDirectoryObject = CType(inactivepanel.dgPanel.Tag, clsDirectoryObject)
        For Each obj As clsDirectoryObject In activepanel.dgPanel.SelectedItems
            sourceobjects.Add(obj)
        Next

        If sourceobjects.Count > 0 AndAlso sourceobjects(0).Domain IsNot destination.Domain Then IMsgBox("Перемещение в другие домены запрещено", "Перемещение объектов", vbOKOnly, vbExclamation) : Exit Sub
        If IMsgBox("Вы уверены?", "Перемещение объектов", vbYesNo, vbQuestion) <> MessageBoxResult.Yes Then Exit Sub

        Try
            For Each obj In sourceobjects
                obj.Entry.MoveTo(destination.Entry)
                obj.NotifyMoved()
            Next
        Catch ex As Exception
            ThrowException(ex, "cmdMove")
        End Try

        activepanel.OpenNode(activepanel.dgPanel.Tag)
        inactivepanel.OpenNode(inactivepanel.dgPanel.Tag)
    End Sub

    Private Sub cmdCreateContainer()
        If activepanel Is Nothing OrElse activepanel.dgPanel.Tag Is Nothing Then Exit Sub
        Dim obj As clsDirectoryObject = CType(activepanel.dgPanel.Tag, clsDirectoryObject)
        If obj Is Nothing Then Exit Sub
        Dim ou As String = IInputBox("Введите название контейнера", "Создание контейнера", vbQuestion)
        If String.IsNullOrEmpty(ou) Then Exit Sub
        ou = "OU=" & ou

        Dim newou As DirectoryEntry = Nothing
        Try
            newou = obj.Entry.Children.Add(ou, "organizationalUnit")
            newou.CommitChanges()
        Catch ex As Exception
            ThrowException(ex, "cmdCreateContainer")
        End Try

        If newou IsNot Nothing Then activepanel.ShowObject(New clsDirectoryObject(newou, obj.Domain))
    End Sub

    Private Sub cmdRemove()
        If activepanel Is Nothing OrElse activepanel.dgPanel.SelectedItem Is Nothing Then Exit Sub
        Dim obj As clsDirectoryObject = CType(activepanel.dgPanel.SelectedItem, clsDirectoryObject)
        If obj Is Nothing Then Exit Sub
        If IMsgBox("Вы уверены?", "Удаление объекта", vbYesNo, vbQuestion) <> MsgBoxResult.Yes Then Exit Sub
        If obj.SchemaClassName = "organizationalUnit" AndAlso IMsgBox("Это контейнер!" & vbCrLf & vbCrLf & "Вы уверены?", "Удаление объекта", vbYesNo, vbExclamation) <> MsgBoxResult.Yes Then Exit Sub

        Try
            obj.Entry.DeleteTree()
        Catch ex As Exception
            ThrowException(ex, "cmdRemove")
        End Try

        activepanel.OpenNode(activepanel.dgPanel.Tag)
    End Sub

    Private Sub cmdCreateObject()
        If activepanel Is Nothing OrElse activepanel.dgPanel.Tag Is Nothing Then Exit Sub
        Dim obj As clsDirectoryObject = CType(activepanel.dgPanel.Tag, clsDirectoryObject)

        Dim w As New wndCreateObject With {.Owner = Me}
        w.cmboUserDomain.SelectedItem = obj.Domain
        w.tbUserContainer.Tag = obj
        w.tbUserContainer.Text = obj.Entry.Path
        w.cmboComputerDomain.SelectedItem = obj.Domain
        w.tbComputerContainer.Tag = obj
        w.tbComputerContainer.Text = obj.Entry.Path
        w.cmboGroupDomain.SelectedItem = obj.Domain
        w.tbGroupContainer.Tag = obj
        w.tbGroupContainer.Text = obj.Entry.Path
        w.cmboContactDomain.SelectedItem = obj.Domain
        w.tbContactContainer.Tag = obj
        w.tbContactContainer.Text = obj.Entry.Path
        w.ShowDialog()

        activepanel.OpenNode(activepanel.dgPanel.Tag)
    End Sub



#End Region

End Class
