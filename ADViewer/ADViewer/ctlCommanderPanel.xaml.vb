Imports System.Collections.ObjectModel
Imports System.DirectoryServices

Public Class ctlCommanderPanel

    Public Shared ReadOnly ContainerProperty As DependencyProperty = DependencyProperty.Register("Container",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlCommanderPanel),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf ContainerPropertyChanged))

    Public Shared ReadOnly DomainProperty As DependencyProperty = DependencyProperty.Register("Domain",
                                                            GetType(clsDomain),
                                                            GetType(ctlCommanderPanel),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf DomainPropertyChanged))

    Private Property _domain As clsDomain
    Private Property _container As clsDirectoryObject
    Private Property _objects As New ObservableCollection(Of clsDirectoryObject)
    Private Property _lastselected As Object


    WithEvents searcher As New clsSearcher

    Private sourceobject As Object
    Private allowdrag As Boolean

    Public Property Container() As clsDirectoryObject
        Get
            Return GetValue(ContainerProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(ContainerProperty, value)
        End Set
    End Property

    Public Property Domain() As clsDomain
        Get
            Return GetValue(DomainProperty)
        End Get
        Set(ByVal value As clsDomain)
            SetValue(DomainProperty, value)
        End Set
    End Property

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub ContainerPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlCommanderPanel = CType(d, ctlCommanderPanel)
        With instance
            ._container = CType(e.NewValue, clsDirectoryObject)
            '._currentdomainobjects.Clear()
            '.lvSelectedObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentobject.member, Nothing)
            '.lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)
        End With
    End Sub

    Private Shared Sub DomainPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlCommanderPanel = CType(d, ctlCommanderPanel)
        With instance
            ._domain = CType(e.NewValue, clsDomain)
        End With
    End Sub

#Region "Events"

    Private Sub ctlCommanderPanel_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        cmboDomains.ItemsSource = domains
        dgPanel.ItemsSource = _objects
    End Sub

    Private Sub cmboDomains_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboDomains.SelectionChanged
        Dim cmbo As ComboBox = CType(sender, ComboBox)
        If cmbo.SelectedItem Is Nothing Then Exit Sub
        Dim root As New clsDirectoryObject(CType(cmbo.SelectedItem, clsDomain).DefaultNamingContext, CType(cmbo.SelectedItem, clsDomain))
        OpenNode(root)
        dgPanel.UpdateLayout()
        If dgPanel.Items.Count > 0 Then
            Dim row As DataGridRow = dgPanel.ItemContainerGenerator.ContainerFromItem(dgPanel.Items(0))
            row.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
        End If
    End Sub

    Private Sub dgPanel_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles dgPanel.PreviewKeyDown
        Select Case e.Key
            Case Key.Enter
                If dgPanel.SelectedItem Is Nothing Then Exit Sub
                Dim current As clsDirectoryObject = CType(dgPanel.SelectedItem, clsDirectoryObject)
                OpenNode(current)
                e.Handled = True
            Case Key.Back
                OpenParentNode()
                e.Handled = True
            Case Key.Home
                If dgPanel.Items.Count > 0 Then dgPanel.SelectedIndex = 0 : dgPanel.CurrentItem = dgPanel.SelectedItem
                e.Handled = True
            Case Key.End
                If dgPanel.Items.Count > 0 Then dgPanel.SelectedIndex = dgPanel.Items.Count - 1 : dgPanel.CurrentItem = dgPanel.SelectedItem
                e.Handled = True
                'Case Key.Tab
                '    If dg Is dgLPanel Then
                '        dgRPanel.Focus()
                '    ElseIf dg Is dgRPanel Then
                '        dgLPanel.Focus()
                '    End If
                '    e.Handled = True
        End Select
    End Sub

    Private Sub dgPanel_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgPanel.MouseDoubleClick
        If dgPanel.SelectedItem Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = CType(dgPanel.SelectedItem, clsDirectoryObject)
        OpenNode(current)
    End Sub

    Private Sub btnBack_Click(sender As Object, e As RoutedEventArgs) Handles btnBack.Click
        OpenParentNode()
    End Sub

#End Region

#Region "Subs"

    'Public Sub HideSelection()
    '    _lastselected = dgPanel.SelectedItems
    'End Sub

    'Public Sub ShowSelection()
    '    dgPanel.se
    'End Sub

    Public Sub OpenNode(current As clsDirectoryObject)
        If current.SchemaClassName = "computer" Or
           current.SchemaClassName = "group" Or
           current.SchemaClassName = "contact" Or
           current.SchemaClassName = "user" Then

            'cmdEdit(current)
        ElseIf current.SchemaClassName = "organizationalUnit" Or
               current.SchemaClassName = "container" Or
               current.SchemaClassName = "domainDNS" Then

            GetContainerChildren(current)
            dgPanel.Tag = current
        End If

        ShowPath()
    End Sub

    Public Sub OpenParentNode()
        Dim current As clsDirectoryObject = CType(dgPanel.Tag, clsDirectoryObject)
        If current Is Nothing OrElse (current.Entry.Parent Is Nothing OrElse current.Entry.Path = current.Domain.DefaultNamingContext.Path) Then Exit Sub
        Dim parent As New clsDirectoryObject(current.Entry.Parent, current.Domain)
        GetContainerChildren(parent)
        dgPanel.Tag = parent

        ShowPath()
    End Sub

    Public Sub ShowObject(obj As clsDirectoryObject)
        If obj Is Nothing Then Exit Sub
        dgPanel.Tag = obj
        OpenParentNode()

        For Each item As clsDirectoryObject In dgPanel.Items
            If item.Entry.Path = obj.Entry.Path Then
                dgPanel.SelectedItem = item
                dgPanel.CurrentItem = dgPanel.SelectedItem
                Exit Sub
            End If
        Next
    End Sub

    Private Sub GetContainerChildren(parent As clsDirectoryObject)
        Dim ds As New DirectorySearcher(parent.Entry)
        ds.PropertiesToLoad.AddRange({"name", "objectClass"})
        'ds.Filter = "(|(objectClass=organizationalUnit)(objectClass=container))"
        ds.SearchScope = SearchScope.OneLevel

        _objects.Clear()
        For Each sr As SearchResult In ds.FindAll()
            _objects.Add(New clsDirectoryObject(sr.GetDirectoryEntry(), parent.Domain))
        Next

        dgPanel.Focus()
        dgPanel.UpdateLayout()
        If dgPanel.Items.Count > 0 Then
            dgPanel.SelectedIndex = 0
            dgPanel.CurrentItem = dgPanel.SelectedItem
        End If
    End Sub

    Private Sub ShowPath()
        Dim current As clsDirectoryObject = CType(dgPanel.Tag, clsDirectoryObject)
        If current Is Nothing Then Exit Sub

        Dim children As New List(Of Button)

        Do
            Dim btn As New Button
            btn.Background = Brushes.Transparent
            btn.Height = 23
            btn.Content = If(current.SchemaClassName = "domainDNS", current.Domain.Name, current.name)
            btn.Margin = New Thickness(2, 0, 2, 0)
            btn.Padding = New Thickness(5, 0, 5, 0)
            btn.Tag = current
            children.Add(btn)

            If current.Entry.Parent Is Nothing OrElse current.Entry.Path = current.Domain.DefaultNamingContext.Path Then
                Exit Do
            Else
                current = New clsDirectoryObject(current.Entry.Parent, current.Domain)
            End If
        Loop

        children.Reverse()

        For Each child As Button In wpPath.Children
            RemoveHandler child.Click, AddressOf btnPath_Click
        Next
        wpPath.Children.Clear()
        For Each child As Button In children
            AddHandler child.Click, AddressOf btnPath_Click
            wpPath.Children.Add(child)
        Next
    End Sub

    Private Sub btnPath_Click(sender As Object, e As RoutedEventArgs)
        If sender.Tag Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = CType(sender.Tag, clsDirectoryObject)
        OpenNode(current)
    End Sub




#End Region

End Class
