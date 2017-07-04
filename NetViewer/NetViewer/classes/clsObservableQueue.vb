Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class clsObservableQueue(Of T)
    Implements IList(Of T)
    Implements INotifyPropertyChanged
    Implements INotifyCollectionChanged

    Private collection As IList(Of T) = New List(Of T)()

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Event CollectionChanged As NotifyCollectionChangedEventHandler Implements INotifyCollectionChanged.CollectionChanged

    Public Sub New()

    End Sub

    Public Sub New(collection As IList(Of T))
        Me.collection = collection
    End Sub

    Public Sub Add(item As T) Implements IList(Of T).Add
        collection.Add(item)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item))
    End Sub

    Public Sub Enqueue(item As T)
        collection.Add(item)
        If collection.Count >= 60 Then collection.RemoveAt(0)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub

    Public Sub Clear() Implements IList(Of T).Clear
        collection.Clear()
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub

    Public Function Contains(item As T) As Boolean Implements IList(Of T).Contains
        Return collection.Contains(item)
    End Function

    Public Sub CopyTo(array As T(), arrayIndex As Integer) Implements IList(Of T).CopyTo
        collection.CopyTo(array, arrayIndex)
    End Sub

    Public ReadOnly Property Count() As Integer Implements IList(Of T).Count
        Get
            Return collection.Count
        End Get
    End Property

    Public ReadOnly Property IsReadOnly() As Boolean Implements IList(Of T).IsReadOnly
        Get
            Return collection.IsReadOnly
        End Get
    End Property

    Public Function Remove(item As T) As Boolean Implements IList(Of T).Remove
        Dim index = collection.IndexOf(item)
        If index = -1 Then
            Return False
        End If
        Dim result = collection.Remove(item)
        If result Then
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End If
        Return result
    End Function

    Public Function GetEnumerator() As IEnumerator(Of T) Implements IList(Of T).GetEnumerator
        Return collection.GetEnumerator()
    End Function

    Private Function System_Collections_IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return collection.GetEnumerator()
    End Function

    Public Function IndexOf(item As T) As Integer Implements IList(Of T).IndexOf
        Return collection.IndexOf(item)
    End Function

    Public Sub Insert(index As Integer, item As T) Implements IList(Of T).Insert
        collection.Insert(index, item)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item, index))
    End Sub

    Public Sub RemoveAt(index As Integer) Implements IList(Of T).RemoveAt
        collection.RemoveAt(index)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub

    Default Public Property Item(index As Integer) As T Implements IList(Of T).Item
        Get
            Return collection(index)
        End Get
        Set
            If collection.Count = 0 OrElse collection.Count <= index Then
                Return
            End If
            collection(index) = Value
        End Set
    End Property

End Class
