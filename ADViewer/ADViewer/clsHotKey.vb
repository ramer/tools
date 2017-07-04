Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Public Class clsHotKey
    Implements IDisposable
    Private Shared _dictHotKeyToCalBackProc As Dictionary(Of Integer, clsHotKey)

    <DllImport("user32.dll")>
    Private Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As UInt32, vlc As UInt32) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    Public Const WmHotKey As Integer = &H312

    Private _disposed As Boolean = False

    Public Property Key() As Key
        Get
            Return m_Key
        End Get
        Private Set
            m_Key = Value
        End Set
    End Property
    Private m_Key As Key
    Public Property KeyModifiers() As KeyModifier
        Get
            Return m_KeyModifiers
        End Get
        Private Set
            m_KeyModifiers = Value
        End Set
    End Property
    Private m_KeyModifiers As KeyModifier
    Public Property Action() As Action(Of clsHotKey)
        Get
            Return m_Action
        End Get
        Private Set
            m_Action = Value
        End Set
    End Property
    Private m_Action As Action(Of clsHotKey)
    Public Property Id() As Integer
        Get
            Return m_Id
        End Get
        Set
            m_Id = Value
        End Set
    End Property
    Private m_Id As Integer

    ' ******************************************************************
    Public Sub New(k As Key, keyModifiers__1 As KeyModifier, action__2 As Action(Of clsHotKey), Optional register__3 As Boolean = True)
        Key = k
        KeyModifiers = keyModifiers__1
        Action = action__2
        If register__3 Then
            Register()
        End If
    End Sub

    ' ******************************************************************
    Public Function Register() As Boolean
        Dim virtualKeyCode As Integer = KeyInterop.VirtualKeyFromKey(Key)
        Id = virtualKeyCode + (KeyModifiers * &H10000)
        Dim result As Boolean = RegisterHotKey(IntPtr.Zero, Id, CType(KeyModifiers, UInteger), CType(virtualKeyCode, UInteger))

        If _dictHotKeyToCalBackProc Is Nothing Then
            _dictHotKeyToCalBackProc = New Dictionary(Of Integer, clsHotKey)()
            AddHandler ComponentDispatcher.ThreadFilterMessage, AddressOf ComponentDispatcherThreadFilterMessage
        End If

        _dictHotKeyToCalBackProc.Add(Id, Me)

        'Debug.Print(result.ToString() + ", " + Id + ", " + virtualKeyCode)
        Return result
    End Function

    ' ******************************************************************
    Public Sub Unregister()
        Dim hotKey As clsHotKey = Nothing
        If _dictHotKeyToCalBackProc.TryGetValue(Id, hotKey) Then
            UnregisterHotKey(IntPtr.Zero, Id)
        End If
    End Sub

    ' ******************************************************************
    Private Shared Sub ComponentDispatcherThreadFilterMessage(ByRef msg As MSG, ByRef handled As Boolean)
        If Not handled Then
            If msg.message = WmHotKey Then
                Dim hotKey As clsHotKey = Nothing

                If _dictHotKeyToCalBackProc.TryGetValue(CInt(msg.wParam), hotKey) Then
                    If hotKey.Action IsNot Nothing Then
                        hotKey.Action.Invoke(hotKey)
                    End If
                    handled = True
                End If
            End If
        End If
    End Sub

    ' ******************************************************************
    ' Implement IDisposable.
    ' Do not make this method virtual.
    ' A derived class should not be able to override this method.
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        ' This object will be cleaned up by the Dispose method.
        ' Therefore, you should call GC.SupressFinalize to
        ' take this object off the finalization queue
        ' and prevent finalization code for this object
        ' from executing a second time.
        GC.SuppressFinalize(Me)
    End Sub

    ' ******************************************************************
    ' Dispose(bool disposing) executes in two distinct scenarios.
    ' If disposing equals true, the method has been called directly
    ' or indirectly by a user's code. Managed and unmanaged resources
    ' can be _disposed.
    ' If disposing equals false, the method has been called by the
    ' runtime from inside the finalizer and you should not reference
    ' other objects. Only unmanaged resources can be _disposed.
    Protected Overridable Sub Dispose(disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not Me._disposed Then
            ' If disposing equals true, dispose all managed
            ' and unmanaged resources.
            If disposing Then
                ' Dispose managed resources.
                Unregister()
            End If

            ' Note disposing has been done.
            _disposed = True
        End If
    End Sub
End Class

' ******************************************************************
<Flags>
    Public Enum KeyModifier
        None = &H0
        Alt = &H1
        Ctrl = &H2
        NoRepeat = &H4000
        Shift = &H4
        Win = &H8
    End Enum

