Public Class clsErrorLog
    Private _timestamp As Date
    Private _command As String
    Private _object As String
    Private _error As Exception

    Sub New()
        _timestamp = Now
    End Sub

    Sub New(Command As String,
        Optional Obj As String = "",
        Optional Err As Exception = Nothing)

        _timestamp = Now
        _command = Command
        _object = Obj
        _error = Err
    End Sub

    Public ReadOnly Property Image() As String
        Get
            Return "img/warning.ico"
        End Get
    End Property

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

    Public ReadOnly Property Command() As String
        Get
            Return _command
        End Get
    End Property

    Public ReadOnly Property Obj() As String
        Get
            Return _object
        End Get
    End Property

    Public ReadOnly Property Err() As String
        Get
            Return _error.Message & vbCrLf & _error.StackTrace
        End Get
    End Property
End Class
