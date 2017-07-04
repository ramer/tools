Public Class clsPowerShellLog
    Private _timestamp As Date
    Private _command As String
    Private _result As String
    Private _error As String

    Sub New()
        _timestamp = Now
    End Sub

    Sub New(Command As String,
            Optional Result As String = "",
            Optional Err As String = "")

        _timestamp = Now
        _command = Command
        _result = Result
        _error = Err
    End Sub

    Public ReadOnly Property Image() As String
        Get
            If Err = "" Or Err Is Nothing Then
                Return "img/ready.ico"
            Else
                Return "img/warning.ico"
            End If
        End Get
    End Property

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

    Public Property Command() As String
        Get
            Return _command
        End Get
        Set(ByVal value As String)
            _command = value
        End Set
    End Property

    Public Property Result() As String
        Get
            Return _result
        End Get
        Set(ByVal value As String)
            _result = value
        End Set
    End Property

    Public Property Err() As String
        Get
            Return _error
        End Get
        Set(ByVal value As String)
            _error = value
        End Set
    End Property
End Class
