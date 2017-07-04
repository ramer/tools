Imports System.Reflection

Public Class wndAbout

    Public ReadOnly Property Version As String
        Get
            Return Assembly.GetExecutingAssembly().GetName().Version.ToString()
        End Get
    End Property

    Public ReadOnly Property Copyright As String
        Get
            Return FileVersionInfo.GetVersionInfo(Me.GetType.Assembly.Location).LegalCopyright
        End Get
    End Property

    Public ReadOnly Property Company As String
        Get
            Return FileVersionInfo.GetVersionInfo(Me.GetType.Assembly.Location).CompanyName
        End Get
    End Property

End Class
