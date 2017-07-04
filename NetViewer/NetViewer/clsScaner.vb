Imports System.Collections.Concurrent
Imports System.Collections.ObjectModel
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Threading

Public Class clsScaner

    Private basicscantasks As New List(Of Task)
    Private basicscantaskscts As New CancellationTokenSource

    Private resultqueue As New ConcurrentQueue(Of clsPingResult)

    Sub New()

    End Sub

    Public Async Function BasicScanAsync(returncollection As clsThreadSafeObservableCollection(Of clsHostTiny), range As String) As Task
        basicscantaskscts.Cancel()
        If basicscantasks.Count > 0 Then Exit Function
        returncollection.Clear()
        basicscantaskscts = New CancellationTokenSource

        Dim hostlist As List(Of IPAddress) = GetHostListFromRange(range)
        Dim hostlists() As List(Of IPAddress) = SplitList(hostlist, 5)

        For Each hl In hostlists

            Dim mt = Task.Run(Sub() BasicScanSync(hl, basicscantaskscts.Token))
            basicscantasks.Add(mt)

        Next

        Dim lt = Task.Run(
                Sub()
                    Do
                        While resultqueue.Count > 0
                            If resultqueue.Count > 0 Then
                                Dim s As clsPingResult = Nothing
                                resultqueue.TryDequeue(s)
                                returncollection.Add(New clsHostTiny(New clsPingResult(s.Status, s.RoundtripTime, s.Address)))
                            End If
                            Thread.Sleep(100)
                        End While
                        If basicscantaskscts.Token.IsCancellationRequested Then Exit Do
                        Thread.Sleep(100)
                    Loop
                End Sub)

        Await Task.WhenAll(basicscantasks.ToArray)
        basicscantasks.Clear()
        basicscantaskscts.Cancel()
        Await lt
    End Function

    Public Sub BasicScanSync(hostlist As List(Of IPAddress), Optional ct As CancellationToken = Nothing)
        Try
            For Each h In hostlist
                If ct.IsCancellationRequested Then Exit Sub

                Dim status As clsPingResult
                status = GetPingReply(h.ToString)
                If status IsNot Nothing Then
                    If status.Status = IPStatus.Success Then resultqueue.Enqueue(status)
                End If
            Next

        Catch ex As Exception
            ThrowException(ex, "BasicScanSync")
        End Try
    End Sub

End Class
