Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Data
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports ServiceStack.OrmLite

Module mdlTools

    Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIP As UInt32, ByVal SrcIP As UInt32, ByVal pMacAddr As Byte(), ByRef PhyAddrLen As Integer) As Integer

    Public preferences As clsPreferences

    Public dbpath As String = IO.Path.Combine(IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NetViewer"), ("hosts.nvh"))
    Public dbFactory As New OrmLiteConnectionFactory(dbpath, SqliteDialect.Provider)
    Public db As IDbConnection

    Public regDomains As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\NetViewerWPF\Domains")
    Public regInterface As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\NetViewerWPF\Interface")
    Public regSettings As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\NetViewerWPF\Settings")
    Public regExternalSoftware As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\NetViewerWPF\ExternalSoftware")

    Public AppInitialized As Boolean = False

    Public hosttypes As New Dictionary(Of String, String) From {
        {"ap", "Wi-Fi точка доступа"},
        {"computer", "Компьютер"},
        {"mfu", "МФУ"},
        {"phone", "IP телефон"},
        {"printer", "Принтер"},
        {"router", "Маршрутизатор"},
        {"server", "Сервер"},
        {"switch", "Коммутатор"}}

    Public pingpriorities As New Dictionary(Of String, String) From {
        {"ipaddress", "IP адрес"},
        {"hostname", "Имя хоста"},
        {"port", "Порт (не работает пока)"}}

    Public WithEvents hosts As New ObservableCollection(Of clsHost)

    Public Sub AppInitialize()
        db = dbFactory.Open()
        db.CreateTableIfNotExists(Of clsHost)
        db.CreateTableIfNotExists(Of clsHistoryItem)

        For Each h In db.Select(Of clsHost)
            hosts.Add(h)
        Next

        AppInitialized = True
    End Sub

    Private Sub hosts_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles hosts.CollectionChanged
        If AppInitialized = False Then Exit Sub

        If e.Action = NotifyCollectionChangedAction.Add Then
            For Each h As clsHost In e.NewItems
                db.Save(h)
            Next
        ElseIf e.Action = NotifyCollectionChangedAction.Remove Then
            For Each h As clsHost In e.OldItems
                h.HistoryClear()
                h.StopPinger()
                db.Delete(h)
            Next
        End If
    End Sub

    Public Function IMsgBox(Promt As String, Optional Title As String = "", Optional Buttons As MsgBoxStyle = vbOKOnly, Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly) As MessageBoxResult
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._content = Promt
        w._title = Title
        w._buttons = Buttons
        w._icon = Icon
        If w.ShowDialog() Then
            Return w._msgboxresult
        Else
            Return vbCancel
        End If
    End Function

    Public Function IInputBox(Promt As String, Optional Title As String = "", Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional DefaultAnswer As String = "") As String
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._inputbox = True
        w._content = Promt
        w._title = Title
        w._buttons = vbOKCancel
        w._icon = Icon
        w._defaultanswer = DefaultAnswer
        If w.ShowDialog() Then
            Return w.tbInput.Text
        Else
            Return Nothing
        End If
    End Function

    Public Function IPasswordBox(Promt As String, Optional Title As String = "", Optional Icon As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional DefaultAnswer As String = "") As String
        Dim w As New wndQuestion() With {.Owner = Application.Current.Windows.OfType(Of Window)().SingleOrDefault(Function(o) o.IsActive)}
        w._passwordbox = True
        w._content = Promt
        w._title = Title
        w._buttons = vbOKCancel
        w._icon = Icon
        If w.ShowDialog() Then
            Return w.tbInput.Text
        Else
            Return Nothing
        End If
    End Function

    Public Sub ThrowException(ByVal ex As Exception, ByVal Procedure As String)
        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog(Procedure,, ex))
    End Sub

    Public Sub ThrowCustomException(Message As String)
        SingleInstanceApplication.tsocErrorLog.Add(New clsErrorLog(Message))
    End Sub

    Public Sub ThrowInformation(Message As String)
        With SingleInstanceApplication.nicon
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = ProductName()
            .BalloonTipText = Message
            .Tag = Nothing
            .Visible = False
            .Visible = True
            .ShowBalloonTip(5000)
        End With
    End Sub

    Public Function ProductName() As String
        If Windows.Application.ResourceAssembly Is Nothing Then
            Return Nothing
        End If

        Return Windows.Application.ResourceAssembly.GetName().Name
    End Function

    Public Function GetHostnameFromIPAddress(IPAddress As String) As String
        Dim hostentry As System.Net.IPHostEntry
        Try
            Dim tsk = System.Net.Dns.GetHostEntryAsync(IPAddress)
            tsk.Wait()
            hostentry = tsk.Result
        Catch ex As Exception
            Return ex.InnerException.Message
        End Try

        If hostentry IsNot Nothing AndAlso hostentry.AddressList.Count > 0 Then Return hostentry.HostName
        Return Nothing
    End Function

    Public Function GetIPAddressFromHostname(Hostname As String) As String
        Dim hostentry As System.Net.IPHostEntry = Nothing
        Try
            Dim tsk = System.Net.Dns.GetHostEntryAsync(Hostname)
            tsk.Wait()
            hostentry = tsk.Result
        Catch ex As Exception
            Return Nothing
        End Try

        If hostentry IsNot Nothing AndAlso hostentry.AddressList.Count > 0 Then Return hostentry.AddressList(0).ToString
        Return Nothing
    End Function

    Public Function GetMACAddressFromIPAddress(IPAddress As String) As String
        If IPAddress Is Nothing Then Return Nothing

        Try
            Dim mac() As Byte = New Byte(6) {}
            Dim len As Integer = mac.Length
            Dim tsk = Task.Run(Sub() SendARP(BitConverter.ToUInt32(Net.IPAddress.Parse(IPAddress).GetAddressBytes(), 0), 0, mac, len))
            tsk.Wait(100)

            Return Join(mac.Select(Function(x As Byte) x.ToString("X2")).ToArray, ":")
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Public Function GetHostListFromRange(ByVal range As String) As List(Of IPAddress)
        Dim result As New List(Of IPAddress)
        Dim parts() As String

        ' Parsing IPRange or  by ","
        parts = Split(range, ",")
        For I As Integer = 0 To parts.Length - 1
            parts(I) = Trim(parts(I))
        Next

        For Each part As String In parts
            If InStr(part, "/") Then ' subnet
                Try
                    Dim cidr() As String = part.Split("/"c)
                    If cidr.Length <> 2 Then Continue For

                    Dim ip As IPAddress = IPAddress.Parse(cidr(0))
                    Dim bits As Integer = CInt(cidr(1))
                    Dim mask As UInteger = Not (UInteger.MaxValue >> bits)
                    Dim ipBytes() As Byte = ip.GetAddressBytes()
                    Dim maskBytes() As Byte = BitConverter.GetBytes(mask).Reverse().ToArray()
                    Dim fromIPBytes(ipBytes.Length - 1) As Byte
                    Dim toIPBytes(ipBytes.Length - 1) As Byte
                    Dim tempIP As IPAddress = Nothing
                    Dim tempIPBytes() As Byte
                    Dim fromIPInt As Integer
                    Dim toIPInt As Integer
                    Dim tempIPInt As Integer

                    For i As Integer = 0 To ipBytes.Length - 1
                        fromIPBytes(i) = CByte(ipBytes(i) And maskBytes(i))
                        toIPBytes(i) = CByte(ipBytes(i) Or (Not maskBytes(i)))
                    Next i

                    fromIPInt = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(fromIPBytes, 0))
                    toIPInt = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(toIPBytes, 0))

                    For tempIPInt = fromIPInt + 1 To toIPInt - 1
                        tempIPBytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(tempIPInt))
                        tempIP = New IPAddress(tempIPBytes)
                        result.Add(tempIP)
                    Next

                Catch ex As Exception
                End Try
            ElseIf InStr(part, "-") Then 'IP range
                Try
                    Dim cidr() As String = part.Split("-"c)
                    If cidr.Length <> 2 Then Continue For

                    Dim ip1 As IPAddress = IPAddress.Parse(cidr(0))
                    Dim ip2 As IPAddress = IPAddress.Parse(cidr(1))
                    Dim tempIP As IPAddress = Nothing
                    Dim tempIPBytes() As Byte
                    Dim fromIPInt As Integer
                    Dim toIPInt As Integer
                    Dim tempIPInt As Integer

                    fromIPInt = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(ip1.GetAddressBytes(), 0))
                    toIPInt = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(ip2.GetAddressBytes(), 0))

                    For tempIPInt = fromIPInt To toIPInt
                        tempIPBytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(tempIPInt))
                        tempIP = New IPAddress(tempIPBytes)
                        result.Add(tempIP)
                    Next

                Catch ex As Exception
                End Try
            Else
                Try
                    result.Add(IPAddress.Parse(part))
                Catch ex As Exception
                End Try
            End If
        Next

        Return result
    End Function

    Public Function GetPingReply(HostnameOrIPAddress As String) As clsPingResult
        Dim result As New clsPingResult(IPStatus.BadDestination, 0, "")

        If HostnameOrIPAddress Is Nothing Then Return result

        If HostnameOrIPAddress.Contains(":") Then ' tcp ping
            Dim tcpClient = New TcpClient()
            Dim addrparts As String() = HostnameOrIPAddress.Split(":", 2, StringSplitOptions.RemoveEmptyEntries)
            If addrparts.Count <> 2 Then Return result
            Dim addr As String = addrparts(0)
            Dim port As Integer
            If Not Integer.TryParse(addrparts(1), port) Then Return result

            Dim timestart As Date = Now

            Dim connectionTask = tcpClient.ConnectAsync(addr, port).ContinueWith(Function(tsk) If(tsk.IsFaulted, Nothing, tcpClient))
            Dim timeoutTask = Task.Delay(1000).ContinueWith(Of TcpClient)(Function(tsk) Nothing)
            Dim resultTask = Task.WhenAny(connectionTask, timeoutTask).Unwrap()

            resultTask.Wait()

            Dim timeend As Date = Now

            Dim resultTcpClient = resultTask.Result

            If resultTcpClient IsNot Nothing Then
                resultTcpClient.Close()
                result.Status = IPStatus.Success
                result.RoundTripTime = (timeend - timestart).TotalMilliseconds
                result.Address = HostnameOrIPAddress
                Return result
            Else
                result.Status = IPStatus.TimedOut
                result.RoundTripTime = 0
                result.Address = ""
                Return result
            End If

        Else ' icmp ping
            Dim pingsender As New Ping
            Dim pingoptions As New PingOptions
            Dim pingtimeout As Integer = 1000
            Dim pingreplytask As Task(Of PingReply) = Nothing
            Dim pingreply As PingReply = Nothing
            Dim pingbuffer() As Byte = Encoding.ASCII.GetBytes(Space(32))

            pingoptions.Ttl = 128
            pingoptions.DontFragment = False
            pingreplytask = Task.Run(Function()
                                         Try
                                             Return pingsender.Send(HostnameOrIPAddress, pingtimeout, pingbuffer, pingoptions)
                                         Catch
                                             Return Nothing
                                         End Try
                                     End Function)
            pingreplytask.Wait()

            pingreply = pingreplytask.Result

            If pingreply IsNot Nothing Then
                result.Status = pingreply.Status
                result.RoundTripTime = pingreply.RoundtripTime
                result.Address = If(pingreply.Address IsNot Nothing, pingreply.Address.ToString, "")
                Return result
            Else
                result.Status = IPStatus.BadDestination
                result.RoundTripTime = 0
                result.Address = ""
                Return result
            End If

        End If
    End Function

    Public Function TraceRoute(IPAddress As IPAddress) As List(Of PingReply)
        If IPAddress Is Nothing Then Return New List(Of PingReply)

        Dim pingsender As New Ping
        Dim pingoptions As New PingOptions
        Dim pingtimeout As Integer = 1000
        Dim pingreplytask As Task(Of PingReply) = Nothing
        Dim pingreply As PingReply = Nothing
        Dim pingbuffer() As Byte = Encoding.ASCII.GetBytes(Space(32))

        Dim resultlist As New List(Of PingReply)

        For ttl As Integer = 1 To 128
            pingoptions.Ttl = ttl
            pingoptions.DontFragment = False

            pingreplytask = Task.Run(Function() pingsender.Send(IPAddress, pingtimeout, pingbuffer, pingoptions))
            pingreplytask.Wait()

            resultlist.Add(pingreplytask.Result)
            If pingreplytask.Result.Status = IPStatus.Success Then Exit For
        Next

        Return resultlist
    End Function

    Public Function ShowObjectProperties(obj As clsHost, Optional owner As Window = Nothing) As Window
        Dim w As wndHost
        If owner IsNot Nothing Then
            For Each wnd As Window In owner.OwnedWindows
                If GetType(wndHost) Is wnd.GetType AndAlso CType(wnd, wndHost).currentobject Is obj Then
                    w = wnd
                    w.Show() : w.Activate()
                    If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                    w.Topmost = True : w.Topmost = False
                    Return Nothing
                End If
            Next
        End If

        w = New wndHost
        If owner IsNot Nothing Then w.Owner = owner
        w.currentobject = obj
        w.Show()
        Return w
    End Function

    Public Function GetApplicationIcon(fileName As String) As ImageSource
        Dim ai As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(ai.Handle, New Int32Rect(0, 0, ai.Width, ai.Height), BitmapSizeOptions.FromEmptyOptions())
    End Function

    Public Function SplitList(Of T)(ByVal list As List(Of T), ByVal count As Integer) As List(Of T)()
        Dim lists As New List(Of List(Of T))
        Dim itemCount As Integer = list.Count
        Dim maxCount = CInt(Math.Ceiling(itemCount / count))
        Dim skipCount = 0

        For number = count To 1 Step -1
            Dim takeCount = Math.Min(maxCount, CInt(Math.Ceiling(itemCount / number)))

            lists.Add(list.Skip(skipCount).Take(takeCount).ToList())
            itemCount -= takeCount
            skipCount += takeCount
        Next

        Return lists.ToArray()
    End Function

End Module
