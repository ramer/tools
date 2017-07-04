Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Printing

Public Class wndMain
    Public WithEvents scaner As New clsScaner
    Public hosts_scaner As New clsThreadSafeObservableCollection(Of clsHostTiny)

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        CType(FindResource("hosts_ro"), CollectionViewSource).Source = hosts
        CType(FindResource("hosts_rw"), CollectionViewSource).Source = hosts
        CType(FindResource("hosts_scaner"), CollectionViewSource).Source = hosts_scaner

        For Each h In hosts
            h.CreateRealTimeModel()
        Next

        dtpHistoryFrom.Value = Now.AddDays(-1)
        dtpHistoryTo.Value = Now
    End Sub

    Private Sub wndMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If preferences.CloseOnXButton Then Application.Current.Shutdown()
    End Sub

    Private Async Sub tbScanerRange_KeyDown(sender As Object, e As KeyEventArgs) Handles tbScanerRange.KeyDown
        If e.Key = Key.Enter Then

            If tbScanerRange.Text.Length <= 0 Then Exit Sub

            tbScanerRange.SelectAll()

            pbScaner.Visibility = Visibility.Visible

            Await scaner.BasicScanAsync(hosts_scaner, tbScanerRange.Text)

            pbScaner.Visibility = Visibility.Hidden

        End If
    End Sub

#Region "Menu"

    Private Sub mnuFileExit_Click(sender As Object, e As RoutedEventArgs) Handles mnuFileExit.Click
        Application.Current.Shutdown()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        Dim w As New wndPreferences With {.Owner = Me}
        w.ShowDialog()
    End Sub

    Private Sub mnuServiceError_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceError.Click
        Dim w As wndError

        For Each wnd As Window In Application.Current.Windows
            If GetType(wndError) Is wnd.GetType Then
                w = wnd
                w.Show()
                w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                Exit Sub
            End If
        Next

        w = New wndError
        w.Show()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim w As New wndAbout With {.Owner = Me}
        w.ShowDialog()
    End Sub

    Private Sub ctxmnuHosts_Opened(sender As Object, e As RoutedEventArgs) Handles ctxmnuHosts.Opened
        ctxmnuHostsExternalSoftware.Items.Clear()
        For Each es As clsExternalSoftware In preferences.ExternalSoftware
            Dim esmnu As New MenuItem
            esmnu.Header = es.Label
            esmnu.Icon = New Image With {.Source = es.Image}
            esmnu.Tag = es
            AddHandler esmnu.PreviewMouseDown, AddressOf ctxmnuHostsExternalSoftwareItem_PreviewMouseDown
            ctxmnuHostsExternalSoftware.Items.Add(esmnu)
        Next
    End Sub

    Private Sub ctxmnuScaner_Opened(sender As Object, e As RoutedEventArgs) Handles ctxmnuScaner.Opened
        ctxmnuScanerExternalSoftware.Items.Clear()
        For Each es As clsExternalSoftware In preferences.ExternalSoftware
            Dim esmnu As New MenuItem
            esmnu.Header = es.Label
            esmnu.Icon = New Image With {.Source = es.Image}
            esmnu.Tag = es
            AddHandler esmnu.PreviewMouseDown, AddressOf ctxmnuScanerExternalSoftwareItem_PreviewMouseDown
            ctxmnuScanerExternalSoftware.Items.Add(esmnu)
        Next
    End Sub

    Private Sub ctxmnuHostsExternalSoftwareItem_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        If dgHosts.SelectedItem Is Nothing Then Exit Sub
        Dim obj As clsHost = dgHosts.SelectedItem
        Dim esmnu As MenuItem = CType(sender, MenuItem)
        If esmnu Is Nothing Then Exit Sub
        Dim es As clsExternalSoftware = CType(esmnu.Tag, clsExternalSoftware)
        If es Is Nothing Then Exit Sub

        Dim args As String = es.Arguments
        If args Is Nothing Then args = ""

        args = Replace(args, "%hostname%", obj.Hostname)
        args = Replace(args, "%ipaddress%", obj.IPAddress)
        args = Replace(args, "%macaddress%", obj.MACAddress)

        Dim psi As New ProcessStartInfo(es.Path, args)

        psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
        psi.UseShellExecute = False
        Process.Start(psi)
    End Sub

    Private Sub ctxmnuScanerExternalSoftwareItem_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        If dgScaner.SelectedItem Is Nothing Then Exit Sub

        Dim obj As clsHostTiny = dgScaner.SelectedItem
        Dim esmnu As MenuItem = CType(sender, MenuItem)
        If esmnu Is Nothing Then Exit Sub
        Dim es As clsExternalSoftware = CType(esmnu.Tag, clsExternalSoftware)
        If es Is Nothing Then Exit Sub

        Dim args As String = es.Arguments
        If args Is Nothing Then args = ""

        args = Replace(args, "%hostname%", obj.Hostname)
        args = Replace(args, "%ipaddress%", obj.IPAddress)
        args = Replace(args, "%macaddress%", obj.MACAddress)

        Dim psi As New ProcessStartInfo(es.Path, args)

        psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
        psi.UseShellExecute = False
        Process.Start(psi)
    End Sub

    Private Sub ctxmnuHostsCopyBasicAttributes_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuHostsCopyBasicAttributes.Click
        If dgHosts.SelectedItems Is Nothing Then Exit Sub

        Try
            Clipboard.SetText(Join(dgHosts.SelectedItems.Cast(Of clsHost).ToArray.Select(Function(x) x.Hostname & vbTab & x.IPAddress & vbTab & x.MACAddress).ToArray, vbCrLf))
        Catch ex As Exception
            ThrowException(ex, "ctxmnuHostsCopyBasicAttributes_Click")
        End Try
    End Sub

    Private Sub ctxmnuHostsSelectAll_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuHostsSelectAll.Click
        dgHosts.SelectAll()
    End Sub

    Private Sub ctxmnuHostsRemove_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuHostsRemove.Click
        If dgHosts.SelectedItems Is Nothing Then Exit Sub

        Try
            If IMsgBox("Вы уверены?", "Удаление объектов из списка", vbYesNo, vbQuestion) = vbYes Then
                Dim itemstoremove As New ObservableCollection(Of clsHost)

                For Each h As Object In dgHosts.SelectedItems
                    If h.GetType Is GetType(clsHost) Then itemstoremove.Add(h)
                Next
                For Each h As clsHost In itemstoremove
                    hosts.Remove(h)
                Next
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuHostsRemove_Click")
        End Try
    End Sub

    Private Sub ctxmnuHostsHistoryClear_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuHostsHistoryClear.Click
        If dgHosts.SelectedItems Is Nothing Then Exit Sub

        Try
            If IMsgBox("Вы уверены?", "Очистка истории", vbYesNo, vbQuestion) = vbYes Then
                For Each h As Object In dgHosts.SelectedItems
                    If h.GetType Is GetType(clsHost) Then CType(h, clsHost).HistoryClear()
                Next
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuHostsHistoryClear_Click")
        End Try
    End Sub

    Private Sub ctxmnuProperties_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuProperties.Click
        If dgHosts.SelectedItem Is Nothing Then Exit Sub

        ShowObjectProperties(dgHosts.SelectedItem, Me)
    End Sub

    Private Sub ctxmnuScanerAddToHosts_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuScanerAddToHosts.Click
        If dgScaner.SelectedItems Is Nothing Then Exit Sub

        For Each h As clsHostTiny In dgScaner.SelectedItems
            hosts.Add(New clsHost(h.IPAddress, h.Hostname, h.MACAddress))
        Next
    End Sub

    Private Sub ctxmnuScanerCopyBasicAttributes_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuScanerCopyBasicAttributes.Click
        If dgScaner.SelectedItems Is Nothing Then Exit Sub

        Try
            Clipboard.SetText(Join(dgScaner.SelectedItems.Cast(Of clsHostTiny).ToArray.Select(Function(x) x.Hostname & vbTab & x.IPAddress & vbTab & x.MACAddress).ToArray, vbCrLf))
        Catch ex As Exception
            ThrowException(ex, "ctxmnuScanerCopyBasicAttributes_Click")
        End Try
    End Sub

    Private Sub ctxmnuScanerSelectAll_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuScanerSelectAll.Click
        dgScaner.SelectAll()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click
        Dim PrintDialog As PrintDialog = New PrintDialog()
        If (PrintDialog.ShowDialog() = True) Then
            PrintDialog.PrintTicket.PageMediaSize = New PageMediaSize(1920, 1080)
            PrintDialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape
            PrintDialog.PrintVisual(grdHistory, "some descr")
        End If
    End Sub

    Private Sub btnDummy_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles btnDummy.MouseDoubleClick
        If e.ChangedButton = MouseButton.Middle Then MsgBox("ura!")
    End Sub

    Private Sub dtpHistoryFrom_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpHistoryFrom.ValueChanged
        If dtpHistoryFrom.Value >= dtpHistoryTo.Value Then dtpHistoryTo.Value = dtpHistoryFrom.Value.Value.AddMinutes(1)
        UpdateHistoryPlots()
    End Sub

    Private Sub dtpHistoryTo_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpHistoryTo.ValueChanged
        If dtpHistoryTo.Value <= dtpHistoryFrom.Value Then dtpHistoryFrom.Value = dtpHistoryTo.Value.Value.AddMinutes(-1)
        UpdateHistoryPlots()
    End Sub

    Private Sub hlHistory10min_Click(sender As Object, e As RoutedEventArgs) Handles hlHistory10min.Click
        dtpHistoryTo.Value = Now
        dtpHistoryFrom.Value = Now.AddMinutes(-10)
    End Sub

    Private Sub hlHistory1hour_Click(sender As Object, e As RoutedEventArgs) Handles hlHistory1hour.Click
        dtpHistoryTo.Value = Now
        dtpHistoryFrom.Value = Now.AddHours(-1)
    End Sub

    Private Sub hlHistory6hour_Click(sender As Object, e As RoutedEventArgs) Handles hlHistory6hour.Click
        dtpHistoryTo.Value = Now
        dtpHistoryFrom.Value = Now.AddHours(-6)
    End Sub

    Private Sub hlHistory1day_Click(sender As Object, e As RoutedEventArgs) Handles hlHistory1day.Click
        dtpHistoryTo.Value = Now
        dtpHistoryFrom.Value = Now.AddDays(-1)
    End Sub

    Private Sub hlHistory7days_Click(sender As Object, e As RoutedEventArgs) Handles hlHistory7days.Click
        dtpHistoryTo.Value = Now
        dtpHistoryFrom.Value = Now.AddDays(-7)
    End Sub

    Private Sub UpdateHistoryPlots()
        pvHistory.Model.Axes(0).Reset()
        pvHistory.Model.Axes(0).Minimum = OxyPlot.Axes.DateTimeAxis.ToDouble(dtpHistoryFrom.Value)
        pvHistory.Model.Axes(0).Maximum = OxyPlot.Axes.DateTimeAxis.ToDouble(dtpHistoryTo.Value)
        pvHistory.Model.InvalidatePlot(True)
    End Sub

#End Region

End Class
