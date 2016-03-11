Public Class ControlWindow
    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        If Not MainWindow.reallyclose Then
            e.Cancel = True
            Me.Hide() 'if close, next F1 will throw exception
        End If
    End Sub

    Private Sub Btn_SwitchImage_Click(sender As Object, e As RoutedEventArgs) Handles Btn_SwitchImage.Click
        Task.Run(Sub()
                     Dispatcher.Invoke(Sub() Btn_SwitchImage.IsEnabled = False)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = False)
                     System.Threading.Thread.Sleep(3000)
                     Dispatcher.Invoke(Sub() Btn_SwitchImage.IsEnabled = True)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = True)
                 End Sub)
        Dim mainwin As MainWindow = Me.Owner
        mainwin.SwitchImage()
    End Sub

    Private Sub Btn_SwitchAudio_Click(sender As Object, e As RoutedEventArgs) Handles Btn_SwitchAudio.Click
        Task.Run(Sub()
                     Dispatcher.Invoke(Sub() Btn_SwitchAudio.IsEnabled = False)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = False)
                     System.Threading.Thread.Sleep(3000)
                     Dispatcher.Invoke(Sub() Btn_SwitchAudio.IsEnabled = True)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = True)
                 End Sub)
        Dim mainwin As MainWindow = Me.Owner
        mainwin.SwitchAudio()
    End Sub

    Private Sub Btn_Restart_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Restart.Click
        Task.Run(Sub()
                     Dispatcher.Invoke(Sub() Btn_SwitchImage.IsEnabled = False)
                     Dispatcher.Invoke(Sub() Btn_SwitchAudio.IsEnabled = False)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = False)
                     System.Threading.Thread.Sleep(3000)
                     Dispatcher.Invoke(Sub() Btn_SwitchImage.IsEnabled = True)
                     Dispatcher.Invoke(Sub() Btn_SwitchAudio.IsEnabled = True)
                     Dispatcher.Invoke(Sub() Btn_Restart.IsEnabled = True)
                 End Sub)
        Dim mainwin As MainWindow = Me.Owner
        mainwin.RestartAll()
    End Sub

    Private Sub Btn_Options_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Options.Click
        Dim optwin As New OptWindow
        optwin.ShowDialog()
        optwin.Close()
    End Sub

    Private Sub Btn_EditSlide_Click(sender As Object, e As RoutedEventArgs) Handles Btn_EditSlide.Click
        Dim editwin As New EditWindow
        editwin.ShowDialog()
        editwin.Close()
    End Sub

    Private Sub Btn_Exit_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Exit.Click
        Dim mainwin As MainWindow = Me.Owner
        mainwin.Close()
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Escape Then
            Me.Hide()
        End If
    End Sub
End Class
