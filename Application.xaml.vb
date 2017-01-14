Class Application

    ' 应用程序级事件(例如 Startup、Exit 和 DispatcherUnhandledException)
    ' 可以在此文件中进行处理。
    Private Sub Application_Startup()
        Dim dict As New ResourceDictionary
        Select Case Threading.Thread.CurrentThread.CurrentCulture.Name.Substring(0, 2)
            Case "zh"
                dict.Source = New Uri("Localization/StringResources.zh-CN.xaml", UriKind.Relative)
            Case Else
                dict.Source = New Uri("Localization/StringResources.xaml", UriKind.Relative)
        End Select
        Me.Resources.MergedDictionaries.Add(dict)

        Dim win As New MainWindow
        Application.Current.MainWindow = win
        win.Show()
    End Sub
End Class
