Public Class OptWindow
    Dim lastPath As String

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        For Each i In MainWindow.folders_image
            LB_ImgFolder.Items.Add(i)
        Next
        For Each i In MainWindow.folders_music
            LB_BGMFolder.Items.Add(i)
        Next
        CB_VerLk.IsChecked = MainWindow.verticalLock
        CB_ResLk.IsChecked = MainWindow.resolutionLock
        CB_VerOptm.IsChecked = MainWindow.verticalOptimize
        CB_HorOptm.IsChecked = MainWindow.horizontalOptimize
        CB_Fadeout.IsChecked = MainWindow.fadeout
        TB_VORatio.Text = MainWindow.verticalOptimizeR
        TB_HORatio.Text = MainWindow.horizontalOptimizeR
        TB_Framerate.Text = MainWindow.framerate
        TB_Duration.Text = MainWindow.duration
        CbB_Transit.Text = MainWindow.transit
    End Sub

    Private Sub ProfileUpdate() Handles CB_Fadeout.Checked, CB_Fadeout.Unchecked, CB_ResLk.Checked, CB_ResLk.Unchecked
        If Me.IsLoaded Then
            If CB_Fadeout.IsChecked = True AndAlso CB_ResLk.IsChecked = False Then
                RB_Qlty.IsChecked = True
            ElseIf CB_Fadeout.IsChecked = False AndAlso CB_ResLk.IsChecked = True Then
                RB_Perf.IsChecked = True
            Else
                RB_Custom.IsChecked = True
            End If
        End If
    End Sub

    Private Sub CB_HorOptm_Change() Handles CB_HorOptm.Unchecked, CB_HorOptm.Checked
        If Me.IsLoaded Then
            If CB_HorOptm.IsChecked Then
                TB_HORatio.IsEnabled = True
            Else
                TB_HORatio.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub CB_VerOptm_Change() Handles CB_VerOptm.Unchecked, CB_VerOptm.Checked
        If Me.IsLoaded Then
            If CB_VerOptm.IsChecked Then
                TB_VORatio.IsEnabled = True
            Else
                TB_VORatio.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub RB_Perf_Checked(sender As Object, e As RoutedEventArgs) Handles RB_Perf.Checked
        If Me.IsLoaded Then
            CB_Fadeout.IsChecked = False
            CB_ResLk.IsChecked = True
        End If
    End Sub

    Private Sub RB_Qlty_Checked(sender As Object, e As RoutedEventArgs) Handles RB_Qlty.Checked
        If Me.IsLoaded Then
            CB_Fadeout.IsChecked = True
            CB_ResLk.IsChecked = False
        End If
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As RoutedEventArgs) Handles Btn_OK.Click
        Dim tmpR As Double = 0
        If Double.TryParse(TB_VORatio.Text, tmpR) AndAlso tmpR > 0 AndAlso tmpR <= 1 Then
            MainWindow.verticalOptimizeR = tmpR
        Else
            MessageBox.Show("Invalid vertical optimize ratio.", "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Exit Sub
        End If
        tmpR = 0
        If Double.TryParse(TB_HORatio.Text, tmpR) AndAlso tmpR > 0 AndAlso tmpR <= 1 Then
            MainWindow.horizontalOptimizeR = tmpR
        Else
            MessageBox.Show("Invalid horizontal optimize ratio.", "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Exit Sub
        End If
        Dim tmpF As UInteger = 0
        If UInteger.TryParse(TB_Framerate.Text, tmpF) AndAlso tmpF > 0 Then
            MainWindow.framerate = tmpF
        Else
            MessageBox.Show("Invalid framerate value.", "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Exit Sub
        End If
        tmpF = 0
        If UInteger.TryParse(TB_Duration.Text, tmpF) AndAlso tmpF > 4 Then
            'not setting mainwindow.picmove_sec to avoid problems. save to config.xml instead for the next load.
            MainWindow.duration = tmpF
        Else
            MessageBox.Show("Invalid duration value.", "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Exit Sub
        End If

        MainWindow.folders_image.Clear()
        For Each i As String In LB_ImgFolder.Items
            MainWindow.folders_image.Add(i)
        Next
        MainWindow.folders_music.Clear()
        For Each i As String In LB_BGMFolder.Items
            MainWindow.folders_music.Add(i)
        Next
        MainWindow.verticalLock = CB_VerLk.IsChecked
        MainWindow.verticalOptimize = CB_VerOptm.IsChecked
        MainWindow.horizontalOptimize = CB_HorOptm.IsChecked
        MainWindow.resolutionLock = CB_ResLk.IsChecked
        MainWindow.fadeout = CB_Fadeout.IsChecked
        MainWindow.horizontalOptimizeR = TB_HORatio.Text
        MainWindow.transit = CbB_Transit.Text

        'saving to file
        Dim config As New XElement("CfgRoot")

        config.Add(New XElement("PicDir"))
        For Each i As String In LB_ImgFolder.Items
            config.Element("PicDir").Add(New XElement("dir", New XCData(i)))
        Next
        config.Add(New XElement("Music"))
        For Each i As String In LB_BGMFolder.Items
            config.Element("Music").Add(New XElement("dir", New XCData(i)))
        Next
        config.Add(New XElement("Framerate", TB_Framerate.Text))
        config.Add(New XElement("Duration", TB_Duration.Text))
        config.Add(New XElement("VerticalLock", CB_VerLk.IsChecked.Value))
        config.Add(New XElement("ResolutionLock", CB_ResLk.IsChecked.Value))
        config.Add(New XElement("VerticalOptimize", CB_VerOptm.IsChecked.Value))
        config.Add(New XElement("HorizontalOptimize", CB_HorOptm.IsChecked.Value))
        config.Add(New XElement("Fadeout", CB_Fadeout.IsChecked.Value))
        config.Add(New XElement("VerticalOptimizeRatio", TB_VORatio.Text))
        config.Add(New XElement("HorizontalOptimizeRatio", TB_HORatio.Text))
        config.Add(New XElement("Transit", CbB_Transit.Text))

        config.Save("config.xml")
        Me.Close()
    End Sub

    Private Sub Btn_Img_Add_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Img_Add.Click, Btn_BGM_Add.Click
        Using dialog As New Forms.FolderBrowserDialog
            dialog.Description = "Select a folder."
            dialog.SelectedPath = lastPath
            If dialog.ShowDialog = Forms.DialogResult.OK Then
                Dim c As Object = VisualTreeHelper.GetParent(sender)
                c.Children(0).Items.Add(dialog.SelectedPath)
                lastPath = dialog.SelectedPath
            End If
        End Using
    End Sub

    Private Sub Btn_Img_Rmv_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Img_Rmv.Click, Btn_BGM_Rmv.Click
        Dim c As Object = VisualTreeHelper.GetParent(sender)
        Dim lb As ListBox = c.Children(0)
        If lb.SelectedIndex <> -1 Then
            lb.Items.RemoveAt(lb.SelectedIndex)
        End If
    End Sub

    Private Sub CbB_Transit_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CbB_Transit.SelectionChanged
        If CbB_Transit.SelectedItem.Content = "Breath" Then
            CB_Fadeout.IsEnabled = False
            RB_Perf.IsEnabled = False
            RB_Qlty.IsEnabled = False
            RB_Custom.IsEnabled = False
        Else
            CB_Fadeout.IsEnabled = True
            RB_Perf.IsEnabled = True
            RB_Qlty.IsEnabled = True
            RB_Custom.IsEnabled = True
        End If
    End Sub
End Class
