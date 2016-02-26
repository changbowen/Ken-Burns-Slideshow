Public Class OptWindow
    Dim lastPath As String

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        CbB_ScaleMode.ItemsSource = MainWindow.ScaleMode_Dic
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
        TB_LoadQuality.Text = MainWindow.loadquality
        CbB_ScaleMode.SelectedItem = New KeyValuePair(Of Integer, String)(MainWindow.scalemode, MainWindow.ScaleMode_Dic(MainWindow.scalemode))
        CbB_BlurMode.Text = MainWindow.blurmode
    End Sub

    'Private Sub ProfileUpdate() Handles CB_Fadeout.Checked, CB_Fadeout.Unchecked
    '    If Me.IsLoaded Then
    '        If CB_Fadeout.IsChecked = True AndAlso CB_ResLk.IsChecked = False Then
    '            RB_Qlty.IsChecked = True
    '        ElseIf CB_Fadeout.IsChecked = False AndAlso CB_ResLk.IsChecked = True Then
    '            RB_Perf.IsChecked = True
    '        Else
    '            RB_Custom.IsChecked = True
    '        End If
    '    End If
    'End Sub

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

    'Private Sub RB_Perf_Checked(sender As Object, e As RoutedEventArgs) Handles RB_Perf.Checked
    '    If Me.IsLoaded Then
    '        CB_Fadeout.IsChecked = False
    '        CB_ResLk.IsChecked = True
    '    End If
    'End Sub

    'Private Sub RB_Qlty_Checked(sender As Object, e As RoutedEventArgs) Handles RB_Qlty.Checked
    '    If Me.IsLoaded Then
    '        CB_Fadeout.IsChecked = True
    '        CB_ResLk.IsChecked = False
    '    End If
    'End Sub

    Private Function CheckConsist(chk_str As String, ByRef target As Double, min As Double, max As Double, inclu_min As Boolean, inclu_max As Boolean, err_msg As String) As Boolean
        Dim tmp As Double = 0
        If Double.TryParse(chk_str, tmp) AndAlso If(inclu_min, tmp >= min, tmp > min) AndAlso If(inclu_max, tmp <= max, tmp < max) Then
            target = tmp
            Return True
        Else
            MessageBox.Show(err_msg, "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Return False
        End If
    End Function

    Private Function CheckConsist(chk_str As String, ByRef target As UInteger, min As UInteger, max As UInteger, inclu_min As Boolean, inclu_max As Boolean, err_msg As String) As Boolean
        Dim tmp As UInteger = 0
        If UInteger.TryParse(chk_str, tmp) AndAlso If(inclu_min, tmp >= min, tmp > min) AndAlso If(inclu_max, tmp <= max, tmp < max) Then
            target = tmp
            Return True
        Else
            MessageBox.Show(err_msg, "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Return False
        End If
    End Function

    Private Function CheckConsist(chk_str As String, ByRef target As Double, min As Double, inclu_min As Boolean, err_msg As String) As Boolean
        Dim tmp As Double = 0
        If Double.TryParse(chk_str, tmp) AndAlso If(inclu_min, tmp >= min, tmp > min) Then
            target = tmp
            Return True
        Else
            MessageBox.Show(err_msg, "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Return False
        End If
    End Function

    Private Function CheckConsist(chk_str As String, ByRef target As UInteger, min As UInteger, inclu_min As Boolean, err_msg As String) As Boolean
        Dim tmp As UInteger = 0
        If UInteger.TryParse(chk_str, tmp) AndAlso If(inclu_min, tmp >= min, tmp > min) Then
            target = tmp
            Return True
        Else
            MessageBox.Show(err_msg, "", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Return False
        End If
    End Function

    Private Sub Btn_OK_Click(sender As Object, e As RoutedEventArgs) Handles Btn_OK.Click
        If Not CheckConsist(TB_VORatio.Text, MainWindow.verticalOptimizeR, 0, 1, False, True, "Invalid vertical optimize ratio.") Then Exit Sub
        If Not CheckConsist(TB_HORatio.Text, MainWindow.horizontalOptimizeR, 0, 1, False, True, "Invalid horizontal optimize ratio.") Then Exit Sub
        If Not CheckConsist(TB_Framerate.Text, MainWindow.framerate, 0, False, "Invalid framerate value.") Then Exit Sub
        If Not CheckConsist(TB_Duration.Text, MainWindow.duration, 4, False, "Invalid duration value.") Then Exit Sub
        If Not CheckConsist(TB_LoadQuality.Text, MainWindow.loadquality, 0, 2, False, True, "Invalid multiplier value.") Then Exit Sub

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
        MainWindow.transit = CbB_Transit.Text
        MainWindow.scalemode = CbB_ScaleMode.SelectedItem.Key
        MainWindow.blurmode = CbB_BlurMode.Text

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
        config.Add(New XElement("LoadQuality", TB_LoadQuality.Text))
        config.Add(New XElement("ScaleMode", CbB_ScaleMode.SelectedItem.Key))
        config.Add(New XElement("BlurMode", CbB_BlurMode.Text))
        Using lop_str = New IO.StringWriter()
            MainWindow.ListOfPic.WriteXml(lop_str)
            Dim lop = XElement.Parse(lop_str.ToString)
            config.Add(lop)
            config.Save("config.xml")
        End Using

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

    Private Sub Btn_Img_Edit_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Img_Edit.Click, Btn_BGM_Edit.Click, LB_ImgFolder.MouseDoubleClick, LB_BGMFolder.MouseDoubleClick
        Dim c As Object = VisualTreeHelper.GetParent(sender)
        Dim lb As ListBox = c.Children(0)
        If lb.SelectedIndex <> -1 Then
            Using dialog As New Forms.FolderBrowserDialog
                dialog.Description = "Select a folder."
                dialog.SelectedPath = lb.SelectedItem.ToString
                If dialog.ShowDialog = Forms.DialogResult.OK Then
                    lb.Items(lb.SelectedIndex) = dialog.SelectedPath
                    lastPath = dialog.SelectedPath
                End If
            End Using
        End If
    End Sub

    Private Sub CbB_Transit_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CbB_Transit.SelectionChanged
        If Me.IsLoaded Then
            If CbB_Transit.SelectedItem.Content = "Throw" Then
                CB_VerOptm.IsEnabled = False
                CB_HorOptm.IsEnabled = False
                TB_VORatio.IsEnabled = False
                TB_HORatio.IsEnabled = False
            Else
                CB_VerOptm.IsEnabled = True
                CB_HorOptm.IsEnabled = True
                TB_VORatio.IsEnabled = True
                TB_HORatio.IsEnabled = True
            End If
        End If
    End Sub

    Private Sub CB_ResLk_Checked(sender As Object, e As RoutedEventArgs) Handles CB_ResLk.Checked, CB_ResLk.Unchecked
        If Me.IsLoaded Then
            If CB_ResLk.IsChecked Then
                TB_LoadQuality.IsEnabled = True
            Else
                TB_LoadQuality.IsEnabled = False
            End If
        End If
    End Sub
End Class
