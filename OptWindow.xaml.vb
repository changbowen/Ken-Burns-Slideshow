Public Class OptWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim config As XElement
        If My.Computer.FileSystem.FileExists("config.xml") Then
            config = XElement.Load("config.xml")
        Else
            MsgBox("Config.xml file is not found at application root.", MsgBoxStyle.Exclamation)
            Me.Close()
            Exit Sub
        End If
        If config.Elements("VerticalLock").Any AndAlso config.Element("VerticalLock").Value.ToLower = "false" Then
            CB_VerLk.IsChecked = False
        End If
        If config.Elements("ResolutionLock").Any AndAlso config.Element("ResolutionLock").Value.ToLower = "false" Then
            CB_ResLk.IsChecked = False
        End If
        If config.Elements("VerticalOptimize").Any AndAlso config.Element("VerticalOptimize").Value.ToLower = "false" Then
            CB_VerOptm.IsChecked = False
        End If
        If config.Elements("HorizontalOptimize").Any AndAlso config.Element("HorizontalOptimize").Value.ToLower = "false" Then
            CB_HorOptm.IsChecked = False
        End If
        If config.Elements("Fadeout").Any AndAlso config.Element("Fadeout").Value.ToLower = "false" Then
            CB_Fadeout.IsChecked = False
        End If
    End Sub

    Private Sub DetailChange() Handles CB_VerOptm.Checked, CB_VerOptm.Unchecked, CB_Fadeout.Checked, CB_Fadeout.Unchecked, CB_HorOptm.Unchecked, CB_HorOptm.Checked, CB_ResLk.Checked, CB_ResLk.Unchecked, CB_VerLk.Checked, CB_VerLk.Unchecked
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
        MainWindow.verticalLock = CB_VerLk.IsChecked
        MainWindow.verticalOptimize = CB_VerOptm.IsChecked
        MainWindow.horizontalOptimize = CB_HorOptm.IsChecked
        MainWindow.resolutionLock = CB_ResLk.IsChecked
        MainWindow.fadeout = CB_Fadeout.IsChecked

        'saving to file
        Dim config As XElement
        If My.Computer.FileSystem.FileExists("config.xml") Then
            config = XElement.Load("config.xml")
        Else
            MsgBox("Config.xml file is not found at application root.", MsgBoxStyle.Exclamation)
            Me.Close()
            Exit Sub
        End If
        If config.Elements("VerticalLock").Any Then
            config.Element("VerticalLock").Value = CB_VerLk.IsChecked.Value
        Else
            config.Add(New XElement("VerticalLock", CB_VerLk.IsChecked.Value))
        End If
        If config.Elements("ResolutionLock").Any Then
            config.Element("ResolutionLock").Value = CB_ResLk.IsChecked.Value
        Else
            config.Add(New XElement("ResolutionLock", CB_ResLk.IsChecked.Value))
        End If
        If config.Elements("VerticalOptimize").Any Then
            config.Element("VerticalOptimize").Value = CB_VerOptm.IsChecked.Value
        Else
            config.Add(New XElement("VerticalOptimize", CB_VerOptm.IsChecked.Value))
        End If
        If config.Elements("HorizontalOptimize").Any Then
            config.Element("HorizontalOptimize").Value = CB_HorOptm.IsChecked.Value
        Else
            config.Add(New XElement("HorizontalOptimize", CB_HorOptm.IsChecked.Value))
        End If
        If config.Elements("Fadeout").Any Then
            config.Element("Fadeout").Value = CB_Fadeout.IsChecked.Value
        Else
            config.Add(New XElement("Fadeout", CB_Fadeout.IsChecked.Value))
        End If
        config.Save("config.xml")
        Me.Close()
    End Sub
End Class
