Public Class EditWindow
    Dim ListOfPic As New System.Data.DataTable("ImageList")
    Dim config As XElement


    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        config = XElement.Load("config.xml")
        Try
            ListOfPic = MainWindow.ListOfPic.Clone
            ListOfPic.ReadXml(New IO.StringReader(config.Element("DocumentElement").ToString))
        Catch
            ListOfPic = MainWindow.ListOfPic.Copy
        End Try
        LB_Pic.ItemsSource = ListOfPic.Rows

        'LB_Pic.ItemTemplate = New DataTemplate(GetType(String)) With {.}
        'For Each i As System.Data.DataRow In ListOfPic.Rows
        '    LB_Pic.Items.Add(i("Path"))
        'Next
    End Sub

    Private Sub LB_Pic_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles LB_Pic.SelectionChanged
        If e.AddedItems.Count = 1 Then
            Dim row = ListOfPic.Rows.Find(e.AddedItems(0)("Path"))
            If row IsNot Nothing Then
                If Not IsDBNull(row("Date")) Then
                    T_DateShown.Text = Date.Parse(row("Date")).ToString("yyyy.M")
                Else
                    T_DateShown.Text = Application.Current.Resources("none")
                End If
                CB_Title.IsChecked = row("CB_Title")
                TB_TextShown.Text = row("Text")

                'showing preview
                Task.Run(Sub()
                             Dim pic As New BitmapImage
                             Try
                                 Dim imgpath As String = e.AddedItems(0)("Path")
                                 Using strm = New IO.FileStream(imgpath, IO.FileMode.Open, IO.FileAccess.Read)
                                     pic.BeginInit()
                                     pic.DecodePixelWidth = 250
                                     pic.CacheOption = BitmapCacheOption.OnLoad
                                     pic.StreamSource = strm
                                     pic.EndInit()
                                 End Using
                             Catch
                                 pic = New BitmapImage
                                 Dim encoder As New BmpBitmapEncoder
                                 Dim bmpsource = BitmapSource.Create(64, 64, 96, 96, PixelFormats.Indexed1, BitmapPalettes.BlackAndWhite, New Byte(64 * 8) {}, 8)
                                 encoder.Frames.Add(BitmapFrame.Create(bmpsource))
                                 Using ms As New IO.MemoryStream
                                     encoder.Save(ms)
                                     ms.Position = 0
                                     pic.BeginInit()
                                     pic.StreamSource = ms
                                     pic.EndInit()
                                 End Using
                             Finally
                                 pic.Freeze()
                             End Try
                             Dispatcher.Invoke(Sub()
                                                   Dim w = pic.PixelWidth
                                                   Dim h = pic.PixelHeight
                                                   Dim tmpcan As New Canvas
                                                   tmpcan.Height = h
                                                   tmpcan.Width = w
                                                   Dim tmpimg As New Image
                                                   With tmpimg
                                                       .Source = pic
                                                       .Height = h
                                                       .Width = w
                                                   End With
                                                   tmpcan.Children.Add(tmpimg)
                                                   If Not IsDBNull(e.AddedItems(0)("Date")) Then
                                                       Dim date_tb As New TextBlock
                                                       With date_tb
                                                           .Text = T_DateShown.Text
                                                           .FontFamily = New FontFamily("Georgia")
                                                           .FontSize = h / 12
                                                           .FontStyle = FontStyles.Italic
                                                           .Foreground = Brushes.White
                                                           .Effect = New Effects.DropShadowEffect With {.ShadowDepth = 2, .Opacity = 0.8}
                                                       End With
                                                       tmpcan.Children.Add(date_tb)
                                                       date_tb.Margin = New Thickness(w / 5, h * 0.75, 0, 0)
                                                   End If
                                                   If Not TB_TextShown.Text = "" Then
                                                       Dim text_tb As New TextBlock
                                                       With text_tb
                                                           .SnapsToDevicePixels = True
                                                           .Text = TB_TextShown.Text
                                                           .FontFamily = New FontFamily("Georgia")
                                                           .FontSize = h / 12
                                                           .Foreground = Brushes.White
                                                           .Effect = New Effects.DropShadowEffect With {.ShadowDepth = 2, .Opacity = 0.8}
                                                       End With
                                                       tmpcan.Children.Add(text_tb)
                                                       text_tb.Margin = New Thickness(w * 0.7, h * 0.15, 0, 0)
                                                   End If
                                                   tmpcan.Arrange(New Rect(0, 0, w, h))
                                                   Dim tmpbmp As New RenderTargetBitmap(w * 144 / 96, h * 144 / 96, 144, 144, PixelFormats.Default)
                                                   tmpbmp.Render(tmpcan)
                                                   IMG_Preview.Source = tmpbmp
                                               End Sub)
                         End Sub)
            End If
        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Using lop_str = New IO.StringWriter()
            ListOfPic.WriteXml(lop_str)
            Dim lop As XElement = XElement.Parse(lop_str.ToString)
            config.Elements("DocumentElement").Remove()
            config.Add(lop)
            config.Save("config.xml")
        End Using
        
        'Using wt = Xml.XmlWriter.Create("config.xml")
        '    wt.WriteStartElement("ConfigSlideRoot")
        '    wt.WriteStartElement("folders_image")
        '    For Each folder In MainWindow.folders_image
        '        wt.WriteStartElement("dir")
        '        wt.WriteCData(folder)
        '        wt.WriteEndElement()
        '    Next
        '    wt.WriteEndElement()
        '    ListOfPic.WriteXml(wt)
        '    wt.WriteEndElement()
        'End Using
        Me.Close()
    End Sub

    Private Sub CB_Title_Checked(sender As Object, e As RoutedEventArgs) Handles CB_Title.Checked, CB_Title.Unchecked
        If Me.IsLoaded Then
            For Each i In LB_Pic.SelectedItems
                ListOfPic.Rows.Find(i("Path"))("CB_Title") = CB_Title.IsChecked
            Next
        End If
    End Sub

    Private Sub Btn_Up_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Up.Click
        Dim allselected As New List(Of System.Data.DataRow)
        For Each row As System.Data.DataRow In LB_Pic.SelectedItems
            Dim i = ListOfPic.Rows.IndexOf(row)
            If i > 0 Then
                Dim newrow = ListOfPic.NewRow
                newrow.ItemArray = row.ItemArray
                ListOfPic.Rows.Remove(row)
                ListOfPic.Rows.InsertAt(newrow, i - 1)
                allselected.Add(newrow)
            Else
                Exit Sub
            End If
        Next
        LB_Pic.Items.Refresh()
        For Each row In allselected
            LB_Pic.SelectedItems.Add(row)
        Next
        LB_Pic.Focus()
    End Sub

    Private Sub Btn_Down_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Down.Click
        Dim maxi As Integer = -1
        For m = 0 To LB_Pic.SelectedItems.Count - 1
            Dim i = ListOfPic.Rows.IndexOf(LB_Pic.SelectedItems(m))
            If i > maxi Then
                maxi = i
            End If
        Next
        If maxi + 1 <= ListOfPic.Rows.Count - 1 Then
            Dim row As System.Data.DataRow = ListOfPic.Rows(maxi + 1)
            Dim newrow = ListOfPic.NewRow
            newrow.ItemArray = row.ItemArray
            ListOfPic.Rows.Remove(row)
            ListOfPic.Rows.InsertAt(newrow, ListOfPic.Rows.IndexOf(LB_Pic.SelectedItems(0)))
        Else
            Exit Sub
        End If

        LB_Pic.Items.Refresh()
        LB_Pic.Focus()
    End Sub

    Private Sub Btn_Reset_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Reset.Click
        LB_Pic.ItemsSource = Nothing
        Using tmpListOfPic = ListOfPic.Copy
            ListOfPic.Clear()
            For Each fd In MainWindow.folders_image
                If My.Computer.FileSystem.DirectoryExists(fd) Then
                    For Each f In My.Computer.FileSystem.GetFiles(fd)
                        Dim filefullname = My.Computer.FileSystem.GetName(f)
                        Dim filename = IO.Path.GetFileNameWithoutExtension(filefullname)
                        Dim ext = IO.Path.GetExtension(filefullname)
                        If MainWindow.PicFormats.Contains(ext.ToLower) Then
                            Dim tmprow = ListOfPic.NewRow
                            tmprow("Path") = f
                            Dim tmpdate As Date
                            If DateTime.TryParse(filename, tmpdate) Then
                                tmprow("Date") = DateTime.Parse(filename).ToString
                            End If
                            Dim match = tmpListOfPic.Rows.Find(f)
                            If match IsNot Nothing Then
                                For i = 2 To ListOfPic.Columns.Count - 1
                                    tmprow(i) = match(i)
                                Next
                            End If
                            ListOfPic.Rows.Add(tmprow)
                        End If
                    Next
                End If
            Next
        End Using
        LB_Pic.ItemsSource = ListOfPic.Rows
    End Sub

    Private Sub TB_TextShown_LostFocus(sender As Object, e As RoutedEventArgs) Handles TB_TextShown.LostFocus
        If Me.IsLoaded Then
            For Each i In LB_Pic.SelectedItems
                ListOfPic.Rows.Find(i("Path"))("Text") = TB_TextShown.Text
            Next
        End If
    End Sub
End Class
