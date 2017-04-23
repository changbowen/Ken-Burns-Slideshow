Imports System.Data

Public Class EditWindow
    Dim ListOfPic As New DataTable("ImageList")
    Dim config As XElement
    Dim changingselection As Boolean = False

    Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        Owner = Application.Current.MainWindow
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        config = XElement.Load(MainWindow.config_path)
        Try
            ListOfPic = MainWindow.ListOfPic.Clone
            ListOfPic.ReadXml(New IO.StringReader(config.Element("DocumentElement").ToString))
        Catch
            ListOfPic = MainWindow.ListOfPic.Copy
        End Try
        LB_Pic.ItemsSource = ListOfPic.Rows
    End Sub

    Private Sub RefreshPreview(row As DataRow)
        If row Is Nothing Then
            Dispatcher.Invoke(Sub()
                                  IMG_Preview.Source = Nothing
                                  changingselection = False
                              End Sub)
            Exit Sub
        End If

        Dim pic As New BitmapImage
        Dim w As Double = 256, h As Double = w / MainWindow.w * MainWindow.h
        Try
            Dim imgpath As String = row("Path")
            Using strm = New IO.FileStream(imgpath, IO.FileMode.Open, IO.FileAccess.Read)
                Dim frame = BitmapFrame.Create(strm, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)
                strm.Position = 0
                pic.BeginInit()
                Try
                    'getting orientation value from exif and rotate image. Rotate before setting DecodePixelHeight/Width?
                    Dim ori As UShort = DirectCast(frame.Metadata, BitmapMetadata).GetQuery("/app1/ifd/{ushort=274}")
                    Select Case ori
                        Case 6
                            pic.Rotation = Rotation.Rotate90
                        Case 3
                            pic.Rotation = Rotation.Rotate180
                        Case 8
                            pic.Rotation = Rotation.Rotate270
                    End Select
                    strm.Position = 0
                Catch
                End Try
                pic.DecodePixelWidth = 256
                pic.CacheOption = BitmapCacheOption.OnLoad
                pic.StreamSource = strm
                pic.EndInit()
            End Using
        Catch
            pic = New BitmapImage
            Dim encoder As New BmpBitmapEncoder
            Dim bmpsource = BitmapSource.Create(256, 192, 96, 96, PixelFormats.Indexed1, BitmapPalettes.BlackAndWhite, New Byte(192 * 32) {}, 32)
            encoder.Frames.Add(BitmapFrame.Create(bmpsource))
            Using ms As New IO.MemoryStream
                encoder.Save(ms)
                ms.Position = 0
                pic.BeginInit()
                pic.CacheOption = BitmapCacheOption.OnLoad
                pic.StreamSource = ms
                pic.EndInit()
            End Using
        Finally
            pic.Freeze()
        End Try
        Dispatcher.Invoke(Sub()
                              Dim tmpcan As New Grid
                              tmpcan.Width = w
                              tmpcan.Height = h
                              Dim tmpimg As New Image
                              With tmpimg
                                  .Source = pic
                                  .Stretch = Stretch.UniformToFill
                              End With
                              tmpcan.Children.Add(tmpimg)
                              If Not IsDBNull(row("Date")) Then
                                  Dim date_tb As New TextBlock
                                  With date_tb
                                      .Text = T_DateShown.Text
                                      .FontFamily = New FontFamily("Georgia")
                                      .FontSize = 12
                                      .FontStyle = FontStyles.Italic
                                      .Foreground = Brushes.White
                                      .Effect = New Effects.DropShadowEffect With {.ShadowDepth = 2, .Opacity = 0.8}
                                  End With
                                  tmpcan.Children.Add(date_tb)
                                  date_tb.Margin = New Thickness(w / 5, h * 0.75, 0, 0)
                              End If
                              If Not TB_Text.Text = "" Then
                                  Dim text_tb As New TextBlock
                                  With text_tb
                                      .Text = TB_Text.Text
                                      .FontFamily = New FontFamily(row("FontFamily"))
                                      .FontSize = MainWindow.h / DirectCast(row("FontSize"), Double) / MainWindow.h * h
                                      .Foreground = New SolidColorBrush(ColorConverter.ConvertFromString(row("FontColor")))
                                      .Effect = New Effects.DropShadowEffect With {.ShadowDepth = 2, .Opacity = 0.8}
                                  End With
                                  tmpcan.Children.Add(text_tb)

                                  Dim startw = (w * 0.7) + (w * DirectCast(row("FontOffsetH"), Double) / 100)
                                  Dim starth = (h * 0.15) + (h * DirectCast(row("FontOffsetV"), Double) / 100)
                                  text_tb.Margin = New Thickness(startw, starth, 0, 0)
                              End If
                              tmpcan.Arrange(New Rect(0, 0, w, h))
                              Dim tmpbmp As New RenderTargetBitmap(w * 144 / 96, h * 144 / 96, 144, 144, PixelFormats.Default)
                              tmpbmp.Render(tmpcan)
                              IMG_Preview.Source = tmpbmp
                              changingselection = False
                          End Sub)
    End Sub

    Private Sub LB_Pic_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles LB_Pic.SelectionChanged
        If e.AddedItems.Count = 1 Then
            changingselection = True
            Dim row = ListOfPic.Rows.Find(e.AddedItems(0)("Path"))
            If row IsNot Nothing Then
                If Not IsDBNull(row("Date")) Then
                    T_DateShown.Text = Date.Parse(row("Date")).ToString("yyyy.M")
                Else
                    T_DateShown.Text = Application.Current.Resources("none")
                End If
                CB_Title.IsChecked = row("TitleSlide")
                TB_Text.Text = row("Text")
                CbB_FontFamily.SelectedValue = row("FontFamily")
                Sld_FontSize.Value = row("FontSize")
                CbB_FontColor.SelectedValue = row("FontColor")
                Sld_FontOffsetH.Value = row("FontOffsetH")
                Sld_FontOffsetV.Value = row("FontOffsetV")
                'showing preview
                Task.Run(Sub() RefreshPreview(e.AddedItems(0)))
            End If
        ElseIf e.AddedItems.Count > 1 Then
            changingselection = True
            RefreshPreview(Nothing) 'clear preview
            T_DateShown.Text = ""
            CB_Title.IsChecked = Nothing
            TB_Text.Text = ""
            CbB_FontFamily.SelectedValue = Nothing
            Sld_FontSize.Value = 12
            CbB_FontColor.SelectedValue = Nothing
            Sld_FontOffsetH.Value = 0
            Sld_FontOffsetV.Value = 0
        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Using lop_str = New IO.StringWriter()
            ListOfPic.WriteXml(lop_str)
            Dim lop As XElement = XElement.Parse(lop_str.ToString)
            config.Elements("DocumentElement").Remove()
            config.Add(lop)
            config.Save(MainWindow.config_path)
        End Using

        'Using wt = Xml.XmlWriter.Create(config_path)
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
        If Interop.ComponentDispatcher.IsThreadModal Then DialogResult = True Else Close()
    End Sub

    Private Sub Btn_Up_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Up.Click
        Dim allselected As New List(Of DataRow)
        For Each row As DataRow In LB_Pic.SelectedItems
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
            Dim row As DataRow = ListOfPic.Rows(maxi + 1)
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
        Dim clearall As Boolean = False
        Dim searchopt = If(MainWindow.recursive_folder, FileIO.SearchOption.SearchAllSubDirectories, FileIO.SearchOption.SearchTopLevelOnly)
        If MsgBox(Application.Current.Resources("msg_resetlist"), MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then clearall = True
        LB_Pic.ItemsSource = Nothing
        Using tmpListOfPic = ListOfPic.Copy
            ListOfPic.Clear()
            For Each fd In MainWindow.folders_image
                If My.Computer.FileSystem.DirectoryExists(fd) Then
                    For Each f In My.Computer.FileSystem.GetFiles(fd, searchopt)
                        Dim filefullname = My.Computer.FileSystem.GetName(f)
                        Dim filename = IO.Path.GetFileNameWithoutExtension(filefullname)
                        Dim ext = IO.Path.GetExtension(filefullname)
                        If MainWindow.PicFormats.Contains(ext.ToLower) Then
                            Dim tmprow = ListOfPic.NewRow
                            tmprow("Path") = f
                            Dim tmpdate As Date
                            If Date.TryParse(filename, tmpdate) Then
                                tmprow("Date") = Date.Parse(filename).ToString
                            End If
                            If Not clearall Then
                                Dim match = tmpListOfPic.Rows.Find(f)
                                If match IsNot Nothing Then
                                    For i = 2 To ListOfPic.Columns.Count - 1
                                        tmprow(i) = match(i)
                                    Next
                                End If
                            End If
                            ListOfPic.Rows.Add(tmprow)
                        End If
                    Next
                End If
            Next
        End Using
        LB_Pic.ItemsSource = ListOfPic.Rows
    End Sub

    Private Sub LostFocuses(sender As Object, e As RoutedEventArgs) Handles TB_Text.LostFocus, CbB_FontFamily.LostFocus, Sld_FontSize.ValueChanged, CbB_FontColor.LostFocus, Sld_FontOffsetH.ValueChanged, Sld_FontOffsetV.ValueChanged
        If Me.IsLoaded AndAlso Not changingselection AndAlso LB_Pic.SelectedItem IsNot Nothing Then
            For Each selected As DataRow In LB_Pic.SelectedItems
                Dim col = sender.Name.ToString.Split("_"c)(1)
                Select Case sender.GetType
                    Case GetType(TextBox)
                        selected(col) = sender.Text
                    Case GetType(ComboBox)
                        If sender.IsEditable Then
                            selected(col) = sender.Text
                        Else
                            selected(col) = sender.SelectedValue
                        End If
                    Case GetType(Slider)
                        selected(col) = sender.Value
                    Case Else
                        Exit For
                End Select
            Next
            If LB_Pic.SelectedItems.Count = 1 Then
                Task.Run(Sub()
                             Dim selected As DataRow
                             Dispatcher.Invoke(Sub() selected = LB_Pic.SelectedItem)
                             RefreshPreview(selected)
                         End Sub)
            End If
        End If
    End Sub

    Private Sub CB_Title_Checked(sender As Object, e As RoutedEventArgs) Handles CB_Title.Checked, CB_Title.Unchecked
        If Me.IsLoaded Then
            For Each i In LB_Pic.SelectedItems
                ListOfPic.Rows.Find(i("Path"))("TitleSlide") = CB_Title.IsChecked
            Next
        End If
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        ListOfPic.Dispose()
        ListOfPic = Nothing
    End Sub
End Class
