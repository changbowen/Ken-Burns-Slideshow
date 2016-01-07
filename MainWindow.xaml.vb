Imports System.Threading

Class MainWindow
    Dim ListOfPic As New Dictionary(Of String, String)
    Dim PicFormats() As String = {".jpg", ".jpeg", ".bmp", ".png", ".tif", ".tiff"}
    Dim ListOfMusic As New List(Of String)
    Dim framerate As Integer = 60
    Dim ran As New Random
    Dim anim_fadein As New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1)))
    Dim anim_fadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
    Dim player As New System.Windows.Media.MediaPlayer
    Dim playing As Integer = 0
    Dim w, h As Double
    Dim position As Integer = 0
    Dim picmove_sec As Integer = 5
    Dim m As Integer = 0, mm As Integer = 0
    Dim pic As Object
    Public Shared verticalLock As Boolean = True
    Public Shared resolutionLock As Boolean = True
    Public Shared verticalOptimize As Boolean = True
    Public Shared horizontalOptimize As Boolean = True
    Public Shared fadeout As Boolean = True

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Me.Background = Brushes.Black
        Me.Topmost = False
        w = Me.Width
        h = Me.Height

        Dim config As XElement
        If My.Computer.FileSystem.FileExists("config.xml") Then
            config = XElement.Load("config.xml")
        Else
            MsgBox("Config.xml file is not found at application root.", MsgBoxStyle.Exclamation)
            Me.Close()
            Exit Sub
        End If

        framerate = config.Element("Framerate").Value
        Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)

        For Each ele In config.Element("PicDir").Elements
            If My.Computer.FileSystem.DirectoryExists(ele.Value) Then
                For Each f In My.Computer.FileSystem.GetFiles(ele.Value)
                    Dim filefullname = My.Computer.FileSystem.GetName(f)
                    Dim filename = IO.Path.GetFileNameWithoutExtension(filefullname)
                    Dim ext = IO.Path.GetExtension(filefullname)
                    If PicFormats.Contains(ext.ToLower) Then
                        Dim tmpdate As Date
                        If DateTime.TryParse(filename, tmpdate) Then
                            ListOfPic.Add(DateTime.Parse(filename).ToString, f)
                        Else
                            ListOfPic.Add("NON-DATE_" & filename, f)
                        End If
                    End If
                Next
            End If
        Next
        If ListOfPic.Count = 0 Then
            MsgBox("No pictures are found in specified folders. Please check the config file.", MsgBoxStyle.Exclamation)
            Me.Close()
            Exit Sub
        End If

        For Each ele In config.Element("Music").Elements
            If My.Computer.FileSystem.DirectoryExists(ele.Value) Then
                For Each f In My.Computer.FileSystem.GetFiles(ele.Value)
                    ListOfMusic.Add(f)
                Next
            End If
        Next
        If ListOfMusic.Count > 0 Then
            player.Open(New Uri(ListOfMusic(0)))
            player.Play()
            AddHandler player.MediaEnded, Sub()
                                              playing += 1
                                              If playing = ListOfMusic.Count Then playing = 0
                                              player.Open(New Uri(ListOfMusic(playing)))
                                              player.Play()
                                          End Sub
        End If

        If config.Elements("Duration").Any Then
            Dim tmp = CInt(config.Element("Duration").Value)
            If tmp >= 4 Then picmove_sec = tmp
        End If
        If config.Elements("VerticalLock").Any AndAlso config.Element("VerticalLock").Value.ToLower = "false" Then
            verticalLock = False
        End If
        If config.Elements("ResolutionLock").Any AndAlso config.Element("ResolutionLock").Value.ToLower = "false" Then
            resolutionLock = False
        End If
        If config.Elements("VerticalOptimize").Any AndAlso config.Element("VerticalOptimize").Value.ToLower = "false" Then
            verticalOptimize = False
        End If
        If config.Elements("HorizontalOptimize").Any AndAlso config.Element("HorizontalOptimize").Value.ToLower = "false" Then
            horizontalOptimize = False
        End If
        If config.Elements("Fadeout").Any AndAlso config.Element("Fadeout").Value.ToLower = "false" Then
            fadeout = False
        End If

        tb_date0.FontSize = Me.Height / 12
        tb_date1.FontSize = Me.Height / 12

        'loading first picture
        LoadNextImg()
        position += 1
        'CType(mainGrid.FindName("slide_img0"), Image).Source = pic
        'CType(mainGrid.FindName("slide_img1"), Image).Source = pic

        Dim worker_pic As New Thread(AddressOf mainThrd)
        worker_pic.IsBackground = True
        worker_pic.Priority = ThreadPriority.BelowNormal
        worker_pic.Start()
    End Sub

    Public Function RandomNum(min As UInteger, max As UInteger, neg As Boolean)
        If neg Then
            If ran.Next(0, 2) = 0 Then
                Return -ran.Next(min, max)
            Else
                Return ran.Next(min, max)
            End If
        Else
            Return ran.Next(min, max)
        End If
    End Function

    Private Sub textThrd(pos As Integer, ByRef output As Integer)
        If Not ListOfPic.Keys(pos - 1).StartsWith("NON-DATE_") Then
            Dim tgt_tb As TextBlock
            'determining future pics with same year and month
            Dim tmpdate = Date.Parse(ListOfPic.Keys(pos - 1)).ToString("yyyy.M")
            Dim tbmove_sec = picmove_sec
            For n = 1 To ListOfPic.Count - pos 'not running if it is the last pic as ListOfPic.Count-pos=0
                If Not ListOfPic.Keys(pos - 1 + n).StartsWith("NON-DATE_") Then
                    If Date.Parse(ListOfPic.Keys(pos - 1 + n)).ToString("yyyy.M") = tmpdate Then
                        tbmove_sec += (picmove_sec - 1)
                        output = pos + n
                        If output = ListOfPic.Count Then
                            output = ListOfPic.Count + 1
                        End If
                    Else
                        output = pos + n
                        Exit For
                    End If
                Else
                    output = pos + n + 1
                    Exit For
                End If
            Next

            If mm = 0 Then
                mm = 1
            Else
                mm = 0
            End If
            Dispatcher.Invoke(Sub()
                                  tgt_tb = CType(mainGrid.FindName("tb_date" & mm), TextBlock)
                                  tgt_tb.Text = tmpdate
                                  Dim anim_tbfadein As New Animation.DoubleAnimation(0, 0.7, New Duration(New TimeSpan(0, 0, 2)))
                                  Dim anim_tbmove As New Animation.ThicknessAnimation(New Thickness(w / 5, h * 0.75, 0, 0), New Thickness(w / 5 + RandomNum(h / 32, h / 25, True), h * 0.75 + RandomNum(h / 32, h / 25, True), 0, 0), New Duration(New TimeSpan(0, 0, tbmove_sec)))
                                  'Dim anim_tbzoom = New Animation.DoubleAnimation(tgt_tb.FontSize * 1.2, New Duration(New TimeSpan(0, 0, tbmove_sec)))
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbfadein, framerate)
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbmove, framerate)
                                  'Animation.Timeline.SetDesiredFrameRate(anim_tbzoom, framerate)

                                  'tgt_tb.BeginAnimation(TextBlock.FontSizeProperty, anim_tbzoom)
                                  tgt_tb.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadein)
                                  tgt_tb.BeginAnimation(TextBlock.MarginProperty, anim_tbmove)
                              End Sub)

            Thread.Sleep((tbmove_sec - 1) * 1000)
            Dispatcher.Invoke(Sub()
                                  Dim anim_tbfadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbfadeout, framerate)
                                  tgt_tb.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadeout, Animation.HandoffBehavior.Compose)
                              End Sub)
        Else
            output = pos + 1
        End If
    End Sub

    Private Sub mainThrd()
        Dim tgt_img As Image
        Dim delta As Double
        Dim tbchkpoint As Integer = 1

        Do
            If position >= tbchkpoint Then
                Task.Run(Sub() textThrd(position, tbchkpoint))
            End If

            Dispatcher.Invoke(
                Sub()
                    Panel.SetZIndex(tb_date0, 2)
                    Panel.SetZIndex(tb_date1, 3)
                    Dim tgt_trasform As ScaleTransform
                    Dim anim_move As Animation.ThicknessAnimation
                    Dim anim_zoomx As Animation.DoubleAnimation
                    Dim anim_zoomy As Animation.DoubleAnimation

                    'switch target
                    If m = 0 Then
                        m = 1
                        tgt_img = slide_img1
                        Panel.SetZIndex(slide_img0, 0)
                        Panel.SetZIndex(slide_img1, 1)
                    Else
                        m = 0
                        tgt_img = slide_img0
                        Panel.SetZIndex(slide_img1, 0)
                        Panel.SetZIndex(slide_img0, 1)
                    End If
                    If pic.PixelWidth / pic.PixelHeight > w / h Then
                        'width is the longer edge comparing to the size of the monitor
                        tgt_img.Height = h
                        tgt_img.Width = tgt_img.Height * pic.PixelWidth / pic.PixelHeight
                        If ran.Next(0, 2) = 0 Then
                            'zoom in
                            delta = tgt_img.Width * 1.2 - w
                            anim_zoomx = New Animation.DoubleAnimation(1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            anim_zoomy = New Animation.DoubleAnimation(1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            tgt_trasform = New ScaleTransform(1, 1)
                            tgt_img.RenderTransform = tgt_trasform
                        Else
                            'zoom out
                            delta = tgt_img.Width - w
                            anim_zoomx = New Animation.DoubleAnimation(1, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            anim_zoomy = New Animation.DoubleAnimation(1, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            tgt_trasform = New ScaleTransform(1.2, 1.2)
                            tgt_img.RenderTransform = tgt_trasform
                        End If

                        Dim startpoint As Double
                        If horizontalOptimize AndAlso delta > w / 2 Then
                            startpoint = -delta / 2
                        Else
                            startpoint = 0
                        End If
                        If ran.Next(0, 2) = 0 Then 'means 0<=ran<2
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Left
                            tgt_img.RenderTransformOrigin = New Point(0, 0.5)
                            anim_move = New Animation.ThicknessAnimation(New Thickness(startpoint, 0, 0, 0), New Thickness(-delta, 0, 0, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        Else
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Right
                            tgt_img.RenderTransformOrigin = New Point(1, 0.5)
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, startpoint, 0), New Thickness(0, 0, -delta, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        End If
                    Else
                        'height is the longer edge comparing to the size of the monitor
                        tgt_img.Width = w
                        tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
                        If ran.Next(0, 2) = 0 Then
                            'zoom in
                            delta = tgt_img.Height * 1.2 - h
                            anim_zoomx = New Animation.DoubleAnimation(1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            anim_zoomy = New Animation.DoubleAnimation(1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            tgt_trasform = New ScaleTransform(1, 1)
                            tgt_img.RenderTransform = tgt_trasform
                        Else
                            'zoom out
                            delta = tgt_img.Height - h
                            anim_zoomx = New Animation.DoubleAnimation(1, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            anim_zoomy = New Animation.DoubleAnimation(1, New Duration(New TimeSpan(0, 0, picmove_sec)))
                            tgt_trasform = New ScaleTransform(1.2, 1.2)
                            tgt_img.RenderTransform = tgt_trasform
                        End If

                        Dim startpoint As Double
                        If verticalOptimize AndAlso delta > h / 2 Then
                            startpoint = -delta / 2
                        Else
                            startpoint = 0
                        End If
                        If verticalLock Then
                            'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                            tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                            tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'this and above line is to make transform align with bottom
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        Else
                            If ran.Next(0, 2) = 0 Then
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Top
                                tgt_img.RenderTransformOrigin = New Point(0.5, 0) 'this and above line is to make transform align with top
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, startpoint, 0, 0), New Thickness(0, -delta, 0, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                            Else
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                                tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
                            End If
                        End If
                    End If
                    tgt_img.Source = pic

                    Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomx, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomy, framerate)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoomx)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoomy)
                    tgt_img.BeginAnimation(Image.MarginProperty, anim_move, Animation.HandoffBehavior.Compose)
                    tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein, Animation.HandoffBehavior.Compose)
                End Sub)
            Thread.Sleep(1000)

            Dim loadtask = Task.Run(
                    Sub()
                        'Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        Try
                            LoadNextImg()
                        Catch
                            pic = BitmapSource.Create(64, 64, 96, 96, PixelFormats.Indexed1, BitmapPalettes.BlackAndWhite, New Byte(64 * 8) {}, 8)
                        Finally
                            position += 1
                            pic.Freeze()
                        End Try
                    End Sub)

            Thread.Sleep((picmove_sec - 2) * 1000)
            If fadeout Then
                Dispatcher.Invoke(Sub()
                                      tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout, Animation.HandoffBehavior.Compose)
                                  End Sub)
            End If
            loadtask.Wait()
        Loop
    End Sub

    Private Sub LoadNextImg()
        Dim s As Size
        Using strm = New IO.FileStream(ListOfPic.Values(position), IO.FileMode.Open, IO.FileAccess.Read)
            Dim frame = BitmapFrame.Create(strm, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)
            s = New Size(frame.PixelWidth, frame.PixelHeight)
        End Using

        pic = New BitmapImage
        pic.BeginInit()
        If resolutionLock Then
            If s.Width > s.Height Then
                If s.Height > h * 1.2 Then
                    pic.DecodePixelHeight = h * 1.2
                End If
            Else
                If s.Width > w * 1.2 Then
                    pic.DecodePixelWidth = w * 1.2
                End If
            End If
        End If
        pic.CacheOption = BitmapCacheOption.OnLoad
        Using strm = New IO.FileStream(ListOfPic.Values(position), IO.FileMode.Open, IO.FileAccess.Read)
            pic.StreamSource = strm
            pic.EndInit()
        End Using

        'reading next picture gentaly proved to be futile
        'pic.StreamSource = New IO.FileStream(ListOfPic.Values(0), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
        'pic.EndInit()
        'Using stream = New IO.FileStream(ListOfPic.Values(position), IO.FileMode.Open, IO.FileAccess.Read)
        '    Dim blocksize As Long = Math.Ceiling(stream.Length / 200)
        '    Dim content(stream.Length - 1) As Byte
        '    Do
        '        If stream.Length - stream.Position < blocksize Then
        '            stream.Read(content, stream.Position, stream.Length - stream.Position)
        '        Else
        '            stream.Read(content, stream.Position, blocksize)
        '        End If
        '        Thread.Sleep(10)
        '    Loop Until stream.Position = stream.Length
        '    stream.Position = 0
        '    pic.StreamSource = stream
        '    pic.EndInit()
        'End Using
        pic.Freeze()
    End Sub

    Private Sub Window_PreviewKeyUp(sender As Object, e As KeyEventArgs)
        If e.Key = Key.F12 Then
            Dim optwin As New OptWindow
            optwin.ShowDialog()
            optwin.Close()
        End If
    End Sub
End Class
