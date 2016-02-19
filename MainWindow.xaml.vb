Imports System.Threading

Class MainWindow
    Public Shared ListOfPic As New System.Data.DataTable("ImageList")
    Dim PicFormats() As String = {".jpg", ".jpeg", ".bmp", ".png", ".tif", ".tiff"}
    Dim ListOfMusic As New List(Of String)
    Dim ran As New Random
    Dim anim_fadein As Animation.DoubleAnimation
    Dim anim_fadeout As Animation.DoubleAnimation
    Dim player As New System.Windows.Media.MediaPlayer
    Dim currentaudio As Integer = 0
    Dim playing As Boolean = False
    Dim audiofading As Boolean = False
    Dim w, h As Double
    Dim position As Integer = 0
    Dim m As Integer = 0, mm As Integer = 0
    Dim pic As Object
    Dim picmove_sec As UInteger = 7
    Dim moveon As Boolean = True
    Dim restart As Boolean = False
    Dim worker_pic As Thread
    Public Shared framerate As UInteger = 60
    Public Shared duration As UInteger = 7 'this only serves as a store. program will read duration value from picmove_sec.
    Public Shared folders_image As New List(Of String)
    Public Shared folders_music As New List(Of String)
    Public Shared verticalLock As Boolean = True
    Public Shared resolutionLock As Boolean = True
    Public Shared verticalOptimize As Boolean = True
    Public Shared horizontalOptimize As Boolean = True
    Public Shared fadeout As Boolean = True
    Public Shared verticalOptimizeR As Double = 0.4
    Public Shared horizontalOptimizeR As Double = 0.4
    Public Shared transit As String = "Ken Burns"
    Public Shared loadquality As Double = 1.2
    Public Shared ScaleMode_Dic As New Dictionary(Of Integer, String)
    Public Shared scalemode As Integer = 2
    Private Declare Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As EXECUTION_STATE) As EXECUTION_STATE
    Dim ExecState_Set As Boolean
    Private Enum EXECUTION_STATE As Integer
        ''' <summary>
        ''' Informs the system that the state being set should remain in effect until the next call that uses ES_CONTINUOUS and one of the other state flags is cleared.
        ''' </summary>
        ES_CONTINUOUS = &H80000000
        ''' <summary>
        ''' Forces the display to be on by resetting the display idle timer.
        ''' </summary>
        ES_DISPLAY_REQUIRED = &H2
        ''' <summary>
        ''' Forces the system to be in the working state by resetting the system idle timer.
        ''' </summary>
        ES_SYSTEM_REQUIRED = &H1
    End Enum

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        w = Me.Width
        h = Me.Height

        'initialize datatable
        ListOfPic.Columns.Add("Path", GetType(String))
        ListOfPic.Columns.Add("Date", GetType(String))
        ListOfPic.Columns.Add("CB_Title", GetType(Boolean)).DefaultValue = False
        ListOfPic.PrimaryKey = {ListOfPic.Columns("Path")}

        'loading other settings
        ScaleMode_Dic.Add(2, "(Default) High")
        ScaleMode_Dic.Add(3, "Medium")
        ScaleMode_Dic.Add(1, "Low")

        'loading config.xml
        Dim config As XElement
        Do
            Try
                config = XElement.Load("config.xml")
                Exit Do
            Catch
                If MsgBox("Error loading configuration. Would you like to open settings?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim optwin As New OptWindow
                    optwin.ShowDialog()
                    optwin.Close()
                Else
                    Me.Close()
                    Exit Sub
                End If
            End Try
        Loop

        'loading image list
        Do
            'loading ListOfPic
            ListOfPic.Clear()
            If config.Elements("DocumentElement").Any Then
                ListOfPic.ReadXml(New IO.StringReader(config.Element("DocumentElement").ToString))
            End If

            folders_image.Clear()
            For Each ele In config.Element("PicDir").Elements
                folders_image.Add(ele.Value)
                If My.Computer.FileSystem.DirectoryExists(ele.Value) Then
                    For Each f In My.Computer.FileSystem.GetFiles(ele.Value)
                        Dim filefullname = My.Computer.FileSystem.GetName(f)
                        Dim filename = IO.Path.GetFileNameWithoutExtension(filefullname)
                        Dim ext = IO.Path.GetExtension(filefullname)
                        If PicFormats.Contains(ext.ToLower) Then
                            Dim row = ListOfPic.Rows.Find(f)
                            If row IsNot Nothing Then
                                Dim tmpdate As Date
                                If DateTime.TryParse(filename, tmpdate) Then
                                    row("Date") = DateTime.Parse(filename).ToString
                                End If
                            Else
                                Dim tmprow = ListOfPic.NewRow
                                tmprow("Path") = f
                                Dim tmpdate As Date
                                If DateTime.TryParse(filename, tmpdate) Then
                                    tmprow("Date") = DateTime.Parse(filename).ToString
                                End If
                                ListOfPic.Rows.Add(tmprow)
                            End If

                        End If
                    Next
                End If
            Next
            'remove old paths
            Dim tmplst As New List(Of System.Data.DataRow)
            For Each row As System.Data.DataRow In ListOfPic.Rows
                If Not folders_image.Contains(IO.Path.GetDirectoryName(row("Path"))) Then
                    tmplst.Add(row)
                ElseIf Not My.Computer.FileSystem.FileExists(row("Path")) Then
                    tmplst.Add(row)
                End If
            Next
            For Each row In tmplst
                ListOfPic.Rows.Remove(row)
            Next

            If ListOfPic.Rows.Count = 0 Then
                If MsgBox("No pictures are found in specified folders. Would you like to open settings?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim optwin As New OptWindow
                    optwin.ShowDialog()
                    optwin.Close()
                    config = XElement.Load("config.xml")
                Else
                    Me.Close()
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        Loop

        'save changes
        Using lop_str = New IO.StringWriter()
            ListOfPic.WriteXml(lop_str)
            Dim lop = XElement.Parse(lop_str.ToString)
            config.Elements("DocumentElement").Remove()
            config.Add(lop)
            config.Save("config.xml")
        End Using

        'loading music list
        folders_music.Clear()
        For Each ele In config.Element("Music").Elements
            folders_music.Add(ele.Value)
            If My.Computer.FileSystem.DirectoryExists(ele.Value) Then
                For Each f In My.Computer.FileSystem.GetFiles(ele.Value)
                    ListOfMusic.Add(f)
                Next
            End If
        Next

        AddHandler player.MediaEnded, Sub()
                                          currentaudio += 1
                                          If currentaudio = ListOfMusic.Count Then currentaudio = 0
                                          player.Open(New Uri(ListOfMusic(currentaudio)))
                                          player.Play()
                                      End Sub

        If config.Elements("Framerate").Any Then framerate = config.Element("Framerate").Value
        If config.Elements("Duration").Any Then
            Dim tmp = Convert.ToUInt32(config.Element("Duration").Value)
            duration = tmp
            If tmp >= 5 Then picmove_sec = tmp
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
        If config.Elements("VerticalOptimizeRatio").Any Then verticalOptimizeR = config.Element("VerticalOptimizeRatio").Value
        If config.Elements("HorizontalOptimizeRatio").Any Then horizontalOptimizeR = config.Element("HorizontalOptimizeRatio").Value
        If config.Elements("Transit").Any Then transit = config.Element("Transit").Value
        If config.Elements("LoadQuality").Any Then loadquality = config.Element("LoadQuality").Value
        If config.Elements("ScaleMode").Any Then scalemode = config.Element("ScaleMode").Value

        tb_date0.FontSize = Me.Height / 12
        tb_date1.FontSize = Me.Height / 12

        'loading first picture
        LoadNextImg()

        Me.Background = Brushes.Black

        Select Case transit
            Case "Ken Burns"
                worker_pic = New Thread(AddressOf mainThrd_KBE)
            Case "Breath"
                worker_pic = New Thread(AddressOf mainThrd_Breath)
            Case Else
                MsgBox("Error reading transit value. Program will exit now.", MsgBoxStyle.Critical)
                Me.Close()
                Exit Sub
        End Select
        worker_pic.IsBackground = True
        worker_pic.Priority = ThreadPriority.Lowest 'this is not the UI thread
        worker_pic.Start()

        'disabling sleep / screensaver
        'this seems to be unnecessary on the dev PC as the sleep timers are ignored anyway without the following two lines.
        SetThreadExecutionState(EXECUTION_STATE.ES_SYSTEM_REQUIRED Or EXECUTION_STATE.ES_DISPLAY_REQUIRED Or EXECUTION_STATE.ES_CONTINUOUS)
        ExecState_Set = True
    End Sub

    Public Function RandomNum(min As UInteger, max As UInteger, neg As Boolean)
        If neg Then
            If ran.Next(2) = 0 Then
                Return -ran.Next(min, max)
            Else
                Return ran.Next(min, max)
            End If
        Else
            Return ran.Next(min, max)
        End If
    End Function

    Private Sub Anim_TitleSlide(tgt As Image)
        'title slide animation
        Thread.Sleep(2500) 'waiting for the former img to fade out
        Dispatcher.Invoke(
            Sub()
                tgt = New Image
                tgt.Name = "tgt"
                mainGrid.Children.Add(tgt)
                tgt.HorizontalAlignment = HorizontalAlignment.Stretch
                tgt.VerticalAlignment = VerticalAlignment.Stretch
                tgt.Source = pic
                tgt.BeginAnimation(Image.OpacityProperty, New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 2))))
            End Sub)
        Thread.Sleep(picmove_sec * 2000)
        Dispatcher.Invoke(Sub() tgt.BeginAnimation(Image.OpacityProperty, New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))))
        Thread.Sleep(2500)
    End Sub

    Private Sub textThrd_KBE(pos As Integer, ByRef output As Integer)
        If Not IsDBNull(ListOfPic.Rows(pos - 1)("Date")) Then
            Dim tgt_tb As TextBlock
            'determining future pics with same year and month
            Dim tmpdate = Date.Parse(ListOfPic.Rows(pos - 1)("Date")).ToString("yyyy.M")
            Dim tbmove_sec = picmove_sec
            For n = 1 To ListOfPic.Rows.Count - pos 'not running if it is the last pic as ListOfPic.Count-pos=0
                If Not IsDBNull(ListOfPic.Rows(pos - 1 + n)("Date")) Then
                    If Date.Parse(ListOfPic.Rows(pos - 1 + n)("Date")).ToString("yyyy.M") = tmpdate Then
                        tbmove_sec += (picmove_sec - 1)
                        output = pos + n
                        If output = ListOfPic.Rows.Count Then
                            output = ListOfPic.Rows.Count + 1
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
                                  tgt_tb.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadeout)
                              End Sub)
        Else
            output = pos + 1
        End If
    End Sub

    'Private Sub textThrd_KBE(force_restart As Boolean)
    '    Dispatcher.Invoke(
    '        Sub()
    '            Dim anim_tbfadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
    '            Animation.Timeline.SetDesiredFrameRate(anim_tbfadeout, framerate)
    '            If tb_date0 IsNot Nothing Then tb_date0.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadeout)
    '            If tb_date1 IsNot Nothing Then tb_date1.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadeout)
    '        End Sub)
    'End Sub

    Private Sub mainThrd_KBE()
        Do While audiofading
            Thread.Sleep(500)
        Loop
        Dispatcher.Invoke(
            Sub()
                If ListOfMusic.Count > 0 Then
                    player.Open(New Uri(ListOfMusic(0)))
                    player.Play()
                    playing = True
                End If

                anim_fadein = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1)))
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)
        Dim tgt_img As Image
        Dim delta As Double
        Dim tbchkpoint As Integer = 1

        Do
            If position >= tbchkpoint Then
                Task.Run(Sub() textThrd_KBE(position, tbchkpoint))
            End If

            If ListOfPic.Rows(position - 1)("CB_Title") Then
                Anim_TitleSlide(tgt_img)
            Else
                Dispatcher.Invoke(
                Sub()
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
                        If ran.Next(2) = 0 Then
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
                        If horizontalOptimize AndAlso delta > w * 0.7 Then
                            startpoint = delta * horizontalOptimizeR
                            If startpoint > tgt_img.Width - w Then
                                startpoint = tgt_img.Width - w
                            End If
                        Else
                            startpoint = 0
                        End If

                        'move left or right
                        If ran.Next(2) = 0 Then 'means 0<=ran<2
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Left
                            tgt_img.RenderTransformOrigin = New Point(0, 0.5)
                            anim_move = New Animation.ThicknessAnimation(New Thickness(-startpoint, 0, 0, 0), New Thickness(-delta, 0, 0, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        Else
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Right
                            tgt_img.RenderTransformOrigin = New Point(1, 0.5)
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, -startpoint, 0), New Thickness(0, 0, -delta, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        End If
                    Else
                        'height is the longer edge comparing to the size of the monitor
                        tgt_img.Width = w
                        tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
                        If ran.Next(2) = 0 Then
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
                        If verticalOptimize AndAlso delta > h * 0.7 Then
                            startpoint = delta * verticalOptimizeR
                            If startpoint > tgt_img.Height - h Then
                                startpoint = tgt_img.Height - h
                            End If
                        Else
                            startpoint = 0
                        End If

                        'move up or down or up only when long
                        If verticalLock AndAlso tgt_img.Height > h * 1.2 Then
                            'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                            tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                            tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'this and above line is to make transform align with bottom
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, -startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
                        Else
                            If ran.Next(2) = 0 Then
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Top
                                tgt_img.RenderTransformOrigin = New Point(0.5, 0) 'this and above line is to make transform align with top
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, -startpoint, 0, 0), New Thickness(0, -delta, 0, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                            Else
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                                tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, -startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
                            End If
                        End If
                    End If
                    tgt_img.Source = pic

                    Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomx, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomy, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoomx)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoomy)
                    tgt_img.BeginAnimation(Image.MarginProperty, anim_move) ', Animation.HandoffBehavior.Compose)
                    tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein) ', Animation.HandoffBehavior.Compose)
                End Sub)
                Thread.Sleep(1000)
            End If

            Dim tmpposition = position
            Dim loadtask = Task.Run(
                    Sub()
                        'Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            If Not ListOfPic.Rows(tmpposition - 1)("CB_Title") Then
                Thread.Sleep((picmove_sec - 2) * 1000)
                Dispatcher.Invoke(Sub() tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout))
            End If

            loadtask.Wait()

            If restart Then
                restart = False
                Exit Sub
            End If

            Do Until moveon
                Thread.Sleep(1000)
            Loop
        Loop
    End Sub

    Private Sub mainThrd_Breath()
        Do While audiofading
            Thread.Sleep(500)
        Loop

        Dim ease_in As Animation.CubicEase
        Dim ease_out As Animation.CubicEase
        Dim ease_inout As Animation.CubicEase
        Dispatcher.Invoke(
            Sub()
                If ListOfMusic.Count > 0 Then
                    player.Open(New Uri(ListOfMusic(0)))
                    player.Play()
                    playing = True
                End If
                ease_in = New Animation.CubicEase With {.EasingMode = Animation.EasingMode.EaseIn}
                ease_out = New Animation.CubicEase With {.EasingMode = Animation.EasingMode.EaseOut}
                ease_inout = New Animation.CubicEase With {.EasingMode = Animation.EasingMode.EaseInOut}
                anim_fadein = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1))) ' With {.EasingFunction = ease_out}
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1))) ' With {.EasingFunction = ease_in}
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)
        Dim tgt_img As Image
        Dim delta As Double
        Dim tbchkpoint As Integer = 1
        Dim last_zoom = False ', last_move As Boolean

        Do
            If position >= tbchkpoint Then
                Task.Run(Sub() textThrd_KBE(position, tbchkpoint))
            End If

            If ListOfPic.Rows(position - 1)("CB_Title") Then
                Anim_TitleSlide(tgt_img)
            Else
                Dispatcher.Invoke(
                Sub()
                    Dim tgt_trasform As ScaleTransform
                    Dim anim_move As Animation.ThicknessAnimation
                    Dim anim_zoomx As Animation.DoubleAnimationUsingKeyFrames
                    Dim anim_zoomy As Animation.DoubleAnimationUsingKeyFrames

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
                        If last_zoom Then ' ran.Next(2) = 0 Then
                            'zoom in
                            last_zoom = False
                            delta = tgt_img.Width * 1.1 - w
                            tgt_trasform = New ScaleTransform(1, 1)
                            tgt_img.RenderTransform = tgt_trasform
                            anim_zoomx = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomy = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
                        Else
                            'zoom out
                            last_zoom = True
                            delta = tgt_img.Width - w
                            tgt_trasform = New ScaleTransform(1.3, 1.3)
                            tgt_img.RenderTransform = tgt_trasform
                            anim_zoomx = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomy = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
                        End If

                        Dim startpoint As Double
                        If horizontalOptimize AndAlso delta > w * 0.7 Then
                            startpoint = delta * horizontalOptimizeR
                            If startpoint > tgt_img.Width - w Then
                                startpoint = tgt_img.Width - w
                            End If
                        Else
                            startpoint = 0
                        End If

                        If ran.Next(2) = 0 Then 'means 0<=ran<2
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Left
                            tgt_img.RenderTransformOrigin = New Point(0, 0.5) 'this and above line is to make transform align with left
                            anim_move = New Animation.ThicknessAnimation(New Thickness(-startpoint, 0, 0, 0), New Thickness(-delta, 0, 0, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                        Else
                            tgt_img.HorizontalAlignment = Windows.HorizontalAlignment.Right
                            tgt_img.RenderTransformOrigin = New Point(1, 0.5) 'this and above line is to make transform align with right
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, -startpoint, 0), New Thickness(0, 0, -delta, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                        End If
                        If tgt_img.Width > w * 1.5 Then
                            anim_move.EasingFunction = ease_inout
                        End If
                    Else
                        'height is the longer edge comparing to the size of the monitor
                        tgt_img.Width = w
                        tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
                        If last_zoom Then 'ran.Next(2) = 0 Then
                            'zoom in
                            last_zoom = False
                            delta = tgt_img.Height * 1.1 - h
                            tgt_trasform = New ScaleTransform(1, 1)
                            tgt_img.RenderTransform = tgt_trasform
                            anim_zoomx = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomy = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0.1, 0.4, 0.7, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0.1, 0.4, 0.7, 1)))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
                        Else
                            'zoom out
                            last_zoom = True
                            delta = tgt_img.Height - h
                            tgt_trasform = New ScaleTransform(1.3, 1.3)
                            tgt_img.RenderTransform = tgt_trasform
                            anim_zoomx = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomy = New Animation.DoubleAnimationUsingKeyFrames
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                            anim_zoomx.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            anim_zoomy.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                            'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
                            'anim_zoomy.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
                        End If
                        Dim startpoint As Double
                        If verticalOptimize AndAlso delta > h * 0.7 Then
                            startpoint = delta * verticalOptimizeR
                            If startpoint > tgt_img.Height - h Then
                                startpoint = tgt_img.Height - h
                            End If
                        Else
                            startpoint = 0
                        End If

                        'move up or down or up only when long
                        If verticalLock AndAlso tgt_img.Height > h * 1.2 Then
                            'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                            tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                            tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'this and above line is to make transform align with bottom
                            anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, -startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                            anim_move.EasingFunction = ease_inout
                        Else
                            If ran.Next(2) = 0 Then
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Top
                                tgt_img.RenderTransformOrigin = New Point(0.5, 0) 'this and above line is to make transform align with top
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, -startpoint, 0, 0), New Thickness(0, -delta, 0, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                            Else
                                tgt_img.VerticalAlignment = Windows.VerticalAlignment.Bottom
                                tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                                anim_move = New Animation.ThicknessAnimation(New Thickness(0, 0, 0, -startpoint), New Thickness(0, 0, 0, -delta), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                            End If
                            If tgt_img.Height > h * 1.5 Then
                                anim_move.EasingFunction = ease_inout
                            End If
                        End If
                    End If

                    tgt_img.Source = pic
                    Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomx, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_zoomy, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
                    Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoomx)
                    tgt_trasform.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoomy)
                    tgt_img.BeginAnimation(Image.MarginProperty, anim_move) ', Animation.HandoffBehavior.Compose)
                    tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein) ', Animation.HandoffBehavior.Compose)
                End Sub)
                Thread.Sleep(1000)
            End If

            Dim tmpposition = position
            Dim loadtask = Task.Run(
                    Sub()
                        'Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            If Not ListOfPic.Rows(tmpposition - 1)("CB_Title") Then
                Thread.Sleep((picmove_sec - 2.5) * 1000)
                Dispatcher.Invoke(Sub() tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout))
            End If

            Thread.Sleep(500)
            loadtask.Wait()

            If restart Then
                restart = False
                Exit Sub
            End If

            Do Until moveon
                Thread.Sleep(1000)
            Loop
        Loop
    End Sub

    Private Sub LoadNextImg()
        Dispatcher.Invoke(Sub()
                              RenderOptions.SetBitmapScalingMode(slide_img0, scalemode)
                              RenderOptions.SetBitmapScalingMode(slide_img1, scalemode)
                          End Sub)
        Try
            Dim s As Size
            Dim imgpath As String = ListOfPic.Rows(position)("Path")
            Using strm = New IO.FileStream(imgpath, IO.FileMode.Open, IO.FileAccess.Read)
                Dim frame = BitmapFrame.Create(strm, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)
                s = New Size(frame.PixelWidth, frame.PixelHeight)
            End Using

            pic = New BitmapImage
            pic.BeginInit()
            If resolutionLock Then
                If s.Width > s.Height Then
                    If s.Height > h * loadquality Then
                        pic.DecodePixelHeight = h * loadquality
                    End If
                Else
                    If s.Width > w * loadquality Then
                        pic.DecodePixelWidth = w * loadquality
                    End If
                End If
            End If
            pic.CacheOption = BitmapCacheOption.OnLoad
            Using strm = New IO.FileStream(imgpath, IO.FileMode.Open, IO.FileAccess.Read)
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
        Catch
            pic = BitmapSource.Create(64, 64, 96, 96, PixelFormats.Indexed1, BitmapPalettes.BlackAndWhite, New Byte(64 * 8) {}, 8)
        Finally
            position += 1
            pic.Freeze()
        End Try
    End Sub

    Private Sub Window_PreviewKeyUp(sender As Object, e As KeyEventArgs)
        If e.Key = Key.F12 Then
            Dim optwin As New OptWindow
            optwin.ShowDialog()
            optwin.Close()
        ElseIf e.Key = Key.F11 Then
            Dim editwin As New EditWindow
            editwin.ShowDialog()
            editwin.Close()
        ElseIf e.Key = Key.P Then
            If Keyboard.Modifiers = ModifierKeys.Control Then 'pause image
                If moveon = True Then
                    moveon = False
                Else
                    moveon = True
                End If
            ElseIf Keyboard.Modifiers = ModifierKeys.Shift Then 'fadeout audio only
                If Not audiofading Then
                    If playing Then
                        Task.Run(AddressOf FadeoutAudio)
                    Else
                        Task.Run(AddressOf FadeinAudio)
                    End If
                End If
            End If
        ElseIf e.Key = Key.R AndAlso Keyboard.Modifiers = ModifierKeys.Control Then
            Task.Run(AddressOf FadeoutAudio)
            Task.Run(Sub()
                         restart = True
                         Dim black As Rectangle
                         Dispatcher.Invoke(Sub()
                                               black = New Rectangle
                                               mainGrid.Children.Add(black)
                                               Panel.SetZIndex(black, 9)
                                               black.Fill = Windows.Media.Brushes.Black
                                               black.Margin = New Thickness(0)
                                               black.BeginAnimation(OpacityProperty, anim_fadein)
                                           End Sub)
                         Thread.Sleep(1000)
                         Dispatcher.Invoke(Sub()
                                               For Each child In mainGrid.Children
                                                   If child.Name.StartsWith("slide_img") OrElse child.Name = "tgt" Then
                                                       CType(child, Image).Source = Nothing
                                                   ElseIf child.Name.StartsWith("tb_date") Then
                                                       CType(child, TextBlock).Text = ""
                                                   End If
                                               Next
                                               mainGrid.Children.Remove(black)
                                               black = Nothing
                                           End Sub)
                         worker_pic.Join()

                         'restart thread
                         position = 0
                         LoadNextImg()
                         Select Case transit
                             Case "Ken Burns"
                                 worker_pic = New Thread(AddressOf mainThrd_KBE)
                             Case "Breath"
                                 worker_pic = New Thread(AddressOf mainThrd_Breath)
                             Case Else
                                 MsgBox("Error reading transit value. Program will exit now.", MsgBoxStyle.Critical)
                                 Me.Close()
                                 Exit Sub
                         End Select
                         worker_pic.IsBackground = True
                         worker_pic.Priority = ThreadPriority.Lowest
                         worker_pic.Start()
                     End Sub)
        End If
    End Sub

    Private Sub FadeoutAudio()
        audiofading = True
        Dim vol As Double
        Do
            Dispatcher.Invoke(Sub()
                                  player.Volume -= 0.01
                                  vol = player.Volume
                              End Sub)
            Thread.Sleep(60)
        Loop Until vol < 0.01
        Dispatcher.Invoke(Sub()
                              player.Pause()
                              player.Volume = 0.5
                          End Sub)
        playing = False
        audiofading = False
    End Sub

    Private Sub FadeinAudio()
        audiofading = True
        Dispatcher.Invoke(Sub()
                              player.Volume = 0
                              player.Play()
                          End Sub)
        Dim vol As Double = 0
        Do
            Dispatcher.Invoke(Sub()
                                  player.Volume += 0.01
                                  vol = player.Volume
                              End Sub)
            Thread.Sleep(60)
        Loop Until vol > 0.49
        'Dispatcher.Invoke(Sub() player.Volume = 0.5)
        playing = True
        audiofading = False
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        If ExecState_Set Then SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS)
    End Sub
End Class
