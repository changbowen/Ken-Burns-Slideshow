Imports System.Threading

Class MainWindow
    Public Shared ListOfPic As New System.Data.DataTable("ImageList")
    Public Shared PicFormats() As String = {".jpg", ".jpeg", ".bmp", ".png", ".tif", ".tiff"}
    Dim ListOfMusic As New List(Of String)
    Dim ran As New Random
    Dim player As New System.Windows.Media.MediaPlayer
    Dim currentaudio As Integer = 0
    Dim playing As Boolean = False
    Dim audiofading As Boolean = False
    Dim w, h As Double
    Dim position As Integer = 0
    Dim m As Integer = 0, mm As Integer = 0
    Dim pic As BitmapImage
    Dim pics() As BitmapImage
    Dim picmove_sec As UInteger = 7
    Dim moveon As Boolean = True
    Dim aborting As Boolean = False
    Dim worker_pic As Thread
    Dim ctrlwindow As Window
    Public Shared framerate As UInteger = 60
    Public Shared duration As UInteger = 7 'only serves as a store. program will read duration value from picmove_sec.
    Public Shared folders_image As New List(Of String)
    Public Shared folders_music As New List(Of String)
    Public Shared verticalLock As Boolean = True
    Public Shared resolutionLock As Boolean = True
    Public Shared verticalOptimize As Boolean = True
    Public Shared horizontalOptimize As Boolean = True
    Public Shared fadeout As Boolean = True
    Public Shared verticalOptimizeR As Double = 0.6
    Public Shared horizontalOptimizeR As Double = 0.6
    Public Shared transit As Integer = 0
    Public Shared loadquality As Double = 1.2
    Public Shared loadmode_next As Integer = 0 'only serves as a store. program will read duration value from loadmode.
    Dim loadmode As Integer = 0
    Public Shared ScaleMode_Dic As New Dictionary(Of Integer, String)
    Public Shared scalemode As Integer = 2
    Public Shared blurmode As Integer = 0
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
        Process.GetCurrentProcess.PriorityClass = ProcessPriorityClass.High

        'initialize datatable
        ListOfPic.Columns.Add("Path", GetType(String))
        ListOfPic.Columns.Add("Date", GetType(String))
        ListOfPic.Columns.Add("CB_Title", GetType(Boolean)).DefaultValue = False
        ListOfPic.PrimaryKey = {ListOfPic.Columns("Path")}

        'loading other settings
        ScaleMode_Dic.Add(2, Application.Current.Resources("(default) high"))
        ScaleMode_Dic.Add(3, Application.Current.Resources("medium"))
        ScaleMode_Dic.Add(1, Application.Current.Resources("low"))

        'loading config.xml
        Dim config As XElement
        Do
            Try
                config = XElement.Load("config.xml")
                Exit Do
            Catch
                If MsgBox(Application.Current.Resources("msg_loadcfgerr"), MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim optwin As New OptWindow
                    optwin.ShowDialog()
                    optwin.Close()
                Else
                    Me.Close()
                    Exit Sub
                End If
            End Try
        Loop

        AddHandler player.MediaEnded, Sub()
                                          currentaudio += 1
                                          If currentaudio = ListOfMusic.Count Then currentaudio = 0
                                          player.Open(New Uri(ListOfMusic(currentaudio)))
                                          player.Play()
                                      End Sub

        'loading settings
        Do
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
            If config.Elements("BlurMode").Any Then blurmode = config.Elements("BlurMode").Value
            If config.Elements("LoadMode").Any Then
                loadmode_next = config.Elements("LoadMode").Value
                loadmode = loadmode_next
            End If

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
                If MsgBox(Application.Current.Resources("msg_noimgerr"), MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
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

        tb_date0.FontSize = Me.Height / 12
        tb_date1.FontSize = Me.Height / 12
        Me.Background = Brushes.Black

        Select Case transit
            Case 0 'Ken Burns
                worker_pic = New Thread(AddressOf mainThrd_KBE)
            Case 1 'Breath
                worker_pic = New Thread(AddressOf mainThrd_Breath)
            Case 2 'Throw
                worker_pic = New Thread(AddressOf mainThrd_Throw)
            Case 3 'Random
                worker_pic = New Thread(AddressOf mainThrd_Mix)
            Case Else
                MsgBox(Application.Current.Resources("msg_transerr"), MsgBoxStyle.Critical)
                Me.Close()
                Exit Sub
        End Select
        worker_pic.IsBackground = True
        worker_pic.Priority = ThreadPriority.Lowest 'this is not the UI thread

        If loadmode = 1 Then
            Task.Run(
                Sub()
                    Dim count = ListOfPic.Rows.Count
                    ReDim pics(count - 1)
                    Dim tb As TextBlock
                    Dispatcher.Invoke(Sub()
                                          tb = New TextBlock With {.Text = "Loading images..."}
                                          tb.FontFamily = New FontFamily("Segoe UI")
                                          tb.FontSize = 16
                                          tb.Foreground = Brushes.WhiteSmoke
                                          tb.Margin = New Thickness(w / 2 - 60, h / 2 - 8, 0, 0)
                                          mainGrid.Children.Add(tb)
                                          Panel.SetZIndex(tb, 10)
                                      End Sub)

                    For i = 0 To count - 1
                        If count > 1 Then
                            Dim ii = i
                            Dispatcher.Invoke(Sub() tb.Text = "Loading images... " & ii * 100 \ (count - 1) & "%")
                        End If

                        Dim img As New BitmapImage
                        Try
                            Dim imgpath As String = ListOfPic.Rows(i)("Path")
                            Using ms = New IO.FileStream(imgpath, IO.FileMode.Open, IO.FileAccess.Read)
                                Dim frame = BitmapFrame.Create(ms, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)
                                Dim s As Size = New Size(frame.PixelWidth, frame.PixelHeight)
                                ms.Position = 0
                                img.BeginInit()
                                If resolutionLock Then
                                    If s.Width > s.Height Then
                                        If s.Height > h * loadquality Then
                                            img.DecodePixelHeight = h * loadquality
                                        End If
                                    Else
                                        If s.Width > w * loadquality Then
                                            img.DecodePixelWidth = w * loadquality
                                        End If
                                    End If
                                End If
                                img.CacheOption = BitmapCacheOption.OnLoad
                                img.StreamSource = ms
                                img.EndInit()
                            End Using
                        Catch
                            img = New BitmapImage
                            Dim encoder As New BmpBitmapEncoder
                            Dim bmpsource = BitmapSource.Create(64, 64, 96, 96, PixelFormats.Indexed1, BitmapPalettes.BlackAndWhite, New Byte(64 * 8) {}, 8)
                            encoder.Frames.Add(BitmapFrame.Create(bmpsource))
                            Using ms As New IO.MemoryStream
                                encoder.Save(ms)
                                ms.Position = 0
                                img.BeginInit()
                                img.StreamSource = ms
                                img.EndInit()
                            End Using
                        Finally
                            img.Freeze()
                        End Try
                        pics(i) = img

                        Dim imgctrl As Image
                        Dispatcher.Invoke(Sub()
                                              imgctrl = New Image
                                              imgctrl.Source = img
                                              RenderOptions.SetBitmapScalingMode(imgctrl, scalemode)
                                              If img.PixelWidth / img.PixelHeight > w / h Then
                                                  imgctrl.Height = h
                                                  imgctrl.Width = h * img.PixelWidth / img.PixelHeight
                                              Else
                                                  imgctrl.Width = w
                                                  imgctrl.Height = w / img.PixelWidth * img.PixelHeight
                                              End If
                                              mainGrid.Children.Add(imgctrl)
                                              imgctrl.Opacity = 0.01
                                          End Sub)
                        Dim scales() As Double
                        'This is like a hack to achieve smooth transit animation. By setting opacity to 0.01 the image is basically 
                        'invisible but it is still drawn to the screen. The below select case block is to force WPF to draw once the 
                        'images at the scales that will be used by each transit animation. Numerous tests indicates that WPF seems 
                        'to cache different ScaleTransform results for further use. Sort of like a background mipmapping.
                        Select Case transit
                            Case 0
                                scales = {1, 1.2}
                            Case 1
                                scales = {1, 1.3}
                            Case 2
                                scales = {0.7, 1}
                            Case 3
                                scales = {0.7, 1, 1.3}
                            Case Else
                                scales = {}
                        End Select
                        For Each d In scales
                            Dispatcher.Invoke(Sub() imgctrl.RenderTransform = New ScaleTransform(d, d))
                            Thread.Sleep(250)
                        Next
                        Dispatcher.Invoke(Sub() mainGrid.Children.Remove(imgctrl))
                    Next
                    Dispatcher.Invoke(Sub() mainGrid.Children.Remove(tb))
                End Sub).ContinueWith(Sub()
                                          LoadNextImg()
                                          worker_pic.Start()
                                          Dispatcher.Invoke(Sub()
                                                                If ctrlwindow Is Nothing Then
                                                                    ctrlwindow = New ControlWindow
                                                                    ctrlwindow.Owner = Me
                                                                End If
                                                                ctrlwindow.Show()
                                                            End Sub)
                                      End Sub)
        ElseIf loadmode = 0 Then
            LoadNextImg()
            worker_pic.Start()
            If ctrlwindow Is Nothing Then
                ctrlwindow = New ControlWindow
                ctrlwindow.Owner = Me
            End If
            ctrlwindow.Show()
        Else
            Me.Close()
        End If


        'disabling sleep / screensaver
        'this seems to be unnecessary on the dev PC as the sleep timers are ignored anyway without the following two lines.
        SetThreadExecutionState(EXECUTION_STATE.ES_SYSTEM_REQUIRED Or EXECUTION_STATE.ES_DISPLAY_REQUIRED Or EXECUTION_STATE.ES_CONTINUOUS)
        ExecState_Set = True
    End Sub

    Private Function RandomNum(min As UInteger, max As UInteger, neg As Boolean)
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
                tgt.Source = pic
                tgt.Width = w
                tgt.Height = h
                mainGrid.Children.Add(tgt)
                tgt.BeginAnimation(Image.OpacityProperty, New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 2))))
            End Sub)
        Thread.Sleep(picmove_sec * 2000)
        Dispatcher.Invoke(Sub() tgt.BeginAnimation(Image.OpacityProperty, New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))))
        Thread.Sleep(1500)
        Dispatcher.Invoke(Sub() mainGrid.Children.Remove(tgt))
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
                                  Dim anim_tbmovex As New Animation.DoubleAnimation(w / 5, w / 5 + RandomNum(h / 32, h / 25, True), New Duration(New TimeSpan(0, 0, tbmove_sec)))
                                  Dim anim_tbmovey As New Animation.DoubleAnimation(h * 0.75, h * 0.75 + RandomNum(h / 32, h / 25, True), New Duration(New TimeSpan(0, 0, tbmove_sec)))
                                  'Dim anim_tbmove As New Animation.ThicknessAnimation(New Thickness(w / 5, h * 0.75, 0, 0), New Thickness(w / 5 + RandomNum(h / 32, h / 25, True), h * 0.75 + RandomNum(h / 32, h / 25, True), 0, 0), New Duration(New TimeSpan(0, 0, tbmove_sec)))
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbfadein, framerate)
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbmovex, framerate)
                                  Animation.Timeline.SetDesiredFrameRate(anim_tbmovey, framerate)
                                  Dim trans_trans As New TranslateTransform(w / 5, h * 0.75)
                                  tgt_tb.RenderTransform = trans_trans

                                  tgt_tb.BeginAnimation(TextBlock.OpacityProperty, anim_tbfadein)
                                  trans_trans.BeginAnimation(TranslateTransform.XProperty, anim_tbmovex)
                                  trans_trans.BeginAnimation(TranslateTransform.YProperty, anim_tbmovey)
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

        Dim ease_in, ease_out, ease_inout As Animation.CubicEase
        Dim anim_fadein, anim_fadeout As Animation.DoubleAnimation
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
                anim_fadein = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1)))
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)
        Dim tgt_img As Image
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
                    SwitchTarget(tgt_img)
                    Anim_KBE(tgt_img, ease_in, ease_out, ease_inout, anim_fadein)
                End Sub)
                Thread.Sleep(1000)
            End If

            Dim tmpposition = position
            Dim loadtask = Task.Run(
                    Sub()
                        Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            If Not ListOfPic.Rows(tmpposition - 1)("CB_Title") Then
                Thread.Sleep((picmove_sec - 2) * 1000)
                Dispatcher.Invoke(Sub()
                                      Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)
                                      tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout)
                                  End Sub)
            End If

            loadtask.Wait()

            If aborting Then
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
        Dim ease_in, ease_out, ease_inout As Animation.CubicEase
        Dim anim_fadein, anim_fadeout As Animation.DoubleAnimation
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
                anim_fadein = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1)))
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)
        Dim tgt_img As Image
        Dim tbchkpoint As Integer = 1
        Dim last_zoom = False

        Do
            If position >= tbchkpoint Then
                Task.Run(Sub() textThrd_KBE(position, tbchkpoint))
            End If

            If ListOfPic.Rows(position - 1)("CB_Title") Then
                Anim_TitleSlide(tgt_img)
            Else
                Dispatcher.Invoke(
                Sub()
                    SwitchTarget(tgt_img)
                    Anim_Breath(tgt_img, ease_in, ease_out, ease_inout, anim_fadein, last_zoom)
                End Sub)
                Thread.Sleep(1000)
            End If

            Dim tmpposition = position
            Dim loadtask = Task.Run(
                    Sub()
                        Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            If Not ListOfPic.Rows(tmpposition - 1)("CB_Title") Then
                Thread.Sleep((picmove_sec - 2.5) * 1000)
                Dispatcher.Invoke(Sub()
                                      Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)
                                      tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout)
                                  End Sub)
            End If

            Thread.Sleep(500)
            loadtask.Wait()

            If aborting Then
                Exit Sub
            End If

            Do Until moveon
                Thread.Sleep(1000)
            Loop
        Loop
    End Sub

    Private Sub mainThrd_Throw()
        Do While audiofading
            Thread.Sleep(500)
        Loop

        Dim ease_in, ease_out, ease_inout As Animation.CubicEase
        Dim anim_fadein, anim_fadeout As Animation.DoubleAnimation
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
                anim_fadein = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 2)))
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)
        Dim tgt_img As Image
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
                    SwitchTarget(tgt_img)
                    Anim_Throw(tgt_img, ease_in, ease_out, ease_inout, anim_fadein)
                End Sub)
            End If

            Dim tmpposition = position
            Dim loadtask = Task.Run(
                    Sub()
                        Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            If Not ListOfPic.Rows(tmpposition - 1)("CB_Title") Then
                Thread.Sleep((picmove_sec - 1) * 1000)
                Dispatcher.Invoke(Sub()
                                      Animation.Timeline.SetDesiredFrameRate(anim_fadeout, framerate)
                                      tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadeout)
                                  End Sub)
            End If

            loadtask.Wait()

            If aborting Then
                Exit Sub
            End If

            Do Until moveon
                Thread.Sleep(1000)
            Loop
        Loop
    End Sub

    Private Sub mainThrd_Mix()
        Do While audiofading
            Thread.Sleep(500)
        Loop
        Dim ease_in, ease_out, ease_inout As Animation.CubicEase
        Dim anim_fadein1, anim_fadein2, anim_fadeout As Animation.DoubleAnimation
        Dim tgt_img As Image
        Dim tbchkpoint As Integer = 1
        Dim last_zoom As Boolean = False
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
                anim_fadein1 = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1)))
                anim_fadein2 = New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 2)))
                anim_fadeout = New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                Panel.SetZIndex(tb_date0, 2)
                Panel.SetZIndex(tb_date1, 3)
            End Sub)

        Do
            If position >= tbchkpoint Then
                Task.Run(Sub() textThrd_KBE(position, tbchkpoint))
            End If

            If ListOfPic.Rows(position - 1)("CB_Title") Then
                Anim_TitleSlide(tgt_img)
            Else
                Dispatcher.Invoke(
                Sub()
                    tgt_img = New Image
                    tgt_img.Name = "tgt"
                    RenderOptions.SetBitmapScalingMode(tgt_img, scalemode)
                    mainGrid.Children.Add(tgt_img)
                End Sub)
                Select Case ran.Next(3)
                    Case 0
                        Dispatcher.Invoke(Sub() Anim_Breath(tgt_img, ease_in, ease_out, ease_inout, anim_fadein1, last_zoom))
                        Thread.Sleep((picmove_sec - 1.5) * 1000)
                        Dispatcher.Invoke(Sub()
                                              Dim fadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                                              AddHandler fadeout.Completed, Sub(s As Animation.AnimationClock, e As EventArgs)
                                                                                mainGrid.Children.Remove(Animation.Storyboard.GetTarget(s.Timeline))
                                                                            End Sub
                                              Animation.Storyboard.SetTarget(fadeout, tgt_img)
                                              tgt_img.BeginAnimation(Image.OpacityProperty, fadeout)
                                          End Sub)
                        Thread.Sleep(500) 'dont miss this
                    Case 1
                        Dispatcher.Invoke(Sub() Anim_Throw(tgt_img, ease_in, ease_out, ease_inout, anim_fadein2))
                        Thread.Sleep((picmove_sec - 1) * 1000)
                        Dispatcher.Invoke(Sub()
                                              Dim fadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                                              AddHandler fadeout.Completed, Sub(s As Animation.AnimationClock, e As EventArgs)
                                                                                mainGrid.Children.Remove(Animation.Storyboard.GetTarget(s.Timeline))
                                                                            End Sub
                                              Animation.Storyboard.SetTarget(fadeout, tgt_img)
                                              tgt_img.BeginAnimation(Image.OpacityProperty, fadeout)
                                          End Sub)
                    Case Else
                        Dispatcher.Invoke(Sub() Anim_KBE(tgt_img, ease_in, ease_out, ease_inout, anim_fadein1))
                        Thread.Sleep((picmove_sec - 1) * 1000)
                        Dispatcher.Invoke(Sub()
                                              Dim fadeout As New Animation.DoubleAnimation(0, New Duration(New TimeSpan(0, 0, 1)))
                                              AddHandler fadeout.Completed, Sub(s As Animation.AnimationClock, e As EventArgs)
                                                                                mainGrid.Children.Remove(Animation.Storyboard.GetTarget(s.Timeline))
                                                                            End Sub
                                              Animation.Storyboard.SetTarget(fadeout, tgt_img)
                                              tgt_img.BeginAnimation(Image.OpacityProperty, fadeout)
                                          End Sub)
                End Select
            End If

            Dim loadtask = Task.Run(
                    Sub()
                        Thread.CurrentThread.Priority = ThreadPriority.Lowest
                        If position = ListOfPic.Rows.Count Then
                            position = 0
                            tbchkpoint = 1
                        End If
                        LoadNextImg()
                    End Sub)

            loadtask.Wait()

            If aborting Then
                Exit Sub
            End If

            Do Until moveon
                Thread.Sleep(1000)
            Loop
        Loop
    End Sub

    Private Sub SwitchTarget(ByRef tgt_img As Image)
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
    End Sub

    Private Sub LoadNextImg()
        Select Case loadmode
            Case 1
                Dispatcher.Invoke(Sub()
                                      RenderOptions.SetBitmapScalingMode(slide_img0, scalemode)
                                      RenderOptions.SetBitmapScalingMode(slide_img1, scalemode)
                                  End Sub)
                pic = pics(position)
                position += 1
            Case 0
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

                        strm.Position = 0
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
                        pic.StreamSource = strm
                        pic.EndInit()
                    End Using

                    'reading next picture gradually proved to be futile
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
                    position += 1
                    pic.Freeze()
                End Try
        End Select
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.F12 Then
            Dim optwin As New OptWindow
            optwin.ShowDialog()
            optwin.Close()
        ElseIf e.Key = Key.F11 Then
            Dim editwin As New EditWindow
            editwin.ShowDialog()
            editwin.Close()
        ElseIf e.Key = Key.F1 Then
            If ctrlwindow Is Nothing Then
                ctrlwindow = New ControlWindow
                ctrlwindow.Owner = Me
            End If
            ctrlwindow.Show()
        ElseIf e.Key = Key.P Then
            If Keyboard.Modifiers = ModifierKeys.Control Then 'pause image
                SwitchImage()
            ElseIf Keyboard.Modifiers = ModifierKeys.Shift Then 'fadeout audio only
                SwitchAudio()
            End If
        ElseIf e.Key = Key.R AndAlso Keyboard.Modifiers = ModifierKeys.Control AndAlso aborting = False Then
            RestartAll()
            'ElseIf e.Key = Key.S AndAlso Keyboard.Modifiers = ModifierKeys.Control AndAlso aborting = False Then
            '    RestartAll()
        End If
    End Sub

    Friend Sub SwitchImage()
        If worker_pic IsNot Nothing AndAlso worker_pic.IsAlive Then
            If moveon = True Then
                moveon = False
            Else
                moveon = True
            End If
        End If
    End Sub

    Friend Sub SwitchAudio()
        If worker_pic IsNot Nothing AndAlso worker_pic.IsAlive Then
            If Not audiofading Then
                If playing Then
                    Task.Run(AddressOf FadeoutAudio)
                Else
                    Task.Run(AddressOf FadeinAudio)
                End If
            End If
        End If
    End Sub

    Friend Sub RestartAll()
        If worker_pic IsNot Nothing AndAlso worker_pic.IsAlive Then
            aborting = True
            Task.Run(AddressOf FadeoutAudio)
            Task.Run(Sub()
                         moveon = True
                         Dim black As Rectangle
                         Dispatcher.Invoke(Sub()
                                               black = New Rectangle
                                               mainGrid.Children.Add(black)
                                               Panel.SetZIndex(black, 9)
                                               black.Fill = Windows.Media.Brushes.Black
                                               black.Width = w
                                               black.Height = h
                                               black.BeginAnimation(OpacityProperty, New Animation.DoubleAnimation(0, 1, New Duration(New TimeSpan(0, 0, 1))))
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
                             Case 0
                                 worker_pic = New Thread(AddressOf mainThrd_KBE)
                             Case 1
                                 worker_pic = New Thread(AddressOf mainThrd_Breath)
                             Case 2
                                 worker_pic = New Thread(AddressOf mainThrd_Throw)
                             Case 3
                                 worker_pic = New Thread(AddressOf mainThrd_Mix)
                             Case Else
                                 MsgBox(Application.Current.Resources("msg_transerr"), MsgBoxStyle.Critical)
                                 Me.Close()
                                 Exit Sub
                         End Select
                         worker_pic.IsBackground = True
                         worker_pic.Priority = ThreadPriority.Lowest
                         aborting = False
                         worker_pic.Start()
                         Dispatcher.Invoke(Sub()
                                               If ctrlwindow Is Nothing Then
                                                   ctrlwindow = New ControlWindow
                                                   ctrlwindow.Owner = Me
                                               End If
                                               ctrlwindow.Show()
                                           End Sub)
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
        ControlWindow.reallyclose = True
        If ctrlwindow IsNot Nothing Then ctrlwindow.Close()
        If ExecState_Set Then SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS)
    End Sub

    Private Sub ApplyBlur(tgt As Image, easeout As Animation.IEasingFunction, easein As Animation.IEasingFunction,
                          Optional fadeindur As Double = 2, Optional fadeoutdur As Double = 2)
        'blur
        If blurmode <> 0 Then
            Dim rad = w * h / 15000
            Dim blur_fx As New Effects.BlurEffect
            Dim blur_anim As New Animation.DoubleAnimationUsingKeyFrames
            If blurmode = 1 Then 'only fade in
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(rad, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(fadeindur)), easeout))
            ElseIf blurmode = 2 Then 'only fade out
                blur_fx.Radius = 0
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec - fadeoutdur))))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(rad, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), easein))
            ElseIf blurmode = 3 Then 'both
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(rad, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(fadeindur)), easeout))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec - fadeoutdur))))
                blur_anim.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(rad, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), easein))
            End If
            tgt.Effect = blur_fx
            Animation.Timeline.SetDesiredFrameRate(blur_anim, framerate)
            blur_fx.BeginAnimation(Effects.BlurEffect.RadiusProperty, blur_anim)
        Else
            tgt.Effect = Nothing
        End If
    End Sub

    Private Sub Anim_KBE(tgt_img As Image,
                         ease_in As Animation.IEasingFunction, ease_out As Animation.IEasingFunction, ease_inout As Animation.IEasingFunction,
                         anim_fadein As Animation.DoubleAnimation)
        Dim delta As Double
        Dim tgt_trasform As New TransformGroup
        Dim trans_scale As ScaleTransform
        Dim trans_trans As TranslateTransform
        Dim anim_move As Animation.DoubleAnimation
        Dim anim_zoom As Animation.DoubleAnimation
        Dim XorY As DependencyProperty

        If pic.PixelWidth / pic.PixelHeight > w / h Then
            'width is the longer edge comparing to the size of the monitor
            tgt_img.Height = h
            tgt_img.Width = tgt_img.Height * pic.PixelWidth / pic.PixelHeight
            XorY = TranslateTransform.XProperty
            Dim zoomin As Boolean
            delta = tgt_img.Width - w

            If ran.Next(2) = 0 Then
                'zoom in
                zoomin = True
                anim_zoom = New Animation.DoubleAnimation(1, 1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
            Else
                'zoom out
                zoomin = False
                anim_zoom = New Animation.DoubleAnimation(1.2, 1, New Duration(New TimeSpan(0, 0, picmove_sec)))
            End If

            Dim startpoint As Double = delta
            If horizontalOptimize AndAlso delta > h * 0.7 Then
                startpoint = delta * horizontalOptimizeR
            End If

            'move left or right
            If ran.Next(2) = 0 Then 'means 0<=ran<2
                'move left
                tgt_img.RenderTransformOrigin = New Point(0, 0.5)
                anim_move = New Animation.DoubleAnimation(startpoint - delta, If(zoomin = True, -delta - tgt_img.Width * 0.2, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
            Else
                tgt_img.RenderTransformOrigin = New Point(1, 0.5)
                anim_move = New Animation.DoubleAnimation(-startpoint, If(zoomin = True, tgt_img.Width * 0.2, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
            End If
        Else
            'height is the longer edge comparing to the size of the monitor
            tgt_img.Width = w
            tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
            XorY = TranslateTransform.YProperty
            Dim zoomin As Boolean
            delta = tgt_img.Height - h

            If ran.Next(2) = 0 Then
                'zoom in
                zoomin = True
                anim_zoom = New Animation.DoubleAnimation(1, 1.2, New Duration(New TimeSpan(0, 0, picmove_sec)))
            Else
                'zoom out
                zoomin = False
                anim_zoom = New Animation.DoubleAnimation(1.2, 1, New Duration(New TimeSpan(0, 0, picmove_sec)))
            End If

            Dim startpoint As Double = delta
            If verticalOptimize AndAlso delta > h * 0.7 Then
                startpoint = delta * verticalOptimizeR
            End If

            'move up or down or up only when long
            If verticalLock AndAlso tgt_img.Height > h * 1.2 Then
                'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'transform align with bottom
                anim_move = New Animation.DoubleAnimation(-startpoint, If(zoomin = True, tgt_img.Height * 0.2, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
            Else
                If ran.Next(2) = 0 Then
                    tgt_img.RenderTransformOrigin = New Point(0.5, 0) 'transform align with top
                    anim_move = New Animation.DoubleAnimation(startpoint - delta, If(zoomin = True, -delta - tgt_img.Height * 0.2, -delta), New Duration(New TimeSpan(0, 0, picmove_sec)))
                Else
                    tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                    anim_move = New Animation.DoubleAnimation(-startpoint, If(zoomin = True, tgt_img.Height * 0.2, 0), New Duration(New TimeSpan(0, 0, picmove_sec)))
                End If
            End If
        End If
        tgt_img.Source = pic
        trans_scale = New ScaleTransform
        tgt_trasform.Children.Add(trans_scale)
        trans_trans = New TranslateTransform
        tgt_trasform.Children.Add(trans_trans)
        tgt_img.RenderTransform = tgt_trasform
        Animation.Timeline.SetDesiredFrameRate(anim_zoom, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
        trans_scale.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoom)
        trans_scale.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoom)
        trans_trans.BeginAnimation(XorY, anim_move)
        tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein)

        ApplyBlur(tgt_img, ease_out, ease_in)
    End Sub

    Private Sub Anim_Breath(tgt_img As Image,
                            ease_in As Animation.IEasingFunction, ease_out As Animation.IEasingFunction, ease_inout As Animation.IEasingFunction,
                            anim_fadein As Animation.DoubleAnimation,
                            ByRef last_zoom As Boolean)
        Dim delta As Double
        Dim tgt_trasform As New TransformGroup
        Dim trans_scale As ScaleTransform
        Dim trans_trans As TranslateTransform
        Dim anim_move As Animation.DoubleAnimation
        Dim anim_zoom As Animation.DoubleAnimationUsingKeyFrames
        Dim XorY As DependencyProperty

        If pic.PixelWidth / pic.PixelHeight > w / h Then
            'width is the longer edge comparing to the size of the monitor
            tgt_img.Height = h
            tgt_img.Width = tgt_img.Height * pic.PixelWidth / pic.PixelHeight
            delta = tgt_img.Width - w
            XorY = TranslateTransform.XProperty

            If last_zoom Then ' ran.Next(2) = 0 Then
                'zoom in
                last_zoom = False
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
            Else
                'zoom out
                last_zoom = True
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
            End If

            Dim startpoint As Double = delta
            If horizontalOptimize AndAlso delta > h * 0.7 Then
                startpoint = delta * horizontalOptimizeR
            End If

            If ran.Next(2) = 0 Then
                tgt_img.RenderTransformOrigin = New Point(0, 0.5) 'transform align with left
                anim_move = New Animation.DoubleAnimation(startpoint - delta, If(last_zoom = False, -delta - tgt_img.Width * 0.1, -delta), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
            Else
                tgt_img.RenderTransformOrigin = New Point(1, 0.5) 'transform align with right
                anim_move = New Animation.DoubleAnimation(-startpoint, If(last_zoom = False, tgt_img.Width * 0.1, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
            End If
            If tgt_img.Width > w * 1.2 Then
                anim_move.EasingFunction = ease_inout
            End If
        Else
            'height is the longer edge comparing to the size of the monitor
            tgt_img.Width = w
            tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
            XorY = TranslateTransform.YProperty
            delta = tgt_img.Height - h

            If last_zoom Then 'ran.Next(2) = 0 Then
                'zoom in
                last_zoom = False
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0.1, 0.4, 0.7, 1)))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.1, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.5, 0, 1, 1)))
            Else
                'zoom out
                last_zoom = True
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.3, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), ease_in))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec * 0.5)), New Animation.KeySpline(0, 0.6, 0.7, 1)))
                'anim_zoomx.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(1.2, Animation.KeyTime.FromTimeSpan(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)), New Animation.KeySpline(0.8, 0, 1, 0.7)))
            End If

            Dim startpoint As Double = delta
            If verticalOptimize AndAlso delta > h * 0.7 Then
                startpoint = delta * verticalOptimizeR
            End If
            'If verticalOptimize AndAlso delta > h * 0.7 Then
            '    startpoint = delta * verticalOptimizeR
            '    If startpoint > tgt_img.Height - h Then
            '        startpoint = tgt_img.Height - h
            '    End If
            'Else
            '    startpoint = 0
            'End If

            'move up or down or up only when long
            If verticalLock AndAlso tgt_img.Height > h * 1.2 Then
                'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'transform align with bottom
                anim_move = New Animation.DoubleAnimation(-startpoint, If(last_zoom = False, tgt_img.Height * 0.1, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500))) With {.EasingFunction = ease_inout}
            Else
                If ran.Next(2) = 0 Then
                    tgt_img.RenderTransformOrigin = New Point(0.5, 0) 'transform align with top
                    anim_move = New Animation.DoubleAnimation(startpoint - delta, If(last_zoom = False, -delta - tgt_img.Height * 0.1, -delta), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                Else
                    tgt_img.RenderTransformOrigin = New Point(0.5, 1) 'transform align with bottom
                    anim_move = New Animation.DoubleAnimation(-startpoint, If(last_zoom = False, tgt_img.Height * 0.1, 0), New Duration(New TimeSpan(0, 0, 0, picmove_sec - 1, 500)))
                End If
                If tgt_img.Height > h * 1.2 Then
                    anim_move.EasingFunction = ease_inout
                End If
            End If
        End If
        tgt_img.Source = pic
        trans_scale = New ScaleTransform
        tgt_trasform.Children.Add(trans_scale)
        trans_trans = New TranslateTransform
        tgt_trasform.Children.Add(trans_trans)
        tgt_img.RenderTransform = tgt_trasform
        Animation.Timeline.SetDesiredFrameRate(anim_zoom, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
        trans_scale.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoom)
        trans_scale.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoom)
        trans_trans.BeginAnimation(XorY, anim_move)
        tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein)

        ApplyBlur(tgt_img, ease_out, ease_in, 2, 2.5)
    End Sub

    Private Sub Anim_Throw(tgt_img As Image,
                            ease_in As Animation.IEasingFunction, ease_out As Animation.IEasingFunction, ease_inout As Animation.IEasingFunction,
                            anim_fadein As Animation.DoubleAnimation)
        'Dim delta As Double
        Dim tgt_trasform As New TransformGroup
        Dim trans_scale As ScaleTransform
        Dim trans_trans As TranslateTransform
        Dim anim_move As Animation.DoubleAnimationUsingKeyFrames
        Dim anim_zoom As Animation.DoubleAnimationUsingKeyFrames
        Dim XorY As DependencyProperty

        If pic.PixelWidth / pic.PixelHeight > w / h Then
            'width is the longer edge comparing to the size of the monitor
            tgt_img.Height = h
            tgt_img.Width = tgt_img.Height * pic.PixelWidth / pic.PixelHeight
            'delta = tgt_img.Width - w
            XorY = TranslateTransform.XProperty

            'move left or right
            If ran.Next(2) = 0 Then 'means 0<=ran<2
                'move left
                tgt_img.RenderTransformOrigin = New Point(0, 0.5)
                anim_move = New Animation.DoubleAnimationUsingKeyFrames
                anim_move.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(w, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_move.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(-tgt_img.Width, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), New Animation.KeySpline(0.2, 0.8, 0.8, 0.2)))
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec / 2)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), ease_in))
            Else
                tgt_img.RenderTransformOrigin = New Point(1, 0.5)
                anim_move = New Animation.DoubleAnimationUsingKeyFrames
                anim_move.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(-tgt_img.Width, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_move.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(w, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), New Animation.KeySpline(0.2, 0.8, 0.8, 0.2)))
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec / 2)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), ease_in))
            End If
        Else
            'height is the longer edge comparing to the size of the monitor
            tgt_img.Width = w
            tgt_img.Height = tgt_img.Width / pic.PixelWidth * pic.PixelHeight
            'delta = tgt_img.Height - h
            XorY = TranslateTransform.YProperty

            'move up or down or up only when long
            If verticalLock AndAlso tgt_img.Height > h * 1.2 Then
                'only move down for pics with height larger than 1.5 * screen height after converted to same width as screen
                'move down
                tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                anim_move = New Animation.DoubleAnimationUsingKeyFrames
                anim_move.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(-tgt_img.Height, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_move.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(h, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), New Animation.KeySpline(0.2, 0.8, 0.8, 0.2)))
                anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec / 2)), ease_out))
                anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), ease_in))
            Else
                If ran.Next(2) = 0 Then
                    tgt_img.RenderTransformOrigin = New Point(0.5, 0)
                    anim_move = New Animation.DoubleAnimationUsingKeyFrames
                    anim_move.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(h, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                    anim_move.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(-tgt_img.Height, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), New Animation.KeySpline(0.2, 0.8, 0.8, 0.2)))
                    anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec / 2)), ease_out))
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), ease_in))
                Else
                    tgt_img.RenderTransformOrigin = New Point(0.5, 1)
                    anim_move = New Animation.DoubleAnimationUsingKeyFrames
                    anim_move.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(-tgt_img.Height, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                    anim_move.KeyFrames.Add(New Animation.SplineDoubleKeyFrame(h, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), New Animation.KeySpline(0.2, 0.8, 0.8, 0.2)))
                    anim_zoom = New Animation.DoubleAnimationUsingKeyFrames
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))))
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(1, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec / 2)), ease_out))
                    anim_zoom.KeyFrames.Add(New Animation.EasingDoubleKeyFrame(0.7, Animation.KeyTime.FromTimeSpan(TimeSpan.FromSeconds(picmove_sec)), ease_in))
                End If
            End If
        End If

        tgt_img.Source = pic
        trans_scale = New ScaleTransform
        tgt_trasform.Children.Add(trans_scale)
        trans_trans = New TranslateTransform
        tgt_trasform.Children.Add(trans_trans)
        tgt_img.RenderTransform = tgt_trasform
        Animation.Timeline.SetDesiredFrameRate(anim_zoom, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_move, framerate)
        Animation.Timeline.SetDesiredFrameRate(anim_fadein, framerate)
        trans_scale.BeginAnimation(ScaleTransform.ScaleXProperty, anim_zoom)
        trans_scale.BeginAnimation(ScaleTransform.ScaleYProperty, anim_zoom)
        trans_trans.BeginAnimation(XorY, anim_move)
        tgt_img.BeginAnimation(Image.OpacityProperty, anim_fadein)

        ApplyBlur(tgt_img, ease_out, ease_in)
    End Sub
End Class

'Class CustomEase
'    Inherits Animation.EasingFunctionBase

'    Protected Overrides Function CreateInstanceCore() As Freezable
'        Return New CustomEase
'    End Function

'    Protected Overrides Function EaseInCore(normalizedTime As Double) As Double
'        If normalizedTime < 0.6011111 Then
'            Return (Math.Sin(5 * normalizedTime + 1.5 * Math.PI) + 1) / 2.26
'        Else
'            Return 0.3 * normalizedTime + 0.7
'        End If
'    End Function
'End Class
