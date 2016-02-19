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
                    T_DateShown.Text = "On-screen date shown:" & vbCrLf & Date.Parse(row("Date")).ToString("yyyy.M")
                Else
                    T_DateShown.Text = "On-screen date shown:" & vbCrLf & "None"
                End If
                CB_Title.IsChecked = row("CB_Title")
            End If

            'For Each row As System.Data.DataRow In ListOfPic.Rows
            '    If row("Path").ToString.Contains(e.AddedItems(0).ToString) Then
            '        If Not IsDBNull(row("Date")) Then
            '            T_DateShown.Text = "On-screen date shown: " & Date.Parse(row("Date")).ToString("yyyy.M")
            '        Else
            '            T_DateShown.Text = "On-screen date shown: None"
            '        End If
            '        CB_Title.IsChecked = row("CB_Title")
            '    End If
            'Next
        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Using lop_str = New IO.StringWriter()
            ListOfPic.WriteXml(lop_str)
            Dim lop = XElement.Parse(lop_str.ToString)
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
End Class
