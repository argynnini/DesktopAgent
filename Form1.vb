Public Class Form1
    Private ReadOnly SearchName(,) As String = New String(,) {
                                               {"Google", "https://www.google.co.jp/search?q="},
                                               {"Yahoo!", "https://search.yahoo.co.jp/search?p="},
                                               {"YouTube", "https://www.youtube.com/search?q="},
                                               {"ニコニコ動画", "https://www.nicovideo.jp/search/"},
                                               {"Twitter", "https://twitter.com/search?q="},
                                               {"Googleマップ", "https://www.google.com/maps/search/"},
                                               {"Amazon", "https://www.amazon.co.jp/s?k="},
                                               {"楽天市場", "https://search.rakuten.co.jp/search/mall/"},
                                               {"ヤフオク！", "https://auctions.yahoo.co.jp/search/search?p="},
                                               {"メルカリ", "https://www.mercari.com/jp/search/?keyword="},
                                               {"Wikipedia", "https://ja.wikipedia.org/wiki/"}}
    '起動時
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AxAgent.Characters.Load("Agent", "DOLPHIN.ACS") 'DOLPHIN.ACSを読み込み
        With AxAgent.Characters("Agent") 'Agentに対して
            .Balloon.Style = 3 '吹き出しの表示スタイルを設定
            .LanguageID = &H411 '言語を日本語に設定
            .Balloon.FontCharSet = 128 '吹き出しのフォントを日本語に設定
            .Balloon.Visible = False '吹き出しを非表示にする
            .AutoPopupMenu = False 'デフォルトの右クリックメニューを無効にする
            .SoundEffectsOn = True '音を有効にする
            .Top = Screen.PrimaryScreen.Bounds.Height - .OriginalHeight - 100 'AgentのY座標をスクリーンの高さ-Agentの高さ-100に設定
            .Left = Screen.PrimaryScreen.Bounds.Width - .OriginalWidth - 50 'AgentのX座標をスクリーンの幅-Agentの幅-50に設定
            If AxAgent.Characters("Agent").VisibilityCause = 0 And
               AxAgent.Characters("Agent").Visible = False Then 'イルカが既出でない場合
                .Show() 'イルカを出す
            End If
        End With 'Agentに対して終了
        ShowInTaskbar = False 'タスクバーに非表示
        FormBorderStyle = FormBorderStyle.None 'フォームの境界線をなくす
        Size = New Size(300, 170) '大きさを変更
        StartPosition = FormStartPosition.Manual 'フォームの表示位置をマニュアルに設定
        Location = New Point(AxAgent.Characters("Agent").Left - 150, AxAgent.Characters("Agent").Top - 150) 'Agentを表示
        Search.BackColor = Color.FromArgb(255, 255, 154) '検索ボタンの背景色を設定
        CloseButton.BackColor = Color.FromArgb(255, 255, 154) '閉じるボタンの背景色を設定
        Label1.BackColor = Color.FromArgb(255, 255, 154) '検索ボタンの背景色を設定
        Dim bmp As New Bitmap(My.Resources.Ballon) '画像を読み込む
        Dim transColor As Color = bmp.GetPixel(0, 0) '透明にする色
        bmp.MakeTransparent(transColor) '画像を透明に
        BackgroundImage = bmp '背景画像を指定
        BackColor = transColor '背景色を指定
        TransparencyKey = transColor '透明を指定
        TopMost = True '一番手前に表示
        AxAgent.Characters("Agent").SoundEffectsOn = My.Settings.DefaultSound
        SearchEngine.Items.Clear()
        SearchEngine.BeginUpdate() '再描画しないようにする
        For i = 0 To (SearchName.Length / 2) - 1
            SearchEngine.Items.Add(SearchName(i, 0)) '配列の内容を一つ一つ追加する
        Next
        SearchEngine.EndUpdate() '再描画するようにする
        SearchEngine.SelectedIndex = My.Settings.DefaultSearchEngine
        ToolTip.SetToolTip(Search, SearchEngine.Text.ToString & " で検索します。")
        ToolTip.SetToolTip(CloseButton, "吹き出しを閉じます。")
        NotifyIcon.ContextMenuStrip = AgentMenu
        With Animation 'アニメーションボタン
            .Items.Add("アニメーション") 'アニメーション（何もしない）を追加
            For Each AnimationList In AxAgent.Characters("Agent").AnimationNames 'イルカのアニメーションを登録
                .Items.Add(AnimationList)
            Next
            .SelectedIndex = .Items.Count() - 1 'アニメーションを既定の選択にする
        End With
    End Sub

    '検索ボックス入力時
    Private Sub SearchBox_KeyPress（sender As Object, e As KeyPressEventArgs） Handles SearchBox.KeyPress
        If e.KeyChar = vbCr Then e.Handled = True
    End Sub

    '検索ボックス入力時
    Private Sub SearchBox_TextUpdate(sender As Object, e As EventArgs） Handles SearchBox.KeyPress
        With AxAgent.Characters("Agent")
            .Balloon.Visible = False
            .Play("Writing")
        End With
    End Sub

    'Agentクリック時
    Private Sub Agent_ActivateInput(sender As Object, e As AxAgentObjects._AgentEvents_ClickEvent) Handles AxAgent.ClickEvent
        With AxAgent.Characters("Agent")
            .StopAll()
            .Play("RestPose")
            .Balloon.Visible = False
        End With
        Select Case e.button
            Case 1 '左クリック時
                AgentMenu.Hide()
                Show()
                Exit Select
            Case 2 '右クリック時
                Animation.SelectedIndex = Animation.Items.Count() - 1 'アニメーションを既定の選択にする
                AgentMenu.Show(Cursor.Position.X, Cursor.Position.Y)
                Exit Select
            Case 3 '真中クリック時
                Exit Select
        End Select
    End Sub

    'Agentドラッグ時
    Private Sub Agent_Dragstart(sender As Object, e As AxAgentObjects._AgentEvents_DragStartEvent) Handles AxAgent.DragStart
        AxAgent.Characters("Agent").Balloon.Visible = False
        Hide()
    End Sub

    'Agentドラッグ終了時
    Private Sub Agent_DragEnd(sender As Object, e As AxAgentObjects._AgentEvents_DragCompleteEvent) Handles AxAgent.DragComplete
        Location = New Point(AxAgent.Characters("Agent").Left - 150, AxAgent.Characters("Agent").Top - 150)
    End Sub

    '検索クリック時
    Private Sub Search_Click(sender As Object, e As EventArgs) Handles Search.Click
        With AxAgent.Characters("Agent")
            Hide()
            .StopAll()
            If SearchBox.Text.Trim = "" Then
                .Play("RestPose")
                Return
            End If
            If SearchBox.Text = "お前を消す方法" Then
                Hide()
                .Balloon.FontSize = 12
                .Speak("質問の意味がわかりません。")
                .Play("wave")
            End If
            '選択された検索エンジン
            Debug.WriteLine(SearchEngine.SelectedIndex)
            Process.Start(SearchName(SearchEngine.SelectedIndex, 1).ToString & Web.HttpUtility.UrlEncode(SearchBox.Text.ToString))
            .Play("RestPose")
        End With
    End Sub

    '閉じるクリック時
    Private Sub Close_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        With AxAgent.Characters("Agent")
            .StopAll()
            .Balloon.Visible = False
            .Play("RestPose")
        End With
        Hide()
    End Sub

    '終了クリック時
    Private Sub MenuExit_Click(sender As Object, e As EventArgs) Handles MenuExit.Click
        Hide()
        AgentMenu.Hide()
        With AxAgent.Characters("Agent")
            .StopAll()
            .Balloon.Visible = False
            .Play("GoodBye")
            .Hide(True)
        End With
        Do While AxAgent.Characters("Agent").Visible = True
            Threading.Thread.Sleep(50)
        Loop
        Application.Exit()
    End Sub

    'アニメーション
    Private Sub Animation_Click(sender As Object, e As EventArgs) Handles Animation.DropDownClosed
        With AxAgent.Characters("Agent")
            .StopAll()
            .Balloon.Visible = False 'イルカの吹き出しを隠す
            Hide()
            AgentMenu.Hide()
            If Animation.Text = "アニメーション" Then
                .Play("RestPose")
                Exit Sub
            Else
                .Play(Animation.SelectedItem.ToString)
            End If
        End With
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon.MouseClick
        If e.Button = MouseButtons.Left Then
            With AxAgent.Characters("Agent")
                .Balloon.Visible = False
                .StopAll()
                .Play("RestPose")
            End With
            If Visible = False Then
                Show()
            Else
                Hide()
            End If
        End If
    End Sub

    Private Sub SearchEngine_Changed(sender As Object, e As EventArgs) Handles SearchEngine.SelectedIndexChanged
        My.Settings.DefaultSearchEngine = SearchEngine.SelectedIndex
        My.Settings.Save()
        ToolTip.SetToolTip(Search, SearchEngine.Text.ToString & " で検索します。")
    End Sub

    Private Sub Sound_Changed(sender As Object, e As EventArgs) Handles Sound.Click
        AxAgent.Characters("Agent").SoundEffectsOn = Not My.Settings.DefaultSound
        My.Settings.DefaultSound = Not My.Settings.DefaultSound
        My.Settings.Save()
    End Sub

    Private Sub TextBox_Enter(sender As Object, e As EventArgs) Handles SearchBox.GotFocus, SearchBox.MouseDown
        SearchBox.SelectAll()
    End Sub
End Class