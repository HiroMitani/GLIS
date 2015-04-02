Imports System.Globalization

Public Class zrireki

    Dim PLACE_ID As String

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Private Sub zrireki_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "在庫履歴検索"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label1.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '倉庫情報取得
        Dim PlaceTable As New DataTable()
        PlaceTable.Columns.Add("ID", GetType(String))
        PlaceTable.Columns.Add("NAME", GetType(String))
        '倉庫情報の取得
        Result = GetPLACEList(PlaceData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        For Count = 0 To PlaceData.Length - 1
            '新しい行を作成
            Dim row As DataRow = PlaceTable.NewRow()

            '各列に値をセット
            row("ID") = PlaceData(Count).ID
            row("NAME") = PlaceData(Count).NAME

            'DataTableに行を追加
            PlaceTable.Rows.Add(row)
        Next

        PlaceTable.AcceptChanges()
        'コンボボックスのDataSourceにDataTableを割り当てる
        ComboBox4.DataSource = PlaceTable

        '表示される値はDataTableのNAME列
        ComboBox4.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox4.ValueMember = "ID"
        '初期値をセット(倉庫データの1件目）
        ComboBox4.SelectedIndex = 0
        PLACE_ID = ComboBox4.SelectedValue.ToString()


        Dim h As Integer
        'ディスプレイの高さ
        h = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height

        If h <= 850 Then
            'フォームサイズを修正
            Me.Size = New Size(1100, 700)
            'DataGridViewサイズ変更
            Me.DataGridView1.Size = New Point(1047, 369)

            'サマリーデータCSV出力ボタン位置移動
            Me.Button3.Location = New Point(614, 625)

            'CSV出力ボタン位置移動
            Me.Button4.Location = New Point(785, 625)
            '閉じるボタン位置移動
            Me.Button1.Location = New Point(962, 625)
        End If

        TextBox1.Focus()

    End Sub

    Private Sub zrireki_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'ダイアログ設定
        Dim closecheck As DialogResult = MessageBox.Show("終了してもいいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)

        If closecheck = DialogResult.Yes Then
            'ダイアログで「はい」が選択された時 
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As StockLog_List = Nothing

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing

        Dim Date_Check_Result As DateTime

        Dim DataTotal As Integer = 0
        Dim DataNumTotal As Integer = 0

        Dim ChkItemCodeString As String = Nothing
        Dim ChkCommentString As String = Nothing

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        '作業日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            MsgBox("作業日は必須項目です。")
            MaskedTextBox1.BackColor = Color.Salmon
            Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("作業日が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '作業日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            MsgBox("作業日は必須項目です。")
            MaskedTextBox2.BackColor = Color.Salmon
            Exit Sub

        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("作業日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        ChkItemCodeString = Trim(TextBox1.Text)
        ChkCommentString = Trim(TextBox2.Text)

        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        'コメントに'が入力されていたらReplaceする。
        ChkCommentString = ChkCommentString.Replace("'", "''")

        '検索Function
        Result = GetStockLogSeach(ChkItemCodeString, Date_From, Date_To, CheckBox7.Checked, CheckBox8.Checked, CheckBox14.Checked, _
                            CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, _
                            CheckBox5.Checked, CheckBox6.Checked, CheckBox9.Checked, CheckBox10.Checked, _
                            CheckBox11.Checked, CheckBox12.Checked, CheckBox13.Checked, ChkCommentString, _
                            PLACE_ID, SearchResult, DataTotal, DataNumTotal, Result, ErrorMessage)
        If Result = False Then
            Label9.Text = "商品数： "
            Label10.Text = "総数： "
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '商品数を表示
        Label9.Text = "商品数： " & DataTotal

        '総数を表示
        Label10.Text = "総数： " & DataNumTotal

        '結果を元にDataGridViewに表示する。
        'DataGridへ入力したデータを挿入
        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '商品コード
                .Cells(0).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(1).Value = SearchResult(Count).I_NAME
                '不良区分
                .Cells(2).Value = SearchResult(Count).NUM
                '数量
                .Cells(3).Value = SearchResult(Count).I_STATUS
                'ステータス
                .Cells(4).Value = SearchResult(Count).I_FLG
                'パッケージNo
                .Cells(5).Value = SearchResult(Count).PACKAGE_NO
                '備考
                .Cells(6).Value = SearchResult(Count).COMMENT
                '作業日時
                .Cells(7).Value = SearchResult(Count).U_DATE
                '倉庫
                .Cells(8).Value = SearchResult(Count).PLACE
                'I_ID
                .Cells(9).Value = SearchResult(Count).I_ID
            End With
            DataGridView1.Rows.Add(SR_List)
        Next
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemName() As Item_List = Nothing

        Dim ChkItemCodeString As String = Nothing

        ChkItemCodeString = Trim(TextBox1.Text)

        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        '商品名欄をクリアする。
        Label7.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '入力チェック
            If Trim(TextBox1.Text) <> "" Then
                'あたかもTabキーが押されたかのようにする
                'Shiftが押されている時は前のコントロールのフォーカスを移動
                Me.ProcessTabKey(Not e.Shift)
                e.Handled = True
                '入力された商品コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetItemName(ChkItemCodeString, 1, ItemName, Result, ErrorMessage)
                If Result = "True" Then
                    Label7.Text = ItemName(0).I_NAME
                    CheckBox1.Focus()
                    TextBox1.BackColor = Color.White
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox1.Focus()
                    TextBox1.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    Label7.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        topmenu.Show()
        Me.Dispose()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'サマリーデータファイル作成
        '作業日付での範囲の入庫と出庫のサマリーをプロダクトラインコード別に出力する。

        Dim Date_From As Date = Nothing
        Dim Date_To As Date = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Date_Check_Result As DateTime

        Dim In_List() As Summary_List = Nothing
        Dim Out_List() As Summary_List = Nothing
        Dim Tanaoroshi_List() As Summary_List = Nothing
        Dim PL_List() As PL_List = Nothing

        Dim dtNow As DateTime = DateTime.Now

        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        '作業日付が設定されていないとエラーメッセージ表示
        '作業日付Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力ならメッセージ表示
            MsgBox("サマリーデータを作成する場合は作業日付を指定してください。")
            MaskedTextBox1.Focus()
            MaskedTextBox1.BackColor = Color.Salmon

            Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("作業日付の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '作業日付Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力ならメッセージ表示
            MsgBox("サマリーデータを作成する場合は作業日付を指定してください。")
            MaskedTextBox2.Focus()
            MaskedTextBox2.BackColor = Color.Salmon
            Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("作業日付の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '日付の比較
        'FromがToより未来ならエラー
        If Date_From > Date_To Then
            MsgBox("作業日付FromがToより未来の日付になっています。")
            Exit Sub
        End If

        '入庫と出庫の履歴を取得
        Result = GetHistorySummary(Date_From, Date_To, PLACE_ID, In_List, Out_List, Tanaoroshi_List, PL_List, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "在庫実績サマリーデータ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "日付,区分,"

        For i = 0 To PL_List.Length - 1
            Sheet1_Header &= PL_List(i).NAME
            If i <> PL_List.Length - 1 Then
                Sheet1_Header &= ","
            End If
        Next

        'ﾌｧｲﾙ名の設定
        strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)
        'headerを設定
        strStreamWriter.WriteLine(Sheet1_Header)

        For i = 0 To In_List.Length - 1
            '入庫データ
            '取得したデータをLineataに格納する。
            'A列に日付
            LineData = """" & In_List(i).WorkDate & ""","
            'B列に区分
            LineData &= """入庫"","
            'C列にライン
            LineData &= In_List(i).Line & ","
            'D列にベイト＆ガルプ
            LineData &= In_List(i).Bait & ","
            'E列にサオ
            LineData &= In_List(i).Rods & ","
            'F列にBaitcast_Reels
            LineData &= In_List(i).B_Reels & ","
            'G列にSpinning_Reels
            LineData &= In_List(i).S_Reels & ","
            'H列にCombos
            LineData &= In_List(i).Combos & ","
            'I列にAccessory
            LineData &= In_List(i).Accessory & ","
            'I列にAccessory
            'LineData &= In_List(i).Accessory2 & ","
            'J列にHard_Lure
            LineData &= In_List(i).Hard_Lure & ","
            'K列にBagApparel
            LineData &= In_List(i).Bag & ","
            'L列にRodOEM
            LineData &= In_List(i).Rod_OEM & ","
            'M列にRodParts
            LineData &= In_List(i).Rod_Parts & ","
            'N列にReelParts
            LineData &= In_List(i).ReelParts & ","
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)

            '出庫データ
            '取得したデータをLineataに格納する。
            'A列に日付
            LineData = """" & Out_List(i).WorkDate & ""","
            'B列に区分
            LineData &= """出庫"","
            'C列にライン
            LineData &= Out_List(i).Line & ","
            'D列にベイト＆ガルプ
            LineData &= Out_List(i).Bait & ","
            'E列にサオ
            LineData &= Out_List(i).Rods & ","
            'F列にBaitcast_Reels
            LineData &= Out_List(i).B_Reels & ","
            'G列にSpinning_Reels
            LineData &= Out_List(i).S_Reels & ","
            'H列にCombos
            LineData &= Out_List(i).Combos & ","
            'I列にAccessory
            LineData &= Out_List(i).Accessory & ","
            'I列にAccessory
            'LineData &= Out_List(i).Accessory2 & ","
            'J列にHard_Lure
            LineData &= Out_List(i).Hard_Lure & ","
            'K列にBagApparel
            LineData &= Out_List(i).Bag & ","
            'L列にRodOEM
            LineData &= Out_List(i).Rod_OEM & ","
            'M列にRodParts
            LineData &= Out_List(i).Rod_Parts & ","
            'N列にReelParts
            LineData &= Out_List(i).ReelParts & ","
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)

            '棚卸データ
            '取得したデータをLineataに格納する。
            'A列に日付
            LineData = """" & Tanaoroshi_List(i).WorkDate & ""","
            'B列に区分
            LineData &= """棚卸"","
            'C列にライン
            LineData &= Tanaoroshi_List(i).Line & ","
            'D列にベイト＆ガルプ
            LineData &= Tanaoroshi_List(i).Bait & ","
            'E列にサオ
            LineData &= Tanaoroshi_List(i).Rods & ","
            'F列にBaitcast_Reels
            LineData &= Tanaoroshi_List(i).B_Reels & ","
            'G列にSpinning_Reels
            LineData &= Tanaoroshi_List(i).S_Reels & ","
            'H列にCombos
            LineData &= Tanaoroshi_List(i).Combos & ","
            'I列にAccessory
            LineData &= Tanaoroshi_List(i).Accessory & ","
            'I列にAccessory
            'LineData &= Tanaoroshi_List(i).Accessory2 & ","
            'J列にHard_Lure
            LineData &= Tanaoroshi_List(i).Hard_Lure & ","
            'K列にBagApparel
            LineData &= Tanaoroshi_List(i).Bag & ","
            'L列にRodOEM
            LineData &= Tanaoroshi_List(i).Rod_OEM & ","
            'M列にRodParts
            LineData &= Tanaoroshi_List(i).Rod_Parts & ","
            'N列にReelParts
            LineData &= Tanaoroshi_List(i).ReelParts & ","
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next

        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "在庫実績データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As StockLog_List = Nothing

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing

        Dim Date_Check_Result As DateTime

        Dim DataTotal As Integer = 0
        Dim DataNumTotal As Integer = 0

        Dim ChkItemCodeString As String = Nothing
        Dim ChkCommentString As String = Nothing

        Dim dtNow As DateTime = DateTime.Now

        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        '作業日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            Date_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("作業日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '作業日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("作業日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        ChkItemCodeString = Trim(TextBox1.Text)
        ChkCommentString = Trim(TextBox2.Text)

        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        'コメントに'が入力されていたらReplaceする。
        ChkCommentString = ChkCommentString.Replace("'", "''")

        '検索Function
        Result = GetStockLogSeach(ChkItemCodeString, Date_From, Date_To, CheckBox7.Checked, CheckBox8.Checked, CheckBox14.Checked, _
                            CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, _
                            CheckBox5.Checked, CheckBox6.Checked, CheckBox9.Checked, CheckBox10.Checked, _
                            CheckBox11.Checked, CheckBox12.Checked, CheckBox13.Checked, ChkCommentString, _
                            PLACE_ID, SearchResult, DataTotal, DataNumTotal, Result, ErrorMessage)
        If Result = False Then
            Label9.Text = "商品数： "
            Label10.Text = "総数： "
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "在庫実績データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "商品コード,商品名,数量,不良区分,ステータス,セット組、ばらしID,コメント,作業日時,倉庫"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If DataTotal = 0 Then
            MsgBox("データがありません。")
            Exit Sub
        End If

        'ﾌｧｲﾙ名の設定
        strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)
        'headerを設定
        strStreamWriter.WriteLine(Sheet1_Header)

        'データ件数分ループ
        For i = 0 To SearchResult.Length - 1
            '取得するデータをLineataに格納する。
            'A列に商品コード
            LineData = """" & SearchResult(i).I_CODE & ""","
            'B列に商品名
            LineData &= """" & SearchResult(i).I_NAME.Replace("""", """""") & ""","
            'C列に数量
            LineData &= SearchResult(i).NUM & " ,"
            'D列に不良区分
            LineData &= """" & SearchResult(i).I_STATUS & ""","
            'E列にステータス
            LineData &= """" & SearchResult(i).I_FLG & ""","
            'F列にセット組、ばらしID
            LineData &= SearchResult(i).PACKAGE_NO & " ,"
            'G列にコメント
            LineData &= """" & SearchResult(i).COMMENT & ""","
            'H列に作業日時
            LineData &= """" & SearchResult(i).U_DATE & ""","
            'I列に作業日時
            LineData &= """" & SearchResult(i).PLACE & """"
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "在庫実績データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

        Dim ToDate As Date

        ToDate = DateTimePicker1.Value.ToShortDateString()

        MaskedTextBox1.Text = Date.ParseExact(DateTimePicker1.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged

        Dim ToDate As Date

        ToDate = DateTimePicker2.Value.ToShortDateString()

        MaskedTextBox2.Text = Date.ParseExact(DateTimePicker2.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub
End Class