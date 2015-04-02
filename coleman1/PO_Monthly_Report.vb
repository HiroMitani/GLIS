Imports System.Windows.Forms
Imports System.Globalization

Public Class PO_Monthly_Report

    Dim PL_Id As Integer

    Dim PLACE_ID As String

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Private Sub PO_Monthly_Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "月別検索"

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

        '左3項目を固定(チェック、商品コード、商品名)
        DataGridView1.Columns(2).Frozen = True

        'DataGridViewを書き込み禁止にする。
        DataGridView1.ReadOnly = True

        'プロダクトラインコード情報取得
        Dim PLTable As New DataTable()
        PLTable.Columns.Add("ID", GetType(String))
        PLTable.Columns.Add("NAME", GetType(String))
        'プロダクトラインの情報を取得する。
        PLResult = GetPLList(PLList, PLResult, PLErrorMessage)
        If PLResult = False Then
            MsgBox(PLErrorMessage)
            Exit Sub
        End If

        For Count = 0 To PLList.Length - 1
            '新しい行を作成
            Dim row As DataRow = PLTable.NewRow()

            '各列に値をセット
            row("ID") = PLList(Count).ID
            row("NAME") = PLList(Count).NAME

            'DataTableに行を追加
            PLTable.Rows.Add(row)
        Next

        PLTable.AcceptChanges()
        'コンボボックスのDataSourceにDataTableを割り当てる
        ComboBox2.DataSource = PLTable

        '表示される値はDataTableのNAME列
        ComboBox2.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox2.ValueMember = "ID"

        '初期値をセット
        ComboBox2.SelectedIndex = -1
        'PL_Id = ComboBox1.SelectedValue.ToString()
        PL_Id = ComboBox2.SelectedValue

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

        ComboBox1.Text = "商品コード"
    End Sub

    Private Sub PO_Monthly_Report_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker1.Value.ToShortDateString()

        MaskedTextBox1.Text = ToDate.ToString("yyyy/MM")

    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker2.Value.ToShortDateString()
        MaskedTextBox2.Text = ToDate.ToString("yyyy/MM")

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        If RadioButton3.Checked = True Then
            RadioButton4.Checked = False
            RadioButton5.Checked = False
        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged

        If RadioButton4.Checked = True Then
            RadioButton3.Checked = False
            RadioButton5.Checked = False
        End If

    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged

        If RadioButton5.Checked = True Then
            RadioButton3.Checked = False
            RadioButton4.Checked = False
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Month As Integer

        Dim ItemName As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ItemJan_Flg As Integer = 0

        Dim ZeroDataFlg As Integer = 0

        Dim Datatype As Integer = 0

        Dim Date_From As Date
        Dim Date_To As Date
        Dim Date_From_String As String
        Dim Date_To_String As String

        Dim Date_Check_Result As DateTime

        Dim SearchResult() As Monthly_Report_Search_List = Nothing

        Dim TMP_Date As Date
        Dim TMP_Date_String As String
        Dim MonthCount As Integer = 0

        Dim checkflg As Integer = 0

        'プロダクトライン
        Dim PL_ID As String = Nothing

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        'DataGridViewのカラムを全てクリアする。
        DataGridView1.Columns.Clear()
        'DataGridViewにカラムを追加する。
        DataGridView1.Columns.Add("column3", "プロダクトライン")
        DataGridView1.Columns.Add("column1", "商品コード")
        DataGridView1.Columns.Add("column2", "商品名")
        DataGridView1.Columns.Add("column4", "JANコード")
        DataGridView1.Columns(2).Width = 150

        '日付として正しいかチェック
        '対象年月Fromのチェック
        If MaskedTextBox1.Text = "    /" Then
            MsgBox("対象年月はFrom、Toが必須項目です。")
            MaskedTextBox1.BackColor = Color.Salmon
            Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("対象年月の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '対象年月Toのチェック
        If MaskedTextBox2.Text = "    /" Then
            MsgBox("対象年月はFrom、Toが必須項目です。")
            MaskedTextBox2.BackColor = Color.Salmon
            Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("対象年月の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        If ComboBox2.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox2.SelectedValue.ToString
        End If


        '検索日付のチェック
        '13ヶ月以上の期間を指定していたらメッセージ表示
        Month = DateDiff(DateInterval.Month, Date_From, Date_To)
        If Month > 12 Then
            MsgBox("検索期間は12ヶ月以内を指定してください。")
        End If

        '検索データが発注数の場合
        If RadioButton3.Checked = True Then
            Datatype = 1
        ElseIf RadioButton4.Checked = True Then
            '検索データが入庫確定数の場合
            Datatype = 2
        ElseIf RadioButton5.Checked = True Then
            '検索データが出庫確定数の場合
            Datatype = 3
        End If

        ChkItemJanString = Trim(TextBox1.Text)
        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        '商品名に'が入力されていたらReplaceする。
        ItemName = Trim(TextBox2.Text)
        ItemName = ItemName.Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox1.Text = "商品コード" Then
            ItemJan_Flg = 1
        ElseIf ComboBox1.Text = "JANコード" Then
            ItemJan_Flg = 2
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox1.BackColor = Color.Salmon
            ComboBox1.Focus()
            Exit Sub
        End If

        '数量0のデータを表示する or しない。
        If RadioButton1.Checked = True Then
            ZeroDataFlg = 1
        ElseIf RadioButton2.Checked = True Then
            ZeroDataFlg = 2
        End If

        Date_From = Date.Parse(MaskedTextBox1.Text)
        Date_To = Date.Parse(MaskedTextBox2.Text)

        Date_From_String = MaskedTextBox1.Text
        Date_To_String = MaskedTextBox2.Text

        '検索Function
        Result = Get_Monthly_Report(Date_From_String, Date_To_String, ItemJan_Flg, ChkItemJanString, ItemName, _
                            PL_ID, ZeroDataFlg, Datatype, PLACE_ID, SearchResult, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '検索タイプにより、DataGridViewのカラムを追加する。
        If Datatype = 1 Then
            'DataGridView1.Columns.Add("column5", "納品確定（入庫予定）数")
            'DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'DataGridView1.Columns(4).Width = 160
        ElseIf Datatype = 3 Then
            DataGridView1.Columns.Add("column5", "発注数")
            DataGridView1.Columns.Add("column6", "在庫数")
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 80
        End If

        TMP_Date_String = Date_From_String
        While checkflg = 0
            TMP_Date = Date.Parse(Date_From).AddMonths(MonthCount)
            TMP_Date_String = TMP_Date.ToString("yyyy/MM")
            If TMP_Date_String > Date_To_String Then
                checkflg = 1
            End If
            MonthCount += 1
        End While

        For i = 0 To MonthCount - 2
            TMP_Date = Date.Parse(Date_From).AddMonths(i)

            TMP_Date_String = TMP_Date.ToString("yyyy/MM")
            DataGridView1.Columns.Add("clmName" & MonthCount, TMP_Date_String)
            If Datatype = 1 Then
                'DataGridView1.Columns(i + 5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'DataGridView1.Columns(i + 5).Width = 80
            ElseIf Datatype = 2 Then
                DataGridView1.Columns(i + 4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridView1.Columns(i + 4).Width = 80
            ElseIf Datatype = 3 Then
                DataGridView1.Columns(i + 6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridView1.Columns(i + 6).Width = 80
            End If
        Next

        '結果を元にDataGridViewに表示する。
        'DataGridへ入力したデータを挿入
        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                'プロダクトライン名
                .Cells(0).Value = SearchResult(Count).PL_NAME
                '商品コード
                .Cells(1).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(2).Value = SearchResult(Count).I_NAME
                'JANコード
                .Cells(3).Value = SearchResult(Count).JAN


                If Datatype = 1 Or Datatype = 2 Then
                    If MonthCount > 1 Then
                        .Cells(4).Value = SearchResult(Count).Month1
                    End If
                    If MonthCount > 2 Then
                        .Cells(5).Value = SearchResult(Count).Month2
                    End If
                    If MonthCount > 3 Then
                        .Cells(6).Value = SearchResult(Count).Month3
                    End If
                    If MonthCount > 4 Then
                        .Cells(7).Value = SearchResult(Count).Month4
                    End If
                    If MonthCount > 5 Then
                        .Cells(8).Value = SearchResult(Count).Month5
                    End If
                    If MonthCount > 6 Then
                        .Cells(9).Value = SearchResult(Count).Month6
                    End If
                    If MonthCount > 7 Then
                        .Cells(10).Value = SearchResult(Count).Month7
                    End If
                    If MonthCount > 8 Then
                        .Cells(11).Value = SearchResult(Count).Month8
                    End If
                    If MonthCount > 9 Then
                        .Cells(12).Value = SearchResult(Count).Month9
                    End If
                    If MonthCount > 10 Then
                        .Cells(13).Value = SearchResult(Count).Month10
                    End If
                    If MonthCount > 11 Then
                        .Cells(14).Value = SearchResult(Count).Month11
                    End If
                    If MonthCount > 12 Then
                        .Cells(15).Value = SearchResult(Count).Month12
                    End If
                ElseIf Datatype = 3 Then
                    .Cells(4).Value = SearchResult(Count).ORDER_NUM
                    .Cells(5).Value = SearchResult(Count).STOCK_NUM

                    If MonthCount > 1 Then
                        .Cells(6).Value = SearchResult(Count).Month1
                    End If
                    If MonthCount > 2 Then
                        .Cells(7).Value = SearchResult(Count).Month2
                    End If
                    If MonthCount > 3 Then
                        .Cells(8).Value = SearchResult(Count).Month3
                    End If
                    If MonthCount > 4 Then
                        .Cells(9).Value = SearchResult(Count).Month4
                    End If
                    If MonthCount > 5 Then
                        .Cells(10).Value = SearchResult(Count).Month5
                    End If
                    If MonthCount > 6 Then
                        .Cells(11).Value = SearchResult(Count).Month6
                    End If
                    If MonthCount > 7 Then
                        .Cells(12).Value = SearchResult(Count).Month7
                    End If
                    If MonthCount > 8 Then
                        .Cells(13).Value = SearchResult(Count).Month8
                    End If
                    If MonthCount > 9 Then
                        .Cells(14).Value = SearchResult(Count).Month9
                    End If
                    If MonthCount > 10 Then
                        .Cells(15).Value = SearchResult(Count).Month10
                    End If
                    If MonthCount > 11 Then
                        .Cells(16).Value = SearchResult(Count).Month11
                    End If
                    If MonthCount > 12 Then
                        .Cells(17).Value = SearchResult(Count).Month12
                    End If

                End If


            End With
            DataGridView1.Rows.Add(SR_List)
        Next


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Month As Integer

        Dim Date_From As Date
        Dim Date_To As Date
        Dim Date_From_String As String
        Dim Date_To_String As String

        Dim Date_Check_Result As DateTime

        Dim SearchResult() As Monthly_Report_Search_List = Nothing

        Dim ItemName As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ItemJan_Flg As Integer = 0

        Dim ZeroDataFlg As Integer = 0

        Dim Datatype As Integer = 0

        Dim TMP_Date As Date
        Dim TMP_Date_String As String
        Dim MonthCount As Integer = 0

        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        Dim Checkflg As Integer = 0

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        '日付として正しいかチェック
        '対象年月Fromのチェック
        If MaskedTextBox1.Text = "    /  " Then
            '未入力なら""を格納
            Date_From = ""

            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("対象年月の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '対象年月Toのチェック
        If MaskedTextBox2.Text = "    /  " Then
            '未入力なら""を格納
            Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("対象年月の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '検索日付のチェック
        '13ヶ月以上の期間を指定していたらメッセージ表示
        Month = DateDiff(DateInterval.Month, Date_From, Date_To)
        If Month > 12 Then
            MsgBox("検索期間は12ヶ月以内を指定してください。")
        End If

        '検索データが発注数の場合
        If RadioButton3.Checked = True Then
            Datatype = 1
        ElseIf RadioButton4.Checked = True Then
            '検索データが入庫確定数の場合
            Datatype = 2
        ElseIf RadioButton5.Checked = True Then
            '検索データが出庫確定数の場合
            Datatype = 3
        End If

        ChkItemJanString = Trim(TextBox1.Text)
        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        '商品名に'が入力されていたらReplaceする。
        ItemName = Trim(TextBox2.Text)
        ItemName = ItemName.Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox1.Text = "商品コード" Then
            ItemJan_Flg = 1
        ElseIf ComboBox1.Text = "JANコード" Then
            ItemJan_Flg = 2
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox1.BackColor = Color.Salmon
            ComboBox1.Focus()
            Exit Sub
        End If

        '数量0のデータを表示する or しない。
        If RadioButton1.Checked = True Then
            ZeroDataFlg = 1
        ElseIf RadioButton2.Checked = True Then
            ZeroDataFlg = 2
        End If

        Date_From = Date.Parse(MaskedTextBox1.Text)
        Date_To = Date.Parse(MaskedTextBox2.Text)

        Date_From_String = MaskedTextBox1.Text
        Date_To_String = MaskedTextBox2.Text

        '検索Function
        Result = Get_Monthly_Report(Date_From_String, Date_To_String, ItemJan_Flg, ChkItemJanString, ItemName, _
                            PL_Id, ZeroDataFlg, Datatype, PLACE_ID, SearchResult, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "月別検索データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        Dim Sheet1_Header As String = Nothing
        'Header設定
        If Datatype = 1 Or Datatype = 2 Then
            Sheet1_Header = "プロダクトライン,商品コード,商品名,JANコード"

        ElseIf Datatype = 3 Then
            Sheet1_Header = "プロダクトライン,商品コード,商品名,JANコード,発注数,在庫数"
        End If

        TMP_Date_String = Date_From_String
        While checkflg = 0
            TMP_Date = Date.Parse(Date_From).AddMonths(MonthCount)
            TMP_Date_String = TMP_Date.ToString("yyyy/MM")
            If TMP_Date_String > Date_To_String Then
                checkflg = 1
            End If
            MonthCount += 1
        End While

        For i = 0 To MonthCount - 2
            TMP_Date = Date.Parse(Date_From).AddMonths(i)

            TMP_Date_String = TMP_Date.ToString("yyyy/MM")
            Sheet1_Header = Sheet1_Header & "," & TMP_Date_String
        Next


        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        'If SearchResult.Length - 1 = 0 Then
        If SearchResult.Length = 0 Then
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
            'A列にプロダクトライン
            LineData = """" & SearchResult(i).PL_NAME & ""","
            'B列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'C列に商品名
            LineData &= """" & SearchResult(i).I_NAME & ""","
            'D列にJANコード
            LineData &= """" & SearchResult(i).JAN & ""","

            If Datatype = 1 Or Datatype = 2 Then
                If MonthCount > 1 Then
                    LineData &= """" & SearchResult(i).Month1 & ""","
                End If
                If MonthCount > 2 Then
                    LineData &= """" & SearchResult(i).Month2 & ""","
                End If
                If MonthCount > 3 Then
                    LineData &= """" & SearchResult(i).Month3 & ""","
                End If
                If MonthCount > 4 Then
                    LineData &= """" & SearchResult(i).Month4 & ""","
                End If
                If MonthCount > 5 Then
                    LineData &= """" & SearchResult(i).Month5 & ""","
                End If
                If MonthCount > 6 Then
                    LineData &= """" & SearchResult(i).Month6 & ""","
                End If
                If MonthCount > 7 Then
                    LineData &= """" & SearchResult(i).Month7 & ""","
                End If
                If MonthCount > 8 Then
                    LineData &= """" & SearchResult(i).Month8 & ""","
                End If
                If MonthCount > 9 Then
                    LineData &= """" & SearchResult(i).Month9 & ""","
                End If
                If MonthCount > 10 Then
                    LineData &= """" & SearchResult(i).Month10 & ""","
                End If
                If MonthCount > 11 Then
                    LineData &= """" & SearchResult(i).Month11 & ""","
                End If
                If MonthCount > 12 Then
                    LineData &= """" & SearchResult(i).Month12 & ""","
                End If
            ElseIf Datatype = 3 Then
                LineData &= """" & SearchResult(i).ORDER_NUM & ""","
                LineData &= """" & SearchResult(i).STOCK_NUM & ""","


                If MonthCount > 1 Then
                    LineData &= """" & SearchResult(i).Month1 & ""","
                End If
                If MonthCount > 2 Then
                    LineData &= """" & SearchResult(i).Month2 & ""","
                End If
                If MonthCount > 3 Then
                    LineData &= """" & SearchResult(i).Month3 & ""","
                End If
                If MonthCount > 4 Then
                    LineData &= """" & SearchResult(i).Month4 & ""","
                End If
                If MonthCount > 5 Then
                    LineData &= """" & SearchResult(i).Month5 & ""","
                End If
                If MonthCount > 6 Then
                    LineData &= """" & SearchResult(i).Month6 & ""","
                End If
                If MonthCount > 7 Then
                    LineData &= """" & SearchResult(i).Month7 & ""","
                End If
                If MonthCount > 8 Then
                    LineData &= """" & SearchResult(i).Month8 & ""","
                End If
                If MonthCount > 9 Then
                    LineData &= """" & SearchResult(i).Month9 & ""","
                End If
                If MonthCount > 10 Then
                    LineData &= """" & SearchResult(i).Month10 & ""","
                End If
                If MonthCount > 11 Then
                    LineData &= """" & SearchResult(i).Month11 & ""","
                End If
                If MonthCount > 12 Then
                    LineData &= """" & SearchResult(i).Month12 & ""","
                End If

            End If
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "CSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub
End Class