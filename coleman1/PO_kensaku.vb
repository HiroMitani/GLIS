Imports System.Windows.Forms
Imports System.Globalization

Public Class PO_kensaku
    'プロダクトラインID
    Dim PL_Id As Integer

    Public FormLord As Boolean = False

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    'False:検索ボタンを押したらFalse
    '入庫確定や変更など行ったら、再度検索を行わせる為のFlg
    Public nSearchFLg As Boolean = False

    Private Sub PO_kensaku_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim PODate_From As String = Nothing
        Dim PODate_To As String = Nothing
        Dim Date_Check_Result As DateTime

        '取込日用
        Dim INSDate_From As String = Nothing
        Dim INSDate_To As String = Nothing

        '発注No
        Dim ChkPO_NO As String
        'ベンダー
        Dim ChkVenderCode As String
        'ベンダー
        Dim ChkRemarks As String
        'ベンダー
        Dim ChkItem_Jan_Code As String
        '1:商品コード、2:JANコード
        Dim ItemJan_Type As Integer = 0

        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim SearchResult() As PO_Search_List = Nothing

        '一括チェックのチェックボックスをクリアする。
        CheckBox9.Checked = False

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        '検索ボタンクリックチェック
        nSearchFLg = False

        '検索条件のチェック

        '希望納期Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            PODate_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                PODate_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '希望納期Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            PODate_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                PODate_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '取込日Fromのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            INSDate_From = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                INSDate_From = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("取込日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '取込日Toのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            INSDate_To = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                INSDate_To = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("取込日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '発注№
        ChkPO_NO = Trim(TextBox2.Text).Replace("'", "''")
        'ベンダー
        ChkVenderCode = Trim(TextBox4.Text).Replace("'", "''")
        '備考
        ChkRemarks = Trim(TextBox8.Text).Replace("'", "''")
        '商品コード or JANコード
        ChkItem_Jan_Code = Trim(TextBox7.Text).Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If Trim(TextBox7.Text) <> "" Then

            If ComboBox1.Text = "商品コード" Then
                ItemJan_Type = 1
            ElseIf ComboBox1.Text = "JANコード" Then
                ItemJan_Type = 2
            Else
                MsgBox("商品コードかJANコードを選択してください。")
                ComboBox1.BackColor = Color.Salmon
                ComboBox1.Focus()
                Exit Sub
            End If
        End If

        'プロダクトライン
        Dim PL_ID As String = Nothing

        If ComboBox2.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox2.SelectedValue.ToString
        End If

        '検索Function
        Result = GetPoSeach(ChkPO_NO, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, _
                            ChkVenderCode, PL_ID, ItemJan_Type, ChkItem_Jan_Code, PODate_From, PODate_To, _
                            INSDate_From, INSDate_To, ChkRemarks, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)
        If Result = False Then
            '商品数、総数をクリア
            Label14.Text = "商品数："
            Label15.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If
        '商品数をセット
        Label14.Text = "商品数：" & Data_Total
        '総数をセット
        Label15.Text = "総数：" & Data_Num_Total


        '結果を元にDataGridViewに表示する。
        'DataGridへ入力したデータを挿入
        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                'ベンダー
                .Cells(1).Value = SearchResult(Count).PO_M_CODE

                '発注No
                .Cells(2).Value = SearchResult(Count).PO_NO
                '商品コード
                .Cells(3).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(4).Value = SearchResult(Count).I_NAME
                'ステータス
                .Cells(5).Value = SearchResult(Count).STATUS
                '発注数
                .Cells(6).Value = SearchResult(Count).PO_NUM
                '発注日
                .Cells(7).Value = SearchResult(Count).ORDER_DATE

                '納品希望日
                .Cells(8).Value = SearchResult(Count).PO_DATE


                '納品確定（入庫予定）数
                .Cells(9).Value = SearchResult(Count).IN_NUM
                '入庫予定日
                .Cells(10).Value = SearchResult(Count).IN_DATE
                '入庫確定数
                .Cells(11).Value = SearchResult(Count).FIX_NUM
                '入庫確定日
                .Cells(12).Value = SearchResult(Count).FIX_DATE
                'キャンセル数
                .Cells(13).Value = SearchResult(Count).CANCEL_NUM
                'キャンセル日
                .Cells(14).Value = SearchResult(Count).CANCEL_DATE
                '発注先名
                .Cells(15).Value = SearchResult(Count).PO_M_NAME
                'ベンダー名（工場名）
                .Cells(16).Value = SearchResult(Count).C_NAME
                'プロダクトライン名
                .Cells(17).Value = SearchResult(Count).PL_NAME
                '備考
                .Cells(18).Value = SearchResult(Count).REMARKS
                'ID
                .Cells(19).Value = SearchResult(Count).ID
                'I_ID
                .Cells(20).Value = SearchResult(Count).I_ID

            End With
            DataGridView1.Rows.Add(SR_List)
        Next

        nSearchFLg = False

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub PO_kensaku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim Disp_Title As String = "発注（PO）関連検索"

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

        '左4項目を固定(チェック、ベンダ、発注No,商品コード)
        DataGridView1.Columns(3).Frozen = True

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
        PL_Id = ComboBox1.SelectedValue

        ComboBox1.Text = "商品コード"

        FormLord = True

        nSearchFLg = False

    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount As Integer = 20

        Dim CheckDataCount As Integer = 0
        Dim CheckData() As Integer = Nothing
        Dim SelectCount As Integer = 0
        Dim LoopCount As Integer = 0
        For Each Row As DataGridViewRow In DataGridView1.SelectedRows
            ReDim Preserve CheckData(0 To SelectCount)
            CheckData(SelectCount) = Row.Index
            SelectCount += 1
        Next Row

        If SelectCount = 0 Then
            For i = 0 To loopcnt - 1
                If DataGridView1(0, i).Value = 1 Then
                    DataGridView1(0, i).Value = 0
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.White
                    Next
                ElseIf DataGridView1(0, i).Value = 0 Then
                    DataGridView1(0, i).Value = 1
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.PaleGreen
                    Next
                End If
            Next
        Else
            For LoopCount = 0 To CheckData.Length - 1
                If DataGridView1(0, CheckData(LoopCount)).Value = 1 Then
                    If CheckBox1.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.White
                        Next
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox1.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = 1
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.PaleGreen
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    'CurrentCellDirtyStateChangedイベントハンドラ
    Private Sub DataGridView1_CurrentCellDirtyStateChanged( _
            ByVal sender As Object, ByVal e As EventArgs) _
            Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.CurrentCellAddress.X = 0 AndAlso _
            DataGridView1.IsCurrentCellDirty Then
            'コミットする
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    'CellValueChangedイベントハンドラ
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, _
            ByVal e As DataGridViewCellEventArgs) _
            Handles DataGridView1.CellValueChanged
        Dim count As Integer
        Dim IndexCount As Integer = 20
        '列のインデックスを確認する
        If e.ColumnIndex = 0 Then
            If FormLord = True Then
                If DataGridView1(e.ColumnIndex, e.RowIndex).Value = 1 Then
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.PaleGreen
                    Next
                Else
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.White
                    Next
                End If
            End If
        End If
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

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker3.Value.ToShortDateString()

        MaskedTextBox3.Text = Date.ParseExact(DateTimePicker3.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)
    End Sub

    Private Sub DateTimePicker4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker4.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker4.Value.ToShortDateString()

        MaskedTextBox4.Text = Date.ParseExact(DateTimePicker4.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        Dim Check_Flg As Boolean = False

        Dim PO_Check_Flg As Boolean = False
        Dim PO_Check_Message As String = Nothing

        Dim PO_Data_Count As Integer = 0
        Dim PO_Check() As PO_In_List = Nothing
        Dim Data_Num_Total As Integer = 0

        Dim ListData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        If nSearchFLg = True Then
            MsgBox("入庫予定、削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        PO_Check_Flg = True

        '倉庫の一覧を取得する
        Result = GetPLACEList(ListData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(5).Value() = "全納品確定" _
               OrElse DataGridView1.Rows(Count).Cells(5).Value() = "全キャンセル" OrElse DataGridView1.Rows(Count).Cells(5).Value() = "一部キャンセル") Then
                PO_Check_Flg = False
            End If
        Next
        If PO_Check_Flg = False Then
            MsgBox("全納品確定、一部キャンセル、全キャンセルのデータにチェックされています。入庫予定は、発注済み、一部納品確定のみ入庫予定データの入力が行えます。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve PO_Check(0 To PO_Data_Count)
                'ベンダー
                PO_Check(PO_Data_Count).VENDER = DataGridView1.Rows(Count).Cells(1).Value()

                '発注No
                PO_Check(PO_Data_Count).PO_NO = DataGridView1.Rows(Count).Cells(2).Value()
                '商品コード
                PO_Check(PO_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(3).Value()
                '商品名
                PO_Check(PO_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(4).Value()
                'ステータス
                PO_Check(PO_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(5).Value()
                '発注数
                PO_Check(PO_Data_Count).PO_NUM = DataGridView1.Rows(Count).Cells(6).Value()
                '納品希望日
                PO_Check(PO_Data_Count).PO_DATE = DataGridView1.Rows(Count).Cells(8).Value()

                '納品確定（入庫予定）数
                PO_Check(PO_Data_Count).IN_NUM = DataGridView1.Rows(Count).Cells(9).Value()
                '入庫予定日
                PO_Check(PO_Data_Count).IN_DATE = DataGridView1.Rows(Count).Cells(10).Value()
                '入庫確定数
                PO_Check(PO_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(11).Value()
                '入庫確定日
                PO_Check(PO_Data_Count).FIX_DATE = DataGridView1.Rows(Count).Cells(12).Value()
                '発注キャンセル数
                PO_Check(PO_Data_Count).CANCELED_NUM = DataGridView1.Rows(Count).Cells(13).Value()
                '発注キャンセル日
                PO_Check(PO_Data_Count).CANCELED_DATE = DataGridView1.Rows(Count).Cells(14).Value()
                'ベンダ名
                PO_Check(PO_Data_Count).C_NAME = DataGridView1.Rows(Count).Cells(15).Value()
                'プロダクトライン
                PO_Check(PO_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(17).Value()
                '備考
                PO_Check(PO_Data_Count).REMARKS = DataGridView1.Rows(Count).Cells(18).Value()
                'ID
                PO_Check(PO_Data_Count).ID = DataGridView1.Rows(Count).Cells(19).Value()
                'I_ID
                PO_Check(PO_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(20).Value()

                PO_Data_Count += 1
            End If
        Next

        'PO_nyuukoyoteiのDataGridViewをクリア
        PO_nyuukoyotei.DataGridView1.Rows.Clear()

        '入庫予定のDataGridViewに表示する。
        For Count = 0 To PO_Check.Length - 1
            Dim IN_Schedule_Data_list As New DataGridViewRow
            IN_Schedule_Data_list.CreateCells(PO_nyuukoyotei.DataGridView1)
            With IN_Schedule_Data_list


                '発注No
                .Cells(0).Value = PO_Check(Count).PO_NO
                '商品コード
                .Cells(1).Value = PO_Check(Count).I_CODE
                '商品名
                .Cells(2).Value = PO_Check(Count).I_NAME
                '発注数
                .Cells(3).Value = PO_Check(Count).PO_NUM
                '納品希望日
                .Cells(4).Value = PO_Check(Count).PO_DATE
                '納品確定数(発注数-納品確定（入庫予定）数)
                .Cells(6).Value = PO_Check(Count).PO_NUM - PO_Check(Count).IN_NUM
                '倉庫（固定で１を表示）
                '.Cells(8).Value = ListData(0).NAME
                '納品確定（入庫予定）数
                .Cells(9).Value = PO_Check(Count).IN_NUM
                '入庫確定数
                .Cells(10).Value = PO_Check(Count).FIX_NUM

                '発注キャンセル数
                .Cells(11).Value = PO_Check(Count).CANCELED_NUM
                '発注キャンセル日
                .Cells(12).Value = PO_Check(Count).CANCELED_DATE
                'ベンダー
                .Cells(13).Value = PO_Check(Count).VENDER
                'ベンダー名
                .Cells(14).Value = PO_Check(Count).C_NAME
                'プロダクトライン
                .Cells(15).Value = PO_Check(Count).PL_NAME
                '備考
                .Cells(16).Value = PO_Check(Count).REMARKS

                'ID
                .Cells(17).Value = PO_Check(Count).ID
                'I_ID
                .Cells(18).Value = PO_Check(Count).I_ID
                'PLACE_ID
                '.Cells(18).Value = ListData(0).ID

            End With
            PO_nyuukoyotei.DataGridView1.Rows.Add(IN_Schedule_Data_list)
        Next

        PO_nyuukoyotei.Show()
        Me.Hide()

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim PODate_From As String = Nothing
        Dim PODate_To As String = Nothing
        Dim Date_Check_Result As DateTime

        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        '取込日用
        Dim INSDate_From As String = Nothing
        Dim INSDate_To As String = Nothing

        '発注No
        Dim ChkPO_NO As String
        'ベンダー
        Dim ChkVenderCode As String
        'ベンダー
        Dim ChkRemarks As String
        'ベンダー
        Dim ChkItem_Jan_Code As String
        '1:商品コード、2:JANコード
        Dim ItemJan_Type As Integer = 0

        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim SearchResult() As PO_Search_List = Nothing

        ''一括チェックのチェックボックスをクリアする。
        'CheckBox9.Checked = False

        ''検索ボタンクリックチェック
        'nSearchFLg = False

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        '検索条件のチェック

        '希望納期Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            PODate_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                PODate_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '希望納期Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            PODate_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                PODate_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '取込日Fromのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            INSDate_From = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                INSDate_From = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("取込日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '取込日Toのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            INSDate_To = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                INSDate_To = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("取込日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '発注№
        ChkPO_NO = Trim(TextBox2.Text).Replace("'", "''")
        'ベンダー
        ChkVenderCode = Trim(TextBox4.Text).Replace("'", "''")
        '備考
        ChkRemarks = Trim(TextBox8.Text).Replace("'", "''")
        '商品コード or JANコード
        ChkItem_Jan_Code = Trim(TextBox7.Text).Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If Trim(TextBox7.Text) <> "" Then

            If ComboBox1.Text = "商品コード" Then
                ItemJan_Type = 1
            ElseIf ComboBox1.Text = "JANコード" Then
                ItemJan_Type = 2
            Else
                MsgBox("商品コードかJANコードを選択してください。")
                ComboBox1.BackColor = Color.Salmon
                ComboBox1.Focus()
                Exit Sub
            End If

        End If

        'プロダクトライン
        Dim PL_ID As String = Nothing

        If ComboBox2.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox2.SelectedValue.ToString
        End If

        '検索Function
        Result = GetPoSeach(ChkPO_NO, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, _
                            ChkVenderCode, PL_ID, ItemJan_Type, ChkItem_Jan_Code, PODate_From, PODate_To, _
                            INSDate_From, INSDate_To, ChkRemarks, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)
        If Result = False Then
            '商品数、総数をクリア
            Label14.Text = "商品数："
            Label15.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "発注関連データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "ベンダコード,発注No,商品コード,商品名,ステータス,発注数,発注日,希望納期,納品確定（入庫予定）数,入庫予定日,入庫確定数,入庫確定日,発注キャンセル数,発注キャンセル日,ベンダ名,工場名,プロダクトライン,備考"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If Data_Total = 0 Then
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
            'A列にベンダ
            LineData = """" & SearchResult(i).PO_M_CODE & ""","
            'B列に発注No
            LineData &= """" & SearchResult(i).PO_NO & ""","
            'C列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'D列に商品名
            LineData &= """" & SearchResult(i).I_NAME & ""","
            'E列にステータス
            LineData &= """" & SearchResult(i).STATUS & ""","
            'F列に発注数
            LineData &= "" & SearchResult(i).PO_NUM & ","
            'G列に発注日
            LineData &= "" & SearchResult(i).ORDER_DATE & ","
            'H列に希望納期
            LineData &= "" & SearchResult(i).PO_DATE & ","
            'I列に納品確定（入庫予定）数
            LineData &= "" & SearchResult(i).IN_NUM & ","
            'J列に入庫予定日
            LineData &= """" & SearchResult(i).IN_DATE & ""","
            'K列に入庫確定数
            LineData &= "" & SearchResult(i).FIX_NUM & ","
            'L列に入庫確定日
            LineData &= """" & SearchResult(i).FIX_DATE & ""","
            'M列に発注キャンセル数
            LineData &= "" & SearchResult(i).CANCEL_NUM & ","
            'N列に発注キャンセル日
            LineData &= """" & SearchResult(i).CANCEL_DATE & ""","
            'O列にベンダ名
            LineData &= """" & SearchResult(i).PO_M_NAME & ""","
            'P列に工場名
            LineData &= """" & SearchResult(i).C_NAME & ""","
            'Q列にプロダクトライン
            LineData &= """" & SearchResult(i).PL_NAME & ""","
            'R列に備考
            LineData &= """" & SearchResult(i).REMARKS & ""","
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "CSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim Check_Flg As Boolean = False

        Dim PO_Check_Flg As Boolean = False
        Dim PO_Check_Message As String = Nothing

        Dim PO_Data_Count As Integer = 0
        Dim PO_Check() As PO_In_List = Nothing
        Dim Data_Num_Total As Integer = 0

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        If nSearchFLg = True Then
            MsgBox("入庫予定、削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        PO_Check_Flg = True
        'チェックされた商品の中でピッキング済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(5).Value() = "一部納品確定" _
                OrElse DataGridView1.Rows(Count).Cells(5).Value() = "全納品確定" OrElse DataGridView1.Rows(Count).Cells(5).Value() = "一部キャンセル" _
                OrElse DataGridView1.Rows(Count).Cells(5).Value() = "全キャンセル") Then
                PO_Check_Flg = False
            End If
        Next
        If PO_Check_Flg = False Then
            MsgBox("一部納品確定、全納品確定、一部キャンセル、全キャンセルのデータにチェックされています。削除は、発注済みデータのみ可能です。")
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("発注データを削除してもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve PO_Check(0 To PO_Data_Count)
                'ID
                PO_Check(PO_Data_Count).ID = DataGridView1.Rows(Count).Cells(18).Value()

                PO_Data_Count += 1
            End If
        Next

        '削除Function
        Result = Del_Po(PO_Check, Result, ErrorMessage)
        If Result = False Then

            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("削除が完了しました。")

        nSearchFLg = True

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim Check_Flg As Boolean = False

        Dim PO_Check_Flg As Boolean = False
        Dim PO_Check_Message As String = Nothing

        Dim PO_Data_Count As Integer = 0
        Dim PO_Check() As PO_In_List = Nothing
        Dim Data_Num_Total As Integer = 0
        If nSearchFLg = True Then
            MsgBox("入庫予定、削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        PO_Check_Flg = True

        'チェックされた商品の中で全納品確定、全キャンセルがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(5).Value() = "全納品確定" _
               OrElse DataGridView1.Rows(Count).Cells(5).Value() = "全キャンセル" OrElse DataGridView1.Rows(Count).Cells(5).Value() = "一部キャンセル") Then
                PO_Check_Flg = False
            End If
        Next
        If PO_Check_Flg = False Then
            MsgBox("全納品確定、一部キャンセル、全キャンセルのデータにチェックされています。発注キャンセルは、発注済み、一部納品確定のみ発注キャンセルが行えます。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve PO_Check(0 To PO_Data_Count)
                'ベンダー
                PO_Check(PO_Data_Count).VENDER = DataGridView1.Rows(Count).Cells(1).Value()

                '発注No
                PO_Check(PO_Data_Count).PO_NO = DataGridView1.Rows(Count).Cells(2).Value()
                '商品コード
                PO_Check(PO_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(3).Value()
                '商品名
                PO_Check(PO_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(4).Value()
                'ステータス
                PO_Check(PO_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(5).Value()
                '発注数
                PO_Check(PO_Data_Count).PO_NUM = DataGridView1.Rows(Count).Cells(6).Value()
                '納品希望日
                PO_Check(PO_Data_Count).PO_DATE = DataGridView1.Rows(Count).Cells(8).Value()
 

                '納品確定（入庫予定）数
                PO_Check(PO_Data_Count).IN_NUM = DataGridView1.Rows(Count).Cells(9).Value()
                '入庫予定日
                PO_Check(PO_Data_Count).IN_DATE = DataGridView1.Rows(Count).Cells(10).Value()
                '入庫確定数
                PO_Check(PO_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(11).Value()
                '入庫確定日
                PO_Check(PO_Data_Count).FIX_DATE = DataGridView1.Rows(Count).Cells(12).Value()
                '発注キャンセル数
                PO_Check(PO_Data_Count).CANCELED_NUM = DataGridView1.Rows(Count).Cells(13).Value()
                '発注キャンセル日
                PO_Check(PO_Data_Count).CANCELED_DATE = DataGridView1.Rows(Count).Cells(14).Value()
                'ベンダ名
                PO_Check(PO_Data_Count).C_NAME = DataGridView1.Rows(Count).Cells(15).Value()
                'プロダクトライン
                PO_Check(PO_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(17).Value()
                '備考
                PO_Check(PO_Data_Count).REMARKS = DataGridView1.Rows(Count).Cells(18).Value()
                'ID
                PO_Check(PO_Data_Count).ID = DataGridView1.Rows(Count).Cells(19).Value()
                'I_ID
                PO_Check(PO_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(20).Value()

                PO_Data_Count += 1
            End If
        Next

        'PO_cancelのDataGridViewをクリア
        PO_cancel.DataGridView1.Rows.Clear()

        'DataGridViewに表示する。
        For Count = 0 To PO_Check.Length - 1
            Dim Cancel_Data_list As New DataGridViewRow
            Cancel_Data_list.CreateCells(PO_cancel.DataGridView1)
            With Cancel_Data_list


                '発注No
                .Cells(0).Value = PO_Check(Count).PO_NO
                '商品コード
                .Cells(1).Value = PO_Check(Count).I_CODE
                '商品名
                .Cells(2).Value = PO_Check(Count).I_NAME
                '発注数
                .Cells(3).Value = PO_Check(Count).PO_NUM
                '納品希望日
                .Cells(4).Value = PO_Check(Count).PO_DATE
                '発注キャンセル(発注数 - 納品確定（入庫予定数） - 過去キャンセル数)
                .Cells(5).Value = PO_Check(Count).PO_NUM - PO_Check(Count).IN_NUM - PO_Check(Count).CANCELED_NUM
                '納品確定（入庫予定）数
                .Cells(6).Value = PO_Check(Count).IN_NUM
                '入庫予定日
                .Cells(7).Value = PO_Check(Count).IN_DATE
                '入庫確定数
                .Cells(8).Value = PO_Check(Count).FIX_NUM
                '過去発注キャンセル数
                .Cells(9).Value = PO_Check(Count).CANCELED_NUM
                '発注キャンセル日
                .Cells(10).Value = PO_Check(Count).CANCELED_DATE
                'ベンダコード
                .Cells(11).Value = PO_Check(Count).VENDER
                'ベンダ名
                .Cells(12).Value = PO_Check(Count).C_NAME
                'プロダクトライン
                .Cells(13).Value = PO_Check(Count).PL_NAME
                '備考
                .Cells(14).Value = PO_Check(Count).REMARKS

                'ID
                .Cells(15).Value = PO_Check(Count).ID
                'I_ID
                .Cells(16).Value = PO_Check(Count).I_ID

            End With
            PO_cancel.DataGridView1.Rows.Add(Cancel_Data_list)
        Next

        PO_cancel.Show()
        Me.Hide()



    End Sub
End Class