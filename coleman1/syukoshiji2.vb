Imports System.Windows.Forms
Imports System.Globalization

Public Class syukoshiji2

    Dim I_Id As String
    Dim PLACE_ID As String

    Public FormLord As Boolean = False

    'False:検索ボタンを押したらFalse
    '確定や変更など行ったら、再度検索を行わせる為のFlg
    Public S_shijiSearchFLg As Boolean = False

    '出荷伝票の出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Check_Flg As Boolean = False
        Dim Out_Check_Flg As Boolean = True

        Dim SearchConditions() As OutShipping_Search_List = Nothing
        Dim SearchResult() As OutShipping_Search_List = Nothing

        Dim Data_Num_Total As Integer = 0

        Dim Data_Count As Integer = 0

        Dim I_STATUS As String = Nothing
        Dim S_STATUS As Integer

        If S_shijiSearchFLg = True Then
            MsgBox("出荷予定変更、出荷指示登録を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True Then
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Out_Check_Flg = True
        'チェックされた商品の中で出荷指示済みデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(14).Value() = "出荷指示登録済") Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("出荷指示登録済のデータにチェックされています。出荷予定データのみ予定変更が行えます。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータの区分、ステータス、商品IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True Then
                ReDim Preserve SearchConditions(0 To Data_Count)
                '区分
                SearchConditions(Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(9).Value()
                'ステータス
                SearchConditions(Data_Count).STATUS = DataGridView1.Rows(Count).Cells(11).Value()
                '商品ID
                SearchConditions(Data_Count).I_ID = DataGridView1.Rows(Count).Cells(14).Value()
                Data_Count += 1
            End If
        Next

        '検索条件から必須項目を取得
        If RadioButton1.Checked = True Then
            S_STATUS = 1
        Else
            S_STATUS = 2
        End If

        'サマリーで表示している為、区分、出荷ステータス、PLACE_ID（倉庫ID）、I_ID（商品ID）を使い、明細データを取得する。
        '出荷予定データ検索Function
        Result = GetOutShipping_Plan_DetailSearch(S_STATUS, PLACE_ID, SearchConditions, SearchResult, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'syukohenko2のDataGridViewをクリア
        syukohenko2.DataGridView1.Rows.Clear()

        '予定変更のDataGridViewに表示する。
        For Count = 0 To SearchResult.Length - 1
            Dim Out_Data_list As New DataGridViewRow
            Out_Data_list.CreateCells(syukohenko2.DataGridView1)
            With Out_Data_list
                '出荷先コード
                .Cells(0).Value = SearchResult(Count).C_CODE
                '出荷先名
                .Cells(1).Value = SearchResult(Count).D_NAME
                '商品コード
                .Cells(2).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(3).Value = SearchResult(Count).I_NAME
                '現出荷希望数（確認用）
                .Cells(4).Value = SearchResult(Count).NUM
                '出荷希望数（変更用）
                .Cells(5).Value = SearchResult(Count).NUM
                '現出荷指示予定数（確認用）
                .Cells(6).Value = SearchResult(Count).PLAN_NUM
                '出荷指示予定数（変更用）
                .Cells(7).Value = SearchResult(Count).PLAN_NUM
                '出荷指示済数
                .Cells(8).Value = SearchResult(Count).FIX_NUM
                '現在庫数
                .Cells(9).Value = SearchResult(Count).STOCK_NUM
                '差し引き在庫数
                '.Cells(6).Value = Out_Check(Count).
                '納品単価
                .Cells(11).Value = SearchResult(Count).D_UNIT_PRICE

                '客先発注No
                .Cells(12).Value = SearchResult(Count).CUSTOMER_ORDER_NO
                '出荷予定日
                .Cells(13).Value = SearchResult(Count).OUT_DATE
                '出荷倉庫
                .Cells(14).Value = SearchResult(Count).P_NAME
                '区分
                .Cells(15).Value = SearchResult(Count).I_STATUS

                '出荷ステータス
                .Cells(16).Value = SearchResult(Count).S_STATUS
                'ステータス
                .Cells(17).Value = SearchResult(Count).STATUS
                'コメント１
                .Cells(18).Value = SearchResult(Count).COMMENT1
                'コメント２
                .Cells(19).Value = SearchResult(Count).COMMENT2
                'ID
                .Cells(20).Value = SearchResult(Count).ID
                'P_ID
                .Cells(21).Value = SearchResult(Count).P_ID
                'C_ID
                .Cells(22).Value = SearchResult(Count).C_ID
                'I_ID
                .Cells(23).Value = SearchResult(Count).I_ID

            End With
            syukohenko2.DataGridView1.Rows.Add(Out_Data_list)
        Next

        syukohenko2.Show()
        Me.Hide()
    End Sub

    Private Sub syukoshiji2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Doc_Result As Boolean = True
        Dim PlaceData() As Place_List = Nothing
        Dim CustomerData() As C_List = Nothing
        Dim ItemData() As Item_List = Nothing

        Dim Disp_Title As String = "出荷指示"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label9.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '入力欄クリア

        TextBox3.Text = ""

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
        ComboBox1.DataSource = PlaceTable

        '表示される値はDataTableのNAME列
        ComboBox1.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox1.ValueMember = "ID"
        '初期値をセット(倉庫データの1件目）
        ComboBox1.SelectedIndex = 0
        PLACE_ID = ComboBox1.SelectedValue.ToString()

        '左1項目を固定(チェック)
        DataGridView1.Columns(0).Frozen = True

        ComboBox4.Text = "良品"
        ComboBox2.Text = "通常出荷"

        FormLord = True
    End Sub
    Private Sub syukoshiji2_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Check_Flg As Boolean = False
        Dim Out_Check_Flg As Boolean = True
        Dim S_STATUS_Check_Flg As Boolean = True
        Dim STATUS_Check_Flg As Boolean = True
        Dim RegistData() As OutShipping_Search_List = Nothing

        Dim Minus_Check_Flg As Boolean = True

        Dim Data_Count As Integer

        Dim dtNow As DateTime
        dtNow = DateTime.Now

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True Then
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'チェックされた商品の中で出荷指示登録済がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True And DataGridView1.Rows(Count).Cells(9).Value() = "出荷指示登録済" Then
                S_STATUS_Check_Flg = False
            End If
        Next
        If S_STATUS_Check_Flg = False Then
            MsgBox("出荷指示登録済のデータがチェックされています。出荷指示登録済のデータは出荷指示登録できません。")
            Exit Sub
        End If
        '伝票出力のみのデータにチェックされていればエラー。
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True And DataGridView1.Rows(Count).Cells(10).Value() = "伝票出力のみ" Then
                STATUS_Check_Flg = False
            End If
        Next
        If STATUS_Check_Flg = False Then
            MsgBox("伝票出力のみのデータがチェックされています。伝票出力のみのデータは出荷指示登録できません。")
            Exit Sub
        End If


        '差引在庫数がマイナスならエラー。
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(11).Value() = "通常出荷" Then
                If DataGridView1.Rows(Count).Cells(0).Value() = True And DataGridView1.Rows(Count).Cells(7).Value() < 0 Then
                    Minus_Check_Flg = False
                End If
            End If
        Next
        If Minus_Check_Flg = False Then
            MsgBox("差引在庫数がマイナスのデータは出荷指示登録が行えません。予定変更画面で出荷指示予定数を調整をしてください。")
            Exit Sub
        End If


        Dim Now_Date As Date = dtNow.ToString("yyyy年MM月dd日")
        Dim OUT_Date As Date = Trim(DateTimePicker1.Text)

        If OUT_Date < Now_Date Then
            MsgBox("出荷予定日は本日以降を選択してください。")
            DateTimePicker1.Focus()
            Exit Sub
        End If


        'ダイアログ設定
        Dim Check_result As DialogResult = MessageBox.Show("出荷指示データに登録してもよろしいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)
        '何が選択されたか調べる()
        If Check_result = DialogResult.No Then
            '「いいえ」が選択された時 
            Exit Sub
        End If



        '必要データを配列に格納
        'DataGridViewからチェックされたデータの区分、ステータス、商品IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True Then
                ReDim Preserve RegistData(0 To Data_Count)
                '区分
                RegistData(Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(9).Value()
                'ステータス
                RegistData(Data_Count).STATUS = DataGridView1.Rows(Count).Cells(11).Value()
                '商品ID
                RegistData(Data_Count).I_ID = DataGridView1.Rows(Count).Cells(14).Value()
                'ID
                RegistData(Data_Count).ID = DataGridView1.Rows(Count).Cells(12).Value()
                Data_Count += 1
            End If
        Next
        '出荷指示登録を行い、OUT_PRT、OUT_TBLへの登録を行う。
        '出荷予定データ検索Function
        Result = InsOutShipping_Plan(PLACE_ID, RegistData, OUT_Date, Result, ErrorMessage)

        If result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("出荷指示登録が完了しました。")



    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim IName_ErrorMessage As String = Nothing
        Dim IName_Result As Boolean = True

        Dim I_Name As String = Nothing
        Dim Iid As Integer = 0
        Dim I_Code As String = Nothing
        Dim Item_Data() As Item_List = Nothing

        Dim Item_List() As Item_List = Nothing

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            '入力チェック
            If Trim(TextBox3.Text) = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox3.Focus()
                TextBox3.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された商品コードを元に商品名を取得する。
            '商品名取得Function
            Result = GetItemName(Trim(TextBox3.Text), 1, Item_Data, Result, ErrorMessage)
            If Result = "True" Then

                Label13.Text = Item_Data(0).I_NAME
                TextBox3.BackColor = Color.White
                'ComboBox2.Focus()
                I_Id = Item_Data(0).ID
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox3.Focus()
                TextBox3.BackColor = Color.Salmon
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As OutShipping_Search_List = Nothing
        Dim Data_Total As Integer
        Dim Data_Num_Total As Integer

        Dim S_Status As Integer

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        'CheckBoxクリア
        CheckBox9.Checked = False

        '検索ボタンクリックチェック
        S_shijiSearchFLg = False

        '検索条件のチェック

        '出荷ステータス
        If RadioButton1.Checked Then
            S_Status = 1
        Else
            S_Status = 2
        End If

        '出荷予定データ検索Function
        Result = GetOutShipping_Plan_Search(Trim(TextBox3.Text), PLACE_ID, ComboBox4.Text, ComboBox2.Text, _
                             S_Status, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            Label10.Text = "商品数："
            Label11.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '件数を表示
        Label10.Text = "商品数：" & Data_Total
        '総商品数を表示
        Label11.Text = "総数：" & Data_Num_Total

        'DataGridへ入力したデータを挿入
        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '商品コード
                .Cells(1).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(2).Value = SearchResult(Count).I_NAME
                '出荷希望数
                .Cells(3).Value = SearchResult(Count).NUM
                '出荷指示予定数
                .Cells(4).Value = SearchResult(Count).PLAN_NUM
                '出荷予定数
                .Cells(5).Value = SearchResult(Count).OUT_NUM
                '現在庫数
                If SearchResult(Count).STATUS = "通常出荷" Then
                    .Cells(6).Value = SearchResult(Count).STOCK_NUM
                ElseIf SearchResult(Count).STATUS = "伝票出力のみ" Then
                    .Cells(6).Value = "－"
                End If

                If SearchResult(Count).STATUS = "通常出荷" Then
                    '差し引き在庫数（現在庫 - 出荷予定テーブルの出荷指示予定数 - OUT_TBLの出荷予定数）

                    .Cells(7).Value = SearchResult(Count).STOCK_NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).OUT_NUM
                    If SearchResult(Count).STOCK_NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).OUT_NUM < 0 Then
                        .Cells(7).Style.BackColor = Color.Salmon
                    Else
                        .Cells(7).Style.BackColor = Color.White
                    End If
                ElseIf SearchResult(Count).STATUS = "伝票出力のみ" Then
                    .Cells(7).Value = "－"
                End If
                '出荷倉庫
                .Cells(8).Value = SearchResult(Count).P_NAME
                '区分
                .Cells(9).Value = SearchResult(Count).I_STATUS
                '出荷ステータス
                .Cells(10).Value = SearchResult(Count).S_STATUS
                'ステータス
                .Cells(11).Value = SearchResult(Count).STATUS
                'ID
                .Cells(12).Value = SearchResult(Count).ID
                'P_ID
                .Cells(13).Value = SearchResult(Count).P_ID
                'I_ID
                .Cells(14).Value = SearchResult(Count).I_ID
            End With
            DataGridView1.Rows.Add(SR_List)
        Next

    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount = 14

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
                If DataGridView1(0, i).Value = True Then
                    DataGridView1(0, i).Value = 0
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.White
                        If Count = 7 And DataGridView1.Rows(i).Cells(11).Value() = "通常出荷" Then
                            If DataGridView1.Rows(i).Cells(7).Value() < 0 Then
                                DataGridView1.Item(Count, i).Style.BackColor = Color.Salmon
                            End If
                        End If
                    Next
                ElseIf DataGridView1(0, i).Value = 0 Then
                    DataGridView1(0, i).Value = True
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.PaleGreen
                    Next
                End If
            Next
        Else
            For LoopCount = 0 To CheckData.Length - 1
                If DataGridView1(0, CheckData(LoopCount)).Value = 1 Then
                    If CheckBox9.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.White
                        Next
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox9.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = True
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
        Dim IndexCount As Integer = 14
        '列のインデックスを確認する
        If e.ColumnIndex = 0 Then
            If FormLord = True Then
                If DataGridView1(e.ColumnIndex, e.RowIndex).Value = True Then
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.PaleGreen
                    Next
                Else
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.White
                        If count = 7 And DataGridView1.Rows(e.RowIndex).Cells(11).Value() = "通常出荷" Then
                            If DataGridView1.Rows(e.RowIndex).Cells(7).Value() < 0 Then
                                DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.Salmon
                            End If

                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Dim Csv_Complete_Message As String = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As OutShipping_Search_List = Nothing
        Dim Data_Total As Integer
        Dim Data_Num_Total As Integer

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing

        Dim S_Status As Integer

        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim Item_Name As String = Nothing

        Dim LineData As String = Nothing

        'DataGridViewをクリアする。
        'DataGridView1.Rows.Clear()
        '検索ボタンクリックチェック
        S_shijiSearchFLg = False

        '検索条件のチェック

        'ステータス
        If RadioButton1.Checked Then
            S_Status = 1
        Else
            S_Status = 2
        End If

        '出荷予定データ検索Function
        Result = GetOutShipping_Plan_Search(Trim(TextBox3.Text), PLACE_ID, ComboBox4.Text, ComboBox2.Text, _
                     S_Status, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            'Label10.Text = "商品数："
            'Label11.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "出荷予定関連データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "商品コード,商品名,出荷希望数,出荷指示予定数,出荷予定数,現在庫数,差引在庫数,出荷倉庫,区分,出荷ステータス,ステータス"

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
            'A列に商品コード
            LineData = """" & SearchResult(i).I_CODE & ""","
            'B列に商品名
            Item_Name = SearchResult(i).I_NAME
            LineData &= """" & Item_Name.Replace("""", ChrW(34) & ChrW(34)) & ""","
            'C列に出荷希望数
            LineData &= """" & SearchResult(i).NUM & ""","
            'D列に出荷希望数
            LineData &= """" & SearchResult(i).PLAN_NUM & ""","
            'E列に出荷予定数
            LineData &= """" & SearchResult(i).OUT_NUM & ""","
            'F列に現在庫数
            If SearchResult(i).STATUS = "通常出荷" Then
                LineData &= """" & SearchResult(i).STOCK_NUM & ""","
            ElseIf SearchResult(i).STATUS = "伝票出力のみ" Then
                LineData &= """" & "－" & ""","
            End If

            'G列に差引在庫数
            If SearchResult(i).STATUS = "通常出荷" Then
                LineData &= """" & SearchResult(i).STOCK_NUM - SearchResult(i).PLAN_NUM - SearchResult(i).OUT_NUM & ""","
            ElseIf SearchResult(i).STATUS = "伝票出力のみ" Then
                LineData &= """" & "－" & ""","
            End If
            'H列に出荷倉庫
            LineData &= """" & SearchResult(i).P_NAME & ""","
            'I列に区分
            LineData &= """" & SearchResult(i).I_STATUS & ""","
            'J列に出荷ステータス
            LineData &= """" & SearchResult(i).S_STATUS & ""","
            'K列にステータス
            LineData &= """" & SearchResult(i).STATUS & """"
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox1.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox1.SelectedValue.ToString()
        End If
    End Sub

End Class