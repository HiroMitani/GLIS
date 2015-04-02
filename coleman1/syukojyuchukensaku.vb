Imports System.Windows.Forms
Imports System.Globalization

Public Class syukojyuchukensaku

    Dim PLACE_ID As String
    Dim I_Id As String
    Dim C_Id As String

    Public FormLord As Boolean = False

    'False:検索ボタンを押したらFalse
    '確定や変更など行ったら、再度検索を行わせる為のFlg
    Public SearchFLg As Boolean = False

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Private Sub syukojyuchukensaku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "受注検索"

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

        '左3項目を固定
        DataGridView1.Columns(2).Frozen = True

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
        ComboBox3.DataSource = PlaceTable

        '表示される値はDataTableのNAME列
        ComboBox3.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox3.ValueMember = "ID"
        '初期値をセット(倉庫データの1件目）
        ComboBox3.SelectedIndex = 0
        PLACE_ID = ComboBox3.SelectedValue.ToString()

        ComboBox1.Text = "良品"
        ComboBox2.Text = "通常出荷"

        FormLord = True
    End Sub

    Private Sub syukojyuchukensaku_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Check_Flg As Boolean = False
        Dim Out_Check_Flg As Boolean = False
        Dim Delete_Data() As OutShipping_Search_List = Nothing
        Dim Data_Count As Integer = 0
        Dim DEL_Result As Boolean = True
        Dim DEL_ErrorMessage As String = Nothing
        Dim Out_AlreadyCheck_Flg As Boolean = True

        If SearchFLg = True Then
            MsgBox("受注内容変更、削除を行った後は再度検索を行って下さい。")
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

        Out_Check_Flg = True
        'チェックされた商品の中でピッキング済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷指示済データ") Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("出荷指示済みデータ、伝票出力のみのデータにチェックされています。出荷予定データのみ出荷確定が行えます。")
            Exit Sub
        End If

        Out_AlreadyCheck_Flg = True
        'チェックされた商品の中で出荷指示済みのデータがチェックされていないか確認
        'チェックされていたらエラー
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(7).Value() <> 0) Then
                Out_AlreadyCheck_Flg = False
            End If
        Next
        If Out_AlreadyCheck_Flg = False Then
            MsgBox("出荷指示済みのデータは削除できません。")
            Exit Sub
        End If


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Delete_Data(0 To Data_Count)
                'ID
                Delete_Data(Data_Count).ID = DataGridView1.Rows(Count).Cells(18).Value()
 
                Data_Count += 1
            End If
        Next

        'ダイアログ設定
        Dim result As DialogResult = MessageBox.Show("データ削除してもよろしいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)
        '何が選択されたか調べる()
        If result = DialogResult.No Then
            '「いいえ」が選択された時 
            Exit Sub
        End If

        '出荷予定データ削除 Function
        DEL_Result = Del_OrderData(Delete_Data, DEL_Result, DEL_ErrorMessage)

        If DEL_Result = False Then
            MsgBox(DEL_ErrorMessage)
            Exit Sub
        End If

        MsgBox("削除が完了しました。")

        SearchFLg = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Check_Flg As Boolean = False
        Dim Out_Check_Flg As Boolean = False
        Dim Update_Data() As OutShipping_Search_List = Nothing
        Dim Data_Count As Integer = 0

        If SearchFLg = True Then
            MsgBox("受注内容変更、削除を行った後は再度検索を行って下さい。")
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

        Out_Check_Flg = True
        'チェックされた商品の中でピッキング済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷指示済データ") Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("出荷指示済みデータ、伝票出力のみのデータにチェックされています。出荷予定データのみ出荷確定が行えます。")
            Exit Sub
        End If


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Update_Data(0 To Data_Count)
                '出荷先コード
                Update_Data(Data_Count).C_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '出荷先名
                Update_Data(Data_Count).C_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                '商品コード
                Update_Data(Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(3).Value()
                '商品名
                Update_Data(Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(4).Value()
                '出荷希望数
                Update_Data(Data_Count).NUM = DataGridView1.Rows(Count).Cells(5).Value()
                '出荷予定指示数
                Update_Data(Data_Count).PLAN_NUM = DataGridView1.Rows(Count).Cells(6).Value()
                '出荷指示済数
                Update_Data(Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(7).Value()
                '納品単価
                Update_Data(Data_Count).D_UNIT_PRICE = DataGridView1.Rows(Count).Cells(8).Value()
                'オーダー番号
                Update_Data(Data_Count).CUSTOMER_ORDER_NO = DataGridView1.Rows(Count).Cells(9).Value()
                '出荷倉庫
                Update_Data(Data_Count).P_NAME = DataGridView1.Rows(Count).Cells(10).Value()
                '区分
                Update_Data(Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(11).Value()
                '出荷ステータス
                Update_Data(Data_Count).S_STATUS = DataGridView1.Rows(Count).Cells(12).Value()
                'ステータス
                Update_Data(Data_Count).STATUS = DataGridView1.Rows(Count).Cells(13).Value()
                'コメント１
                Update_Data(Data_Count).COMMENT1 = DataGridView1.Rows(Count).Cells(14).Value()
                'コメント２
                Update_Data(Data_Count).COMMENT2 = DataGridView1.Rows(Count).Cells(15).Value()
                '出荷予定日
                Update_Data(Data_Count).OUT_DATE = DataGridView1.Rows(Count).Cells(16).Value()
                '登録日時
                Update_Data(Data_Count).U_DATE = DataGridView1.Rows(Count).Cells(17).Value()
                'ID
                Update_Data(Data_Count).ID = DataGridView1.Rows(Count).Cells(18).Value()
                'CID
                Update_Data(Data_Count).C_ID = DataGridView1.Rows(Count).Cells(19).Value()
                'PID
                Update_Data(Data_Count).P_ID = DataGridView1.Rows(Count).Cells(20).Value()
                'IID
                Update_Data(Data_Count).I_ID = DataGridView1.Rows(Count).Cells(21).Value()


                Data_Count += 1
            End If
        Next

        'のDataGridViewをクリア
        syukojyuchuhenkou.DataGridView1.Rows.Clear()

        '出庫確定のDataGridViewに表示する。
        For Count = 0 To Update_Data.Length - 1
            Dim Update_Data_list As New DataGridViewRow
            Update_Data_list.CreateCells(syukojyuchuhenkou.DataGridView1)
            With Update_Data_list
                '出荷先コード
                .Cells(0).Value = Update_Data(Count).C_CODE
                '出荷先名
                .Cells(1).Value = Update_Data(Count).C_NAME
                '商品コード
                .Cells(2).Value = Update_Data(Count).I_CODE
                '商品名
                .Cells(3).Value = Update_Data(Count).I_NAME
                '出荷希望数
                .Cells(4).Value = Update_Data(Count).NUM
                '出荷予定指示数
                .Cells(5).Value = Update_Data(Count).PLAN_NUM
                '出荷指示済数
                .Cells(6).Value = Update_Data(Count).FIX_NUM
                '納品単価
                .Cells(7).Value = Update_Data(Count).D_UNIT_PRICE
                'オーダー番号
                .Cells(8).Value = Update_Data(Count).CUSTOMER_ORDER_NO
                '出荷倉庫
                .Cells(9).Value = Update_Data(Count).P_NAME
                '区分
                .Cells(10).Value = Update_Data(Count).I_STATUS
                '出荷ステータス
                .Cells(11).Value = Update_Data(Count).S_STATUS
                'ステータス
                .Cells(12).Value = Update_Data(Count).STATUS
                'コメント１
                .Cells(13).Value = Update_Data(Count).COMMENT1
                'コメント２
                .Cells(14).Value = Update_Data(Count).COMMENT2
                '出荷予定日
                .Cells(15).Value = Update_Data(Count).OUT_DATE
                '登録日時
                .Cells(16).Value = Update_Data(Count).U_DATE
                'ID
                .Cells(17).Value = Update_Data(Count).ID
                'CID
                .Cells(18).Value = Update_Data(Count).C_ID

                'PID
                .Cells(19).Value = Update_Data(Count).P_ID
                'IID
                .Cells(20).Value = Update_Data(Count).I_ID
            End With
            syukojyuchuhenkou.DataGridView1.Rows.Add(Update_Data_list)
        Next

        syukojyuchuhenkou.Show()
        Me.Hide()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox3.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox3.SelectedValue.ToString()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
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
            If Trim(TextBox1.Text) = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox1.Focus()
                TextBox1.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された商品コードを元に商品名を取得する。
            '商品名取得Function
            Result = GetItemName(Trim(TextBox1.Text), 1, Item_Data, Result, ErrorMessage)
            If Result = "True" Then

                Label15.Text = Item_Data(0).I_NAME
                TextBox1.BackColor = Color.White
                I_Id = Item_Data(0).ID
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox1.Focus()
                TextBox1.BackColor = Color.Salmon
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

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim CustomerName As String = Nothing
        Dim ID As Integer = 0
        Dim C_Code As String = Nothing
        Dim ChkCustomerCodeString As String = Nothing

        '出荷先コード欄をクリアする。
        Label6.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '入力チェック
            If Trim(TextBox2.Text) <> "" Then
                ChkCustomerCodeString = Trim(TextBox2.Text)
                'アイテムコードに'が入力されていたらReplaceする。
                ChkCustomerCodeString = ChkCustomerCodeString.Replace("'", "''")

                Me.ProcessTabKey(Not e.Shift)
                e.Handled = True
                '入力された出荷先コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetCustomerName(ChkCustomerCodeString, 1, CustomerName, ID, C_Code, Result, ErrorMessage)
                If Result = "True" Then
                    Label16.Text = CustomerName
                    TextBox2.BackColor = Color.White
                    C_Id = ID
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox2.Focus()
                    TextBox2.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    Label6.Text = ""
                End If
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

        Dim ChkCommentString As String = Nothing

        Dim OUT_Date_From As String = Nothing
        Dim OUT_Date_To As String = Nothing
        Dim REGIST_Date_From As String = Nothing
        Dim REGIST_Date_To As String = Nothing
        Dim Date_Check_Result As DateTime

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        'CheckBoxクリア
        CheckBox1.Checked = False

        '検索ボタンクリックチェック
        SearchFLg = False

        '検索条件のチェック

        '出荷ステータス
        If RadioButton1.Checked Then
            S_Status = 1
        Else
            S_Status = 2
        End If

        '出荷予定日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            OUT_Date_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUT_Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '出荷予定日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            OUT_Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUT_Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '登録日時Fromのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            REGIST_Date_From = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                REGIST_Date_From = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("データ登録日時の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If
        '登録日時Toのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            REGIST_Date_To = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                REGIST_Date_To = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("データ登録日時の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '納品先コードに'が入っていたらReplaceする。
        ChkCommentString = Trim(TextBox4.Text).Replace("'", "''")

        '受注検索データ検索Function
        Result = GetOrder_Search(Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text), _
                                OUT_Date_From, OUT_Date_To, REGIST_Date_From, _
                                REGIST_Date_To, ComboBox1.Text, ComboBox2.Text, PLACE_ID, _
                             S_Status, ChkCommentString, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

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

        If Result = True Then
            'DataGridへ入力したデータを挿入
            For Count = 0 To SearchResult.Length - 1
                Dim SR_List As New DataGridViewRow
                SR_List.CreateCells(DataGridView1)
                With SR_List
                    '出荷先コード
                    .Cells(1).Value = SearchResult(Count).C_CODE
                    '出荷先名
                    .Cells(2).Value = SearchResult(Count).D_NAME
                    '商品コード
                    .Cells(3).Value = SearchResult(Count).I_CODE
                    '商品名
                    .Cells(4).Value = SearchResult(Count).I_NAME
                    '出荷希望数
                    .Cells(5).Value = SearchResult(Count).NUM
                    '出荷予定指示数
                    .Cells(6).Value = SearchResult(Count).PLAN_NUM
                    '出荷指示済数
                    .Cells(7).Value = SearchResult(Count).FIX_NUM
                    '納品単価
                    .Cells(8).Value = SearchResult(Count).D_UNIT_PRICE
                    '納品単価
                    .Cells(9).Value = SearchResult(Count).CUSTOMER_ORDER_NO

                    '出荷倉庫
                    .Cells(10).Value = SearchResult(Count).P_NAME

                    '区分
                    .Cells(11).Value = SearchResult(Count).I_STATUS
                    '出荷ステータス
                    .Cells(12).Value = SearchResult(Count).S_STATUS
                    'ステータス
                    .Cells(13).Value = SearchResult(Count).STATUS
                    'コメント１
                    .Cells(14).Value = SearchResult(Count).COMMENT1
                    'コメント２
                    .Cells(15).Value = SearchResult(Count).COMMENT2
                    '出荷予定日
                    .Cells(16).Value = SearchResult(Count).OUT_DATE
                    '登録日時
                    .Cells(17).Value = SearchResult(Count).U_DATE
                    'ID
                    .Cells(18).Value = SearchResult(Count).ID
                    'CID
                    .Cells(19).Value = SearchResult(Count).C_ID
                    'PID
                    .Cells(20).Value = SearchResult(Count).P_ID
                    'IID
                    .Cells(21).Value = SearchResult(Count).I_ID
                End With
                DataGridView1.Rows.Add(SR_List)
            Next
        Else
            MsgBox(ErrorMessage)
            Exit Sub
        End If
    End Sub


    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount = 21

        Dim CheckDataCount As Integer = 0

        '2012/02/10 菊池様の依頼により、DataGridViewをドラッグで複数行指定し、
        'チェックボタンを押すことにより、指定された行のチェックをつけたりはずしたりする機能を実装

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

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim count As Integer
        Dim IndexCount As Integer = 21
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As OutShipping_Search_List = Nothing
        Dim Data_Total As Integer
        Dim Data_Num_Total As Integer

        Dim S_Status As Integer

        Dim ChkCommentString As String = Nothing

        Dim OUT_Date_From As String = Nothing
        Dim OUT_Date_To As String = Nothing
        Dim REGIST_Date_From As String = Nothing
        Dim REGIST_Date_To As String = Nothing
        Dim Date_Check_Result As DateTime
        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        'CheckBoxクリア
        CheckBox1.Checked = False

        '検索ボタンクリックチェック
        SearchFLg = False

        '検索条件のチェック

        '出荷ステータス
        If RadioButton1.Checked Then
            S_Status = 1
        Else
            S_Status = 2
        End If

        '出荷予定日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            OUT_Date_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUT_Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '出荷予定日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            OUT_Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUT_Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '登録日時Fromのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            REGIST_Date_From = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                REGIST_Date_From = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("データ登録日時の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If
        '登録日時Toのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            REGIST_Date_To = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                REGIST_Date_To = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("データ登録日時の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '納品先コードに'が入っていたらReplaceする。
        ChkCommentString = Trim(TextBox4.Text).Replace("'", "''")

        '受注検索データ検索Function
        Result = GetOrder_Search(Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text), _
                                OUT_Date_From, OUT_Date_To, REGIST_Date_From, _
                                REGIST_Date_To, ComboBox1.Text, ComboBox2.Text, PLACE_ID, _
                             S_Status, ChkCommentString, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア

            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "受注検索データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "出荷先コード,出荷先名,商品コード,商品名,出荷希望数,出荷予定指示数,出荷指示済数,納品単価,オーダー番号,出荷倉庫,区分,出荷ステータス,ステータス,コメント１,コメント２,出荷予定日,登録日時"

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

        Dim Item_Name As String = Nothing

        'データ件数分ループ
        For i = 0 To SearchResult.Length - 1

            '取得するデータをLineataに格納する。
            'A列に出荷先コード
            LineData = """" & SearchResult(i).C_CODE & ""","
            'B列に出荷先名
            LineData &= """" & SearchResult(i).C_NAME & ""","
            'C列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'D列に商品名
            'LineData &= """" & SearchResult(i).I_NAME & ""","
            Item_Name = SearchResult(i).I_NAME
            LineData &= """" & Item_Name.Replace("""", ChrW(34) & ChrW(34)) & ""","
            'E列に出荷希望数
            LineData &= """" & SearchResult(i).NUM & ""","
            'F列に出荷予定指示数
            LineData &= """" & SearchResult(i).PLAN_NUM & ""","
            'G列に出荷指示済数
            LineData &= """" & SearchResult(i).FIX_NUM & ""","
            'H列に納品単価
            LineData &= """" & SearchResult(i).D_UNIT_PRICE & ""","
            'I列にオーダー番号
            LineData &= """" & SearchResult(i).CUSTOMER_ORDER_NO & ""","
            'J列に出荷倉庫
            LineData &= """" & SearchResult(i).P_NAME & ""","
            'K列に区分
            LineData &= """" & SearchResult(i).I_STATUS & ""","
            'L列に出荷ステータス
            LineData &= """" & SearchResult(i).S_STATUS & ""","
            'M列にステータス
            LineData &= """" & SearchResult(i).STATUS & ""","
            'N列にコメント１
            LineData &= """" & SearchResult(i).COMMENT1 & ""","
            'O列にコメント２
            LineData &= """" & SearchResult(i).COMMENT2 & ""","
            'P列に出荷予定日
            LineData &= """" & SearchResult(i).OUT_DATE & ""","
            'Q列に登録日時
            LineData &= """" & SearchResult(i).U_DATE & """"

            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub
End Class