Public Class zhenpin
    Dim C_Id As String

    Dim Form_Load As Boolean = False

    Private Sub zhenpin_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '登録ボタンが押されたら

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Return_Check() As Tanaoroshi_List = Nothing
        Dim Return_Data_Count As Integer = 0

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        '出荷指示ファイル名が既に取り込まれているかチェックする為の変数
        Dim FNameErrorMessage As String = Nothing
        Dim FNameResult As Boolean = True

        Dim ChkFileNameString As String = Nothing
        Dim ChkSheetNoString As String = Nothing
        Dim ChkOrderNoString As String = Nothing
        Dim ChkCommentString As String = Nothing

        Dim DataGridErrorMessage As String = Nothing

        '入力チェック

        '出荷指示ファイル名
        If Trim(TextBox4.Text) = "" Then
            MsgBox("出荷指示ファイル名を入力してください。")
            TextBox4.BackColor = Color.Salmon
            TextBox4.Focus()
            Exit Sub
        End If

        '伝票番号
        If Trim(TextBox1.Text) = "" Then
            MsgBox("伝票番号を入力してください。")
            TextBox1.BackColor = Color.Salmon
            TextBox1.Focus()
            Exit Sub
        End If
        TextBox1.BackColor = Color.White
        'オーダー番号
        If Trim(TextBox3.Text) = "" Then
            MsgBox("オーダー番号を入力してください。")
            TextBox3.BackColor = Color.Salmon
            TextBox3.Focus()
            Exit Sub
        End If
        TextBox3.BackColor = Color.White

        '返品先
        If Trim(ComboBox1.Text) = "" Then
            MsgBox("返品先を選択してください。")
            ComboBox1.BackColor = Color.Salmon
            ComboBox1.Focus()
            Exit Sub
        End If
        ComboBox1.BackColor = Color.White

        ChkFileNameString = Trim(TextBox4.Text)
        ChkSheetNoString = Trim(TextBox1.Text)
        ChkOrderNoString = Trim(TextBox3.Text)
        ChkCommentString = Trim(TextBox2.Text)


        '出荷指示ファイル名に'が入力されていたらReplaceする。
        ChkFileNameString = ChkFileNameString.Replace("'", "''")

        '伝票番号に'が入力されていたらReplaceする。
        ChkSheetNoString = ChkSheetNoString.Replace("'", "''")
        'オーダー番号に'が入力されていたらReplaceする。
        ChkOrderNoString = ChkOrderNoString.Replace("'", "''")
        'メモに'が入力されていたらReplaceする。
        ChkCommentString = ChkCommentString.Replace("'", "''")

        '返品ファイル名がすでに登録されていないか確認する。
        '取り込もうとしている出荷指示ファイル名が、すでにDBに登録されているかチェックを行う。
        FNameResult = FileName_Check(ChkFileNameString, FNameResult, FNameErrorMessage)

        If FNameResult = False Then
            MsgBox("ファイル名""" & TextBox4.Text & """は" & FNameErrorMessage)
            TextBox4.Focus()
            Exit Sub
        End If

        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Return_Check(0 To Count)
            '商品ID
            Return_Check(Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
            '商品コード
            Return_Check(Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
            '返品数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の返品数量は" & NumChkErrorMessage & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(4, Count).Style.BackColor = Color.White
                '棚卸後数量
                Return_Check(Count).NEW_NUM = DataGridView1.Rows(Count).Cells(4).Value()
                '現数量
                Return_Check(Count).NUM = DataGridView1.Rows(Count).Cells(3).Value()
            End If
            '倉庫
            Return_Check(Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
            'STOCK_ID
            Return_Check(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
            'P_ID
            Return_Check(Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
            '不良区分
            If DataGridView1.Rows(Count).Cells(8).Value() = "良品" Then
                Return_Check(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(8).Value() = "不良品" Then
                Return_Check(Count).I_STATUS = 2
            End If
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("返品出荷を登録してもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If



        '返品出荷登録Function
        Result = Ins_Return_Out(Return_Check, ChkFileNameString, ChkSheetNoString, ChkOrderNoString, ComboBox1.SelectedValue.ToString, DateTimePicker1.Value.ToShortDateString(), _
                                ChkCommentString, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("返品出荷が完了しました。")

        zkensaku.zSearchFLg = True


        Me.Dispose()
        zkensaku.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'キャンセルボタンが押されたら
        Me.Dispose()
        zkensaku.Show()

    End Sub

    Private Sub zhenpin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "返品出荷確定"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label5.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        'ウインドウを表示したとき、１行目の数量にフォーカスを移動。
        DataGridView1.Visible = True
        DataGridView1.Select()
        DataGridView1.CurrentCell = DataGridView1(4, 0)

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim CustomerData() As C_List = Nothing

        '返品出荷指示ファイル名用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '入庫元情報取得
        Dim CustomerTable As New DataTable()
        CustomerTable.Columns.Add("ID", GetType(String))
        CustomerTable.Columns.Add("NAME", GetType(String))
        '入荷元（M_CUSTOMER)情報の取得
        Result = GetCustomerNameList(CustomerData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        For Count = 0 To CustomerData.Length - 1
            '新しい行を作成
            Dim row As DataRow = CustomerTable.NewRow()

            '各列に値をセット
            row("ID") = CustomerData(Count).ID
            row("NAME") = CustomerData(Count).NAME

            'DataTableに行を追加
            CustomerTable.Rows.Add(row)
        Next

        CustomerTable.AcceptChanges()
        'コンボボックスのDataSourceにDataTableを割り当てる
        ComboBox1.DataSource = CustomerTable

        '表示される値はDataTableのNAME列
        ComboBox1.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox1.ValueMember = "ID"
        '初期値をセット
        ComboBox1.SelectedIndex = 1
        C_Id = ComboBox1.SelectedValue.ToString()

        '出荷指示ファイル名を自動的に入れる。
        'HAIKI+日時
        TextBox4.Text = "HAIKI" & dtNow.ToString("yyyyMMddHHmmss")

        Form_Load = True

    End Sub
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If Form_Load = True Then
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            '数量から棚卸数量をひいて、差分に結果を表示する。
            'もし在庫数量欄にマイナスを入れたらメッセージを表示する。

            Dim NumChkErrorMessage As String = Nothing
            Dim NumChkResult As Boolean = True

            If NumChkVal(Trim(DataGridView1.Rows(Row).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(4, Row).Style.BackColor = Color.Salmon
                MsgBox(NumChkErrorMessage)
                Exit Sub
            Else
                DataGridView1(4, Row).Style.BackColor = Color.White
                DataGridView1(5, Row).Value = DataGridView1(3, Row).Value - DataGridView1(4, Row).Value
            End If

        End If
    End Sub

End Class