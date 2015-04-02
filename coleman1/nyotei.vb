Public Class nyotei

    Dim C_Id As String
    Dim PLACE_ID As String

    Private Sub nyotei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Doc_Result As Boolean = True
        Dim CustomerData() As C_List = Nothing
        Dim Count As Integer = 0
        Dim DocNo As String = Nothing
        Dim PlaceData() As Place_List = Nothing

        Dim Disp_Title As String = "入庫予定入力"

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

        '初期表示時、ドキュメント№にフォーカス
        TextBox1.Focus()

        'DataGridViewのクリア
        DataGridView1.Rows.Clear()
        '入庫元ComboBoxのクリア
        ComboBox1.Items.Clear()
        '入力項目のクリア
        'ドキュメント№
        TextBox1.Text = ""
        '商品コード
        TextBox2.Text = ""
        '商品名
        TextBox3.Text = ""
        '数量
        TextBox4.Text = ""

        '数量の欄の右クリックで表示されるコンテキストメニューを表示させない
        TextBox4.ContextMenu = New ContextMenu

        '最終入力済みドキュメント№を取得
        Doc_Result = GetDocNo(DocNo, Result, ErrorMessage)
        If Doc_Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '最終入力済みドキュメント№を表示
        Label10.Text = "最終入力済みドキュメント№：" & DocNo

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
        '初期値をセット(PFJ-アメリカにあわせる）
        ComboBox1.SelectedIndex = 1
        C_Id = ComboBox1.SelectedValue.ToString()

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

        '不良区分に初期値（良品）を設定
        ComboBox2.SelectedIndex = 0

        '入庫種別に初期値（通常入庫）を設定
        ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub nyotei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        Dim Count As Integer

        Dim Category As Integer
        Dim Defect_Type As Integer

        Dim DocNo_Check_Result As Boolean = True
        Dim DocNo_Check_ErrorMessage As String = Nothing

        'ドキュメント№、数値入力チェック用
        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        Dim ChkDocNoString As String = Nothing

        Dim C_Code As String = Nothing
        Dim CId As Integer = 0
        Dim C_Name As String = Nothing

        Dim Doc_Result As Boolean = True
        Dim DocNo As String = Nothing

        '入力、文字数チェック
        'ドキュメント№
        ChkDocNoString = Trim(TextBox1.Text)

        'ドキュメント№に'が入力されたいたらReplaceする。
        ChkDocNoString = ChkDocNoString.Replace("'", "''")

        '入力チェック Start -------
        'ドキュメント№
        If Trim(TextBox1.Text) = "" Then
            MsgBox("ドキュメント№は必須項目です。入力してください。")
            TextBox1.Focus()
            TextBox1.BackColor = Color.Salmon
            Exit Sub
        End If

        '文字数チェック
        'ドキュメント№が17文字以上ならエラーメッセージを表示
        If ChkDocNoString.Length > 16 Then
            MsgBox("ドキュメント№の文字数が長すぎます。16文字までで設定して下さい。")
            Exit Sub
        End If

        'チェックOKなら背景色は白にする。
        TextBox1.BackColor = Color.White

        If DataGridView1(0, 0).Value = "" Then
            MsgBox("商品が登録されていません。")
            TextBox2.Focus()
            Exit Sub
        End If

        'ドキュメント№がすでに登録されているかチェック。
        DocNo_Check_Result = DocNo_Check(ChkDocNoString, DocNo_Check_Result, DocNo_Check_ErrorMessage)
        If DocNo_Check_Result = False Then
            MsgBox(DocNo_Check_ErrorMessage)
            Exit Sub

        End If

        'DataGridViewに登録されたデータ件数を調べ配列に格納
        Dim dt() As Ins_Item_List = Nothing
        ReDim dt(DataGridView1.Rows.Count - 2)

        For Count = 0 To DataGridView1.Rows.Count - 2
            dt(Count).ITEMCODE = DataGridView1(0, Count).Value
            dt(Count).NAME = DataGridView1(1, Count).Value
            dt(Count).NUM = DataGridView1(2, Count).Value
            dt(Count).ID = DataGridView1(3, Count).Value
        Next

        '種別のチェック
        If ComboBox3.Text = "通常入荷" Then
            Category = 1
        Else
            '返品入荷なら2
            Category = 2
        End If

        '種別のチェック
        If ComboBox2.Text = "良品" Then
            Defect_Type = 1
        Else
            '返品入荷なら2
            Defect_Type = 2
        End If
        '入庫予定データ登録
        Result = InsItem(ChkDocNoString, Defect_Type, Category, DateTimePicker1.Value.ToShortDateString(), _
                        C_Id, R_User, PLACE_ID, dt, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("入庫予定の登録が完了しました。")

        '2012/2/1　原さんからの依頼により、登録後は登録内容をクリアし連続入力が行えるよう対応

        '入力項目のクリア

        'ドキュメント№
        TextBox1.Text = ""
        '入庫予定日(今日の日付）
        DateTimePicker1.Text = DateTime.Now.ToString("yyyy/MM/dd")

        '入庫元（プルダウン）
        '初期値をセット
        ComboBox1.SelectedIndex = 1
        C_Id = ComboBox1.SelectedValue.ToString()

        Result = GetCustomerName(C_Id, 2, C_Name, CId, C_Code, Result, ErrorMessage)

        'C_IDを元にC_Codeを取得しセットする。
        TextBox5.Text = C_Code

        '不良区分、入庫種別は初期値にしなくてもOK（2012/02/01 原さんからの依頼）
        '不良区分
        'ComboBox2.Text = "良品"
        '入庫種別
        'ComboBox3.Text = "通常入荷"
        '商品コード
        TextBox2.Text = ""
        '商品名
        TextBox3.Text = ""
        '数量
        TextBox4.Text = ""
        'DataGridView
        DataGridView1.Rows.Clear()

        '最終入力済みドキュメント№を取得
        Doc_Result = GetDocNo(DocNo, Result, ErrorMessage)
        If Doc_Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '最終入力済みドキュメント№を表示
        Label10.Text = "最終入力済みドキュメント№：" & DocNo

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
        topmenu.Show()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True
            DateTimePicker1.Focus()
        End If
    End Sub
    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemName() As Item_List = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True


        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(TextBox4.Text), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                MsgBox(NumChkErrorMessage)
                TextBox4.BackColor = Color.Salmon
                Exit Sub
            End If

            TextBox4.BackColor = Color.White

            '商品コードが入力されているかチェック
            If TextBox2.Text = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox2.BackColor = Color.Salmon
                TextBox2.Focus()
                Exit Sub
            End If

            'DataGrid登録前に再度商品コードと商品名をチェックする。
            '商品コードが入力されてもTabで移動されていると
            '商品名が取得できていないのでチェック
            If TextBox3.Text = "" Then
                '入力された商品コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetItemName(Trim(TextBox2.Text), 1, ItemName, Result, ErrorMessage)
                If Result = "True" Then
                    TextBox2.BackColor = Color.White
                    TextBox3.Text = ItemName(0).I_NAME
                    Label11.Text = ItemName(0).ID

                    TextBox4.Focus()
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox2.Focus()
                    TextBox2.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    TextBox3.Text = ""
                    Exit Sub
                End If
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            If DataGridView1.Rows.Count = 0 Then
                MsgBox("商品が登録されていません。")
                TextBox2.Focus()
                Exit Sub
            End If
            DataGridView1.Visible = True
            DataGridView1.Select()

            'DataGridの行が何行目が選択されているかチェック。
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            '一番下のデータなら、データを追加する。
            'データの存在する行だったらデータを上書きする。

            If DataGridView1.Rows(Row).Cells(0).Value() <> "" Then
                '上書きする
                '商品コード
                DataGridView1(0, Row).Value = Trim(TextBox2.Text)
                '商品名
                DataGridView1(1, Row).Value = TextBox3.Text
                '数量
                DataGridView1(2, Row).Value = Integer.Parse(Trim(TextBox4.Text))
                '商品ID
                DataGridView1(3, Row).Value = Label11.Text

            ElseIf DataGridView1.Rows(Row).Cells(0).Value() = "" Then
                '商品コードが空の行は新規の行と認識し、データを追加する。

                'DataGridへ入力したデータを挿入
                Dim item As New DataGridViewRow
                item.CreateCells(DataGridView1)
                With item
                    .Cells(0).Value = TextBox2.Text
                    .Cells(1).Value = TextBox3.Text
                    .Cells(2).Value = Integer.Parse(Trim(TextBox4.Text))
                    '非表示項目に商品ID
                    .Cells(3).Value = Label11.Text

                End With
                DataGridView1.Rows.Add(item)

                '最新の行を取得
                Dim datacnt As Integer
                datacnt = DataGridView1.Rows.Count - 1
                '挿入した行へフォーカスを移動
                DataGridView1.CurrentCell = DataGridView1(0, datacnt)
            End If

            '商品コードをクリア
            TextBox2.Text = ""
            '商品名をクリア
            TextBox4.Text = ""
            '数量をクリア
            TextBox3.Text = ""
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemName() As Item_List = Nothing

        TextBox3.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            '入力チェック
            If Trim(TextBox2.Text) = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox2.Focus()
                TextBox2.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された商品コードを元に商品名を取得する。
            '商品名称取得Function
            Result = GetItemName(Trim(TextBox2.Text), 1, ItemName, Result, ErrorMessage)
            If Result = "True" Then
                'topmenu.Show()
                'Me.Hide()
                TextBox2.BackColor = Color.White
                TextBox3.Text = ItemName(0).I_NAME
                Label11.Text = ItemName(0).ID
                TextBox4.Focus()
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox2.Focus()
                TextBox2.BackColor = Color.Salmon
                'エラーの場合、商品名もクリア。
                TextBox3.Text = ""
            End If
        End If
    End Sub
    Private Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing
        Dim C_Name As String = Nothing
        Dim Cid As Integer = 0
        Dim C_Code As String = Nothing

        '選択されていればSelectedValueに入っている
        If ComboBox1.SelectedIndex <> -1 Then
            'ラベルに表示
            C_Id = ComboBox1.SelectedValue.ToString()

            '選択されたプルダウンの企業名から取得された企業IDを元に企業コードを取得する。
            '納入元名取得Function
            Result = GetCustomerName(C_Id, 2, C_Name, Cid, C_Code, Result, ErrorMessage)

            TextBox5.BackColor = Color.White
            TextBox5.Text = C_Code

        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim CName_ErrorMessage As String = Nothing
        Dim CName_Result As Boolean = True

        Dim C_Name As String = Nothing
        Dim Cid As Integer = 0
        Dim C_Code As String = Nothing
        Dim Customer_List() As C_List = Nothing

        'TextBox3.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            '入力チェック
            If Trim(TextBox5.Text) = "" Then
                MsgBox("入庫元コードを入力してください。")
                TextBox5.Focus()
                TextBox5.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された入庫元コードを元に商品名を取得する。
            '納入元名取得Function
            Result = GetCustomerName(Trim(TextBox5.Text), 1, C_Name, Cid, C_Code, Result, ErrorMessage)
            If Result = "True" Then
                'topmenu.Show()
                'Me.Hide()

                '企業マスタ全件を取得する。
                Result = GetCustomerNameList(Customer_List, CName_Result, CName_ErrorMessage)
                If CName_Result = False Then
                    MsgBox(CName_ErrorMessage)
                    Exit Sub
                End If

                '1件目からループして、入力した入庫元コードが何番目かを調べ
                'comboboxにセットする。
                For i = 0 To Customer_List.Length - 1
                    If Cid = Customer_List(i).ID Then
                        ComboBox1.SelectedIndex = i
                    End If
                Next
                TextBox5.BackColor = Color.White
                ComboBox2.Focus()
                C_Id = Cid
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox5.Focus()
                TextBox5.BackColor = Color.Salmon
            End If
        End If
    End Sub


    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim Row As Integer = 0
        '選択されている行を取得
        Row = e.RowIndex

        '選択されている行のデータを取得し、DataGridの上のTextboxに表示する。
        '商品コード
        TextBox2.Text = DataGridView1(0, e.RowIndex).Value
        '商品名
        TextBox3.Text = DataGridView1(1, e.RowIndex).Value
        '数量
        TextBox4.Text = DataGridView1(2, e.RowIndex).Value
        '商品ID
        Label11.Text = DataGridView1(3, e.RowIndex).Value
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged

        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class
