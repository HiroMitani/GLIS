Public Class syukoyotei

    Dim C_Id As String
    Dim C_DISCOUNT_RATE As String
    Dim PLACE_ID As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub syukoyotei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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


    Private Sub syukoyotei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Doc_Result As Boolean = True
        Dim CustomerData() As C_List = Nothing
        Dim PlaceData() As Place_List = Nothing

        Dim Disp_Title As String = "受注入力"

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

        'DataGridViewのクリア
        DataGridView1.Rows.Clear()
        '出荷先名ComboBoxのクリア
        ComboBox1.Items.Clear()
        '倉庫ComboBoxのクリア
        ComboBox2.Items.Clear()
        '入力項目のクリア
        '出荷先ＩＤクリア
        TextBox1.Text = ""
        '客先発注Noクリア
        TextBox2.Text = ""
        'コメント１クリア
        TextBox3.Text = ""
        'コメント２クリア
        TextBox4.Text = ""

        '出荷先情報取得
        Dim CustomerTable As New DataTable()
        CustomerTable.Columns.Add("ID", GetType(String))
        CustomerTable.Columns.Add("C_DISCOUNT_RATE", GetType(String))
        CustomerTable.Columns.Add("NAME", GetType(String))
        '出荷先（M_CUSTOMER)情報の取得
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
            row("C_DISCOUNT_RATE") = CustomerData(Count).DISCOUNT_RATE
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
        '初期値をセット(1件目にあわせる）
        ComboBox1.SelectedIndex = 0
        C_Id = ComboBox1.SelectedValue.ToString()
        C_DISCOUNT_RATE = CustomerData(0).DISCOUNT_RATE
        Label14.Text = CustomerData(0).DISCOUNT_RATE


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
        ComboBox2.DataSource = PlaceTable

        '表示される値はDataTableのNAME列
        ComboBox2.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox2.ValueMember = "ID"
        '初期値をセット(倉庫データの1件目）
        ComboBox2.SelectedIndex = 0
        PLACE_ID = ComboBox2.SelectedValue.ToString()

        ComboBox3.Text = "良品"
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim CName_ErrorMessage As String = Nothing
        Dim CName_Result As Boolean = True

        Dim C_Name As String = Nothing
        Dim Cid As Integer = 0
        Dim C_Code As String = Nothing
        Dim Customer_List() As C_List = Nothing
        Dim Discount_rate As Decimal

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            '入力チェック
            If Trim(TextBox1.Text) = "" Then
                MsgBox("出荷先コードを入力してください。")
                TextBox1.Focus()
                TextBox1.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された入庫元コードを元に商品名を取得する。
            '納入元名取得Function
            Result = GetCustomerName(Trim(TextBox1.Text), 1, C_Name, Cid, C_Code, Discount_rate, Result, ErrorMessage)
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
                TextBox1.BackColor = Color.White
                ComboBox2.Focus()
                C_Id = Cid
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox1.Focus()
                TextBox1.BackColor = Color.Salmon
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemData() As Item_List = Nothing

        TextBox6.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            '入力チェック
            If Trim(TextBox5.Text) = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox5.Focus()
                TextBox5.BackColor = Color.Salmon
                Exit Sub
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            '入力された商品コードを元に商品名を取得する。
            '商品名称取得Function
            Result = GetItemName(Trim(TextBox5.Text), 1, ItemData, Result, ErrorMessage)
            If Result = "True" Then
                'topmenu.Show()
                'Me.Hide()
                TextBox5.BackColor = Color.White
                TextBox6.Text = ItemData(0).I_NAME
                '納品単価
                ' TextBox7.Text = ItemData(0).PURCHASE_PRICE
                '価格*M_CUSTOMERの掛け率を算出（結果を四捨五入）

                TextBox7.Text = Math.Round(ItemData(0).PRICE * Label14.Text)

                '免責金額
                TextBox10.Text = ItemData(0).IMMUNITY_PRICE
                '修理金額
                TextBox9.Text = ItemData(0).REPAIR_PRICE
                '納品金額
                Label16.Text = ItemData(0).ID

                TextBox7.Focus()
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox5.Focus()
                TextBox5.BackColor = Color.Salmon
                'エラーの場合、商品名もクリア。
                TextBox6.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
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
            If NumChkVal(Trim(TextBox8.Text), "INTEGER", False, False, True, NumChkResult, NumChkErrorMessage) = False Then
                MsgBox(NumChkErrorMessage)
                TextBox8.BackColor = Color.Salmon
                Exit Sub
            End If

            TextBox8.BackColor = Color.White

            '商品コードが入力されているかチェック
            If TextBox5.Text = "" Then
                MsgBox("商品コードを入力してください。")
                TextBox5.BackColor = Color.Salmon
                TextBox5.Focus()
                Exit Sub
            End If

            'DataGrid登録前に再度商品コードと商品名をチェックする。
            '商品コードが入力されてもTabで移動されていると
            '商品名が取得できていないのでチェック
            If TextBox6.Text = "" Then
                '入力された商品コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetItemName(Trim(TextBox5.Text), 1, ItemName, Result, ErrorMessage)
                If Result = "True" Then
                    TextBox5.BackColor = Color.White
                    TextBox6.Text = ItemName(0).I_NAME

                    '納品単価
                    TextBox7.Text = Math.Round(ItemName(0).PRICE * Label14.Text)
                    '免責金額
                    TextBox10.Text = ItemName(0).IMMUNITY_PRICE
                    '修理金額
                    TextBox9.Text = ItemName(0).REPAIR_PRICE
                    'ID
                    Label16.Text = ItemName(0).ID

                    TextBox7.Focus()
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox5.Focus()
                    TextBox5.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    TextBox6.Text = ""
                    Exit Sub
                End If
            End If

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            If DataGridView1.Rows.Count = 0 Then
                MsgBox("商品が登録されていません。")
                TextBox5.Focus()
                Exit Sub
            End If
            DataGridView1.Visible = True
            DataGridView1.Select()

            'DataGridの行が何行目が選択されているかチェック。
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            '一番下のデータなら、データを追加する。
            'データの存在する行だったらデータを上書きする。

            If DataGridView1.Rows(Row).Cells(0).Value() <> "" Then
                '上書きする前に登録しようとしている商品コードがすでにDataGridにあればエラー。
                For Count = 0 To DataGridView1.Rows.Count - 1
                    '商品コード
                    If Trim(TextBox5.Text).ToUpper() = DataGridView1(0, Count).Value And DataGridView1.CurrentCellAddress.Y <> Count Then
                        'MsgBox(DataGridView1.CurrentCell.Value)

                        MsgBox("登録しようとしている商品コードはすでに登録されています。")
                        TextBox5.Focus()
                        Exit Sub
                    End If
                Next

                '上書きする
                '商品コード
                DataGridView1(0, Row).Value = Trim(TextBox5.Text).ToUpper()
                '商品名
                DataGridView1(1, Row).Value = TextBox6.Text
                '納品単価
                DataGridView1(2, Row).Value = TextBox7.Text

                '数量
                DataGridView1(3, Row).Value = Integer.Parse(Trim(TextBox8.Text))
                '納品金額
                DataGridView1(4, Row).Value = TextBox7.Text * Integer.Parse(Trim(TextBox8.Text))
                '商品ID
                DataGridView1(5, Row).Value = Label16.Text

            ElseIf DataGridView1.Rows(Row).Cells(0).Value() = "" Then
                '商品コードが空の行は新規の行と認識し、データを追加する。

                '登録しようとしている商品コードがすでにDataGridにあればエラー。
                For Count = 0 To DataGridView1.Rows.Count - 1
                    '商品コード
                    If Trim(TextBox5.Text).ToUpper() = DataGridView1(0, Count).Value Then
                        MsgBox("登録しようとしている商品コードはすでに登録されています。")
                        TextBox5.Focus()
                        Exit Sub
                    End If
                Next

                'DataGridへ入力したデータを挿入
                Dim item As New DataGridViewRow
                item.CreateCells(DataGridView1)
                With item
                    .Cells(0).Value = TextBox5.Text.ToUpper()
                    .Cells(1).Value = TextBox6.Text
                    '納品単価
                    .Cells(2).Value = Integer.Parse(Trim(TextBox7.Text))
                    '数量
                    .Cells(3).Value = Integer.Parse(Trim(TextBox8.Text))
                    '納品金額
                    .Cells(4).Value = Integer.Parse(Trim(TextBox7.Text)) * Integer.Parse(Trim(TextBox8.Text))
                    '商品ID
                    .Cells(5).Value = Label16.Text

                End With
                '先頭行に追加
                DataGridView1.Rows.Insert(0, item)

                '最新の行を取得
                Dim datacnt As Integer
                datacnt = DataGridView1.Rows.Count - 1
                '挿入した行へフォーカスを移動
                DataGridView1.CurrentCell = DataGridView1(0, datacnt)
            End If

            '商品コードをクリア
            TextBox5.Text = ""
            '商品名をクリア
            TextBox6.Text = ""
            '納品単価をクリア
            TextBox7.Text = ""
            '数量をクリア
            TextBox8.Text = ""
            '修理金額をクリア
            TextBox9.Text = ""
            '免責金額をクリア
            TextBox10.Text = ""

            TextBox5.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Tmp_Comment1 As String = Nothing
        Dim Tmp_Comment2 As String = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim NUM_Check As Boolean = True

        Dim Order_NO_Result As Boolean = True
        Dim Order_NO_ErrorMessage As String = Nothing

        '入力チェック Start
        '出荷先名の指定がなければエラー
        If Trim(ComboBox1.Text) = "" Then
            MsgBox("出荷先名は必須項目です。入力してください。")
            ComboBox1.Focus()
            Exit Sub
        End If

        'オーダー番号が入力されていなければエラー
        If Trim(TextBox2.Text) = "" Then
            MsgBox("オーダー番号は必須項目です。入力してください。")
            TextBox2.Focus()
            Exit Sub
        End If

        'オーダー番号は、24文字以上入力されていたらエラー
        If TextBox2.Text.Length > 24 Then
            MsgBox("オーダー番号が長すぎます。24文字以下にしてください。")
            TextBox2.Focus()
            Exit Sub
        End If

        '出荷予定日が入力されていなかったらエラー。
        If Trim(DateTimePicker1.Text) = "" Then
            MsgBox("出荷予定日は必須項目です。入力してください。")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        '出荷倉庫
        If Trim(ComboBox2.Text) = "" Then
            MsgBox("出荷倉庫は必須項目です。入力してください。")
            ComboBox2.Focus()
            Exit Sub
        End If

        '区分
        If Trim(ComboBox3.Text) = "" Then
            MsgBox("区分は必須項目です。入力してください。")
            ComboBox3.Focus()
            Exit Sub
        End If

        'コメント１は未入力可だが、400文字以上入力されていたらエラー
        If TextBox3.Text.Length > 401 Then
            MsgBox("コメント１が長すぎます。400文字以下にしてください。")
            TextBox3.Focus()
            Exit Sub
        Else
            '　アポストロフィ対応とバックスラッシュ対応
            Tmp_Comment1 = Trim(TextBox3.Text)
            Tmp_Comment1 = Tmp_Comment1.Replace("'", "''")
            Tmp_Comment1 = Tmp_Comment1.Replace("\", "\\")
        End If

        'コメント２は未入力可だが、400文字以上入力されていたらエラー
        If TextBox4.Text.Length > 401 Then
            MsgBox("コメント２が長すぎます。400文字以下にしてください。")
            TextBox4.Focus()
            Exit Sub
        Else
            '　アポストロフィ対応とバックスラッシュ対応
            Tmp_Comment2 = Trim(TextBox4.Text)
            Tmp_Comment2 = Tmp_Comment2.Replace("'", "''")
            Tmp_Comment2 = Tmp_Comment2.Replace("\", "\\")
        End If

        'DataGridが登録されているかチェック。
        If DataGridView1(0, 0).Value = "" Then
            MsgBox("商品情報が登録されていません。")
            TextBox5.Focus()
            Exit Sub
        End If

        '登録されていれば、数量がすべてプラスかチェック。
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1(3, Count).Value < 0 Then
                NUM_Check = False
            End If
        Next

        If NUM_Check = False Then
            MsgBox("出荷希望数は数量が全て1以上でないと登録されません。")
            Exit Sub
        End If

        '登録データを格納
        'DataGridViewに登録されたデータ件数を調べ配列に格納
        Dim dt() As S_Yotei_List = Nothing
        ReDim dt(DataGridView1.Rows.Count - 2)

        For Count = 0 To DataGridView1.Rows.Count - 2
            '商品コード
            dt(Count).I_CODE = DataGridView1(0, Count).Value
            '納品単価
            dt(Count).UNIT_PRICE = DataGridView1(2, Count).Value
            '受注数
            dt(Count).NUM = DataGridView1(3, Count).Value
            '納品金額
            dt(Count).PURCHASE_PRICE = DataGridView1(4, Count).Value
            '商品ID
            dt(Count).I_ID = DataGridView1(5, Count).Value
        Next

        'オーダー番号がすでに登録されているかチェックする。
        Order_NO_Result = OrderNo_Check(Trim(TextBox2.Text), Order_NO_Result, Order_NO_ErrorMessage)
        If Order_NO_Result = False Then
            MsgBox(Order_NO_ErrorMessage)
            Exit Sub
        End If

        '登録処理
        '出荷予定予定データ登録
        Result = Ins_S_Yotei(C_Id, Trim(TextBox2.Text), Trim(DateTimePicker1.Value.ToShortDateString()), PLACE_ID, _
                        ComboBox3.Text, Tmp_Comment1, Tmp_Comment2, 1, dt, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("出荷予定の登録が完了しました。")

        '入力内容をクリアする

        '出荷先ＩＤ，出荷先名を1（初期値）に戻す。
        TextBox1.Text = ""
        ComboBox1.SelectedIndex = 1
        C_Id = ComboBox1.SelectedValue.ToString()
        '客先注文No
        TextBox2.Text = ""
        '出荷予定日
        DateTimePicker1.Text = DateTime.Today
        '出荷倉庫
        ComboBox2.SelectedIndex = 0
        PLACE_ID = ComboBox2.SelectedValue.ToString()
        '区分
        ComboBox3.Text = "良品"
        'コメント1
        TextBox3.Text = ""
        'コメント2
        TextBox4.Text = ""

        '商品コード
        TextBox5.Text = ""
        '商品名
        TextBox6.Text = ""
        '納品単価
        TextBox7.Text = ""
        '出荷希望数
        TextBox8.Text = ""
        '免責金額
        TextBox9.Text = ""
        '修理金額
        TextBox10.Text = ""
        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim Tmp_Comment1 As String = Nothing
        Dim Tmp_Comment2 As String = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim NUM_Check As Boolean = True

        Dim Order_NO_Result As Boolean = True
        Dim Order_NO_ErrorMessage As String = Nothing

        '入力チェック Start
        '出荷先名の指定がなければエラー
        If Trim(ComboBox1.Text) = "" Then
            MsgBox("出荷先名は必須項目です。入力してください。")
            ComboBox1.Focus()
            Exit Sub
        End If

        'オーダー番号が入力されていなければエラー
        If Trim(TextBox2.Text) = "" Then
            MsgBox("オーダー番号は必須項目です。入力してください。")
            TextBox2.Focus()
            Exit Sub
        End If

        'オーダー番号は、24文字以上入力されていたらエラー
        If TextBox2.Text.Length > 24 Then
            MsgBox("オーダー番号が長すぎます。24文字以下にしてください。")
            TextBox2.Focus()
            Exit Sub
        End If

        '出荷予定日が入力されていなかったらエラー。
        If Trim(DateTimePicker1.Text) = "" Then
            MsgBox("出荷予定日は必須項目です。入力してください。")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        '出荷倉庫
        If Trim(ComboBox2.Text) = "" Then
            MsgBox("出荷倉庫は必須項目です。入力してください。")
            ComboBox2.Focus()
            Exit Sub
        End If

        '区分
        If Trim(ComboBox3.Text) = "" Then
            MsgBox("区分は必須項目です。入力してください。")
            ComboBox3.Focus()
            Exit Sub
        End If

        'コメント１は未入力可だが、400文字以上入力されていたらエラー
        If TextBox3.Text.Length > 401 Then
            MsgBox("コメント１が長すぎます。400文字以下にしてください。")
            TextBox3.Focus()
            Exit Sub
        Else
            '　アポストロフィ対応とバックスラッシュ対応
            Tmp_Comment1 = Trim(TextBox3.Text)
            Tmp_Comment1 = Tmp_Comment1.Replace("'", "''")
            Tmp_Comment1 = Tmp_Comment1.Replace("\", "\\")
        End If

        'コメント２は未入力可だが、400文字以上入力されていたらエラー
        If TextBox4.Text.Length > 401 Then
            MsgBox("コメント２が長すぎます。400文字以下にしてください。")
            TextBox4.Focus()
            Exit Sub
        Else
            '　アポストロフィ対応とバックスラッシュ対応
            Tmp_Comment2 = Trim(TextBox4.Text)
            Tmp_Comment2 = Tmp_Comment2.Replace("'", "''")
            Tmp_Comment2 = Tmp_Comment2.Replace("\", "\\")
        End If

        'DataGridが登録されているかチェック。
        If DataGridView1(0, 0).Value = "" Then
            MsgBox("商品情報が登録されていません。")
            TextBox5.Focus()
            Exit Sub
        End If

        '登録データを格納
        'DataGridViewに登録されたデータ件数を調べ配列に格納
        Dim dt() As S_Yotei_List = Nothing
        ReDim dt(DataGridView1.Rows.Count - 2)

        For Count = 0 To DataGridView1.Rows.Count - 2
            '商品コード
            dt(Count).I_CODE = DataGridView1(0, Count).Value
            '納品単価
            dt(Count).UNIT_PRICE = DataGridView1(2, Count).Value
            '受注数
            dt(Count).NUM = DataGridView1(3, Count).Value
            '納品金額
            dt(Count).PURCHASE_PRICE = DataGridView1(4, Count).Value
            '商品ID
            dt(Count).I_ID = DataGridView1(5, Count).Value
        Next

        'オーダー番号がすでに登録されているかチェックする。
        Order_NO_Result = OrderNo_Check(Trim(TextBox2.Text), Order_NO_Result, Order_NO_ErrorMessage)
        If Order_NO_Result = False Then
            MsgBox(Order_NO_ErrorMessage)
            Exit Sub
        End If

        '登録処理
        '出荷予定予定データ登録
        Result = Ins_S_Yotei(C_Id, Trim(TextBox2.Text), Trim(DateTimePicker1.Value.ToShortDateString()), PLACE_ID, _
                        ComboBox3.Text, Tmp_Comment1, Tmp_Comment2, 2, dt, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("伝票のみの出荷予定の登録が完了しました。")

        '入力内容をクリアする

        '出荷先ＩＤ，出荷先名を1（初期値）に戻す。
        TextBox1.Text = ""
        ComboBox1.SelectedIndex = 1
        C_Id = ComboBox1.SelectedValue.ToString()
        '客先注文No
        TextBox2.Text = ""
        '出荷予定日
        DateTimePicker1.Text = DateTime.Today
        '出荷倉庫
        ComboBox2.SelectedIndex = 0
        PLACE_ID = ComboBox2.SelectedValue.ToString()
        '区分
        ComboBox3.Text = "良品"
        'コメント1
        TextBox3.Text = ""
        'コメント2
        TextBox4.Text = ""

        '商品コード
        TextBox5.Text = ""
        '商品名
        TextBox6.Text = ""
        '納品単価
        TextBox7.Text = ""
        '出荷希望数
        TextBox8.Text = ""
        '免責金額
        TextBox9.Text = ""
        '修理金額
        TextBox10.Text = ""
        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox2.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox2.SelectedValue.ToString()
        End If
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
            Result = GetCustomerName(C_Id, 2, C_Name, Cid, C_Code, C_DISCOUNT_RATE, Result, ErrorMessage)

            TextBox1.BackColor = Color.White
            TextBox1.Text = C_Code
            Label14.Text = C_DISCOUNT_RATE

            'DataGridや入力欄をクリアする。
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""

            'DataGridViewのクリア
            DataGridView1.Rows.Clear()
        End If
    End Sub

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then

            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub
End Class