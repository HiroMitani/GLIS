Public Class zset

    Public I_Id As String
    Dim PLACE_ID As String

    Public Structure Set_Item_Check_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim NUM As Integer
        Dim BEFORE_NUM As Integer
        Dim I_STATUS As Integer
        Dim PLACE As Integer
    End Structure

    Dim Form_Load As Boolean = False

    Private Sub zset_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemData() As Item_List = Nothing

        Dim Location As String = Nothing
        Dim LocationErrorMessage As String = Nothing
        Dim LocationResult As Boolean = True
        Dim ItemCode As String = Nothing

        Dim PlaceData() As Place_List = Nothing

        Dim Disp_Title As String = "セット組"

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

        'ウインドウを表示したとき、１行目の数量にフォーカスを移動。
        DataGridView1.Visible = True
        DataGridView1.Select()
        DataGridView1.CurrentCell = DataGridView1(0, 0)

        '商品（パッケージ商品のみ）情報取得
        Dim CustomerTable As New DataTable()
        CustomerTable.Columns.Add("ID", GetType(String))
        CustomerTable.Columns.Add("NAME", GetType(String))
        '商品（M_ITEM)情報の取得
        Result = GetItemList(1, ItemData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)

            Exit Sub
        End If

        For Count = 0 To ItemData.Length - 1
            '新しい行を作成
            Dim row As DataRow = CustomerTable.NewRow()

            '各列に値をセット
            row("ID") = ItemData(Count).ID
            row("NAME") = ItemData(Count).I_NAME

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
        ComboBox1.SelectedIndex = 0
        I_Id = ComboBox1.SelectedValue.ToString()

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

        '初期値の商品IDの基本ロケーションを取得する。
        LocationResult = GetItemLocation(I_Id, Location, ItemCode, LocationResult, LocationErrorMessage)
        If LocationResult = False Then
            MsgBox(LocationErrorMessage)
            Exit Sub
        End If

        TextBox2.Text = Location
        TextBox3.Text = ItemCode

        Form_Load = True
        ComboBox4.Focus()

    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim GetItemResult As Boolean = True
        Dim GetItemErrorMessage As String = Nothing

        Dim Item_Data() As Stock_Item_List = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        If Form_Load = True Then

            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            Dim Column As Integer = DataGridView1.CurrentCell.ColumnIndex

            'カラムが0なら商品コード
            If Column = 0 Then
                '入力された商品コードで商品名とJANコード、基本ロケーションを取得する。
                GetItemResult = GetStockItem(Trim(DataGridView1(Column, Row).Value), PLACE_ID, Item_Data, GetItemResult, GetItemErrorMessage)
                If GetItemResult = False Then
                    MsgBox(GetItemErrorMessage)

                    '該当行の表示項目をクリア
                    DataGridView1(0, Row).Style.BackColor = Color.Salmon
                    Exit Sub
                End If

                '取得できたら、商品名、JANコード、基本ロケーション、商品IDをセットする。
                DataGridView1(1, Row).Value = Item_Data(0).I_NAME
                DataGridView1(2, Row).Value = Item_Data(0).JAN
                DataGridView1(3, Row).Value = Item_Data(0).NUM
                DataGridView1(5, Row).Value = Item_Data(0).LOCATION
                DataGridView1(6, Row).Value = Item_Data(0).PL_NAME
                DataGridView1(7, Row).Value = Item_Data(0).I_STATUS
                DataGridView1(8, Row).Value = Item_Data(0).STOCK_ID
                DataGridView1(9, Row).Value = Item_Data(0).I_ID

                DataGridView1(0, Row).Style.BackColor = Color.White

                'データが入った時点で倉庫を変更不可能にする。
                ComboBox4.Enabled = False

            ElseIf Column = 4 Then
                ' 整数値、未入力NG、0の値NG、マイナス値NG
                If NumChkVal(Trim(DataGridView1(Column, Row).Value), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView1(Column, Row).Style.BackColor = Color.Salmon

                    MsgBox(NumChkErrorMessage)
                Else
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView1(Column, Row).Style.BackColor = Color.White
                End If
                End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim Location As String = Nothing
        Dim LocationErrorMessage As String = Nothing
        Dim LocationResult As Boolean = True
        Dim ItemCode As String = Nothing

        If Form_Load = True Then
            '    '商品の基本ロケーションを取得する。
            LocationResult = GetItemLocation(ComboBox1.SelectedValue.ToString, Location, ItemCode, LocationResult, LocationErrorMessage)
            If LocationResult = False Then
                MsgBox(LocationErrorMessage)
                Exit Sub
            Else
                '選択されていればSelectedValueに入っている
                If ComboBox1.SelectedIndex <> -1 Then
                    'ラベルに表示
                    I_Id = ComboBox1.SelectedValue.ToString()

                    'comboboxで選択したら入庫元コード入力欄をクリアする。
                    TextBox3.BackColor = Color.White
                    TextBox3.Text = ItemCode
                End If
            End If
        End If
        TextBox2.Text = Location
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        '確認ウインドウにセット組み換え内容を表示する。
        Dim SetItemList() As Set_Item_Check_List = Nothing

        '入力チェック

        'DataGridViewにデータが１件も登録されていなかったらエラーを表示。
        If DataGridView1.Rows(0).Cells(0).Value() = "" Then
            MsgBox("セット組する商品データを入力してください。")
            DataGridView1(0, 0).Style.BackColor = Color.Salmon
            DataGridView1.Visible = True
            DataGridView1.Select()
            DataGridView1.CurrentCell = DataGridView1(0, 0)
            Exit Sub
        End If

        'パッケージ商品が選択されていなかったらエラーを表示。
        If ComboBox1.Text = "" Then
            MsgBox("セット商品を選択してください。")
            ComboBox1.Focus()
            Exit Sub
        End If

        'セットの数量が選択されていなかったらエラーを表示。
        If Trim(TextBox1.Text) = "" Then
            MsgBox("セット商品の数量を入力してください。")
            TextBox1.BackColor = Color.Salmon
            TextBox1.Focus()
            Exit Sub
        Else

            If NumChkVal(Trim(TextBox1.Text), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                MsgBox(NumChkErrorMessage)
                TextBox1.BackColor = Color.Salmon
                TextBox1.Focus()
                Exit Sub
            Else
                TextBox1.BackColor = Color.White
            End If
        End If

        'ロケーションが選択されていなかったらエラーを表示。
        If TextBox2.Text = "" Then
            MsgBox("ロケーションを入力してください。")
            TextBox2.BackColor = Color.Salmon
            TextBox2.Focus()
            Exit Sub
        End If

        'DataGridView1の必要数量チェック
        For Count = 0 To DataGridView1.Rows.Count - 2
            '数値の妥当性チェック

            If DataGridView1.Rows(Count).Cells(4).Value() = "" Then
                MsgBox("数量を入力してください。")
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                Exit Sub
            End If

            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                MsgBox(NumChkErrorMessage)
                Exit Sub
            Else
                DataGridView1(4, Count).Style.BackColor = Color.White
            End If
        Next

        'DataGridView1の値を配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 2
            ReDim Preserve SetItemList(0 To Count)
            'STOCK_ID
            SetItemList(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(8).Value()
            '商品ID
            SetItemList(Count).I_ID = DataGridView1.Rows(Count).Cells(9).Value()
            '商品コード
            SetItemList(Count).I_CODE = DataGridView1.Rows(Count).Cells(0).Value()
            '商品名
            SetItemList(Count).I_NAME = DataGridView1.Rows(Count).Cells(1).Value()
            '数量
            SetItemList(Count).BEFORE_NUM = DataGridView1.Rows(Count).Cells(3).Value()
            '組み換え後数量
            SetItemList(Count).NUM = DataGridView1.Rows(Count).Cells(3).Value() - (DataGridView1.Rows(Count).Cells(4).Value() * Trim(TextBox1.Text))
            '不良区分
            If DataGridView1.Rows(Count).Cells(7).Value() = "良品" Then
                SetItemList(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(7).Value() = "不良品" Then
                SetItemList(Count).I_STATUS = 2
            End If
            '倉庫
            SetItemList(Count).PLACE = Place_ID

        Next

        'zset_confのLabel2に組み替え情報を表示
        For i = 0 To SetItemList.Length - 1
            zset_conf.Label2.Text &= "商品名：" & SetItemList(i).I_NAME & "　　　　数量：" & SetItemList(i).BEFORE_NUM & " → " & SetItemList(i).NUM & vbCr
        Next

        zset_conf.Label2.Text &= vbCr & "を以下の商品にセット組します。" & vbCr
        '追加でセット商品情報を表示
        zset_conf.Label2.Text &= vbCr & "商品名：" & ComboBox1.Text & "　　　　数量：" & Trim(TextBox1.Text)

        '倉庫
        zset_conf.Label3.Text = PLACE_ID
        Me.Hide()
        zset_conf.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
        topmenu.Show()
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim IName_ErrorMessage As String = Nothing
        Dim IName_Result As Boolean = True


        Dim I_Name As String = Nothing

        Dim Iid As Integer = 0

        Dim Item_Data() As Item_List = Nothing
        Dim Item_List() As Item_List = Nothing

        Dim ChkItemCodeString As String = Nothing


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


            ChkItemCodeString = Trim(TextBox3.Text)
            '商品コードに'が入力されたいたらReplaceする。
            ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

            '入力された入庫元コードを元に商品名を取得する。
            '商品名取得Function
            Result = GetSetItemName(ChkItemCodeString, 1, Item_Data, Result, ErrorMessage)
            If Result = "True" Then
                'topmenu.Show()
                'Me.Hide()

                'セット商品リストを取得する。
                IName_Result = GetItemList(1, Item_List, IName_Result, IName_ErrorMessage)
                If IName_Result = False Then
                    MsgBox(IName_ErrorMessage)
                    Exit Sub
                End If

                '1件目からループして、入力した入庫元コードが何番目かを調べ
                'comboboxにセットする。
                For i = 0 To Item_List.Length - 1
                    If Item_Data(0).ID = Item_List(i).ID Then
                        ComboBox1.SelectedIndex = i
                    End If
                Next
                TextBox3.BackColor = Color.White
                ComboBox1.Focus()
                I_Id = Item_Data(0).ID
            ElseIf Result = "False" Then
                MsgBox(ErrorMessage)
                TextBox3.Focus()
                TextBox3.BackColor = Color.Salmon
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub
End Class