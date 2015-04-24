Public Class mitem_modify

    'プロダクトラインID
    Dim PL_Id As Integer


    Private Sub mitem_modify_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub mitem_modify_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim Disp_Title As String = "商品マスタ登録修正"

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
        ComboBox3.DataSource = PLTable

        '表示される値はDataTableのNAME列
        ComboBox3.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox3.ValueMember = "ID"

        '初期値をセット
        ComboBox3.SelectedIndex = -1
        'PL_Id = ComboBox1.SelectedValue.ToString()
        PL_Id = ComboBox3.SelectedValue

        'カーソルを検索条件の商品コードへ。
        TextBox1.Focus()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Itemcode As String = Nothing

        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim Search_Data() As M_Item_List = Nothing

        '入力エリアをクリアする。

        '商品コード
        TextBox3.Text = ""
        '商品名
        TextBox4.Text = ""
        'JANコード
        TextBox6.Text = ""
        '価格
        TextBox5.Text = ""
        '仕入価格
        TextBox10.Text = ""
        '免責価格
        TextBox11.Text = ""
        '修理価格
        TextBox12.Text = ""

        'ベンダーコード
        TextBox7.Text = ""
        'ロケーション
        TextBox8.Text = ""
        '入数
        TextBox2.Text = ""
        'マスターカートンサイズ
        TextBox9.Text = ""

        'プロダクトライン
        ComboBox3.Text = ""

        '商品区分
        ComboBox4.Text = ""

        If TextBox1.Text = "" Then
            MsgBox("商品コードを入力してください。")
            'カーソルを検索条件の商品コードへ。
            TextBox1.Focus()
            Exit Sub
        End If

        '商品コード
        Itemcode = Trim(TextBox1.Text).Replace("'", "''")
        Itemcode = Trim(Itemcode).Replace("\", "\\")

        Result = Get_Mitem_Modify(Itemcode, Search_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        If Search_Data(0).I_NAME = "" Then
            MsgBox("データがみつかりません")
        End If

        '入力欄に値をセットする。
        '商品コード
        TextBox3.Text = Search_Data(0).I_CODE
        '商品名
        TextBox4.Text = Search_Data(0).I_NAME
        'JANコード
        TextBox6.Text = Search_Data(0).JAN
        '価格
        TextBox5.Text = Search_Data(0).PRICE
        '仕入金額
        TextBox10.Text = Search_Data(0).PURCHASE_PRICE
        '免責金額
        TextBox11.Text = Search_Data(0).IMMUNITY_PRICE
        '修理金額
        TextBox12.Text = Search_Data(0).REPAIR_PRICE

        'ベンダーコード
        TextBox7.Text = Search_Data(0).C_CODE
        'ベンダー名
        Label3.Text = Search_Data(0).C_NAME
        'ロケーション
        TextBox8.Text = Search_Data(0).LOCATION
        '入数
        TextBox2.Text = Search_Data(0).IN_BOX_NUM
        'マスターカートンサイズ
        TextBox9.Text = Search_Data(0).MASTER_CARTON_SIZE

        'プロダクトライン
        ComboBox3.SelectedValue = Search_Data(0).PL_ID

        '商品ID
        Label4.Text = Search_Data(0).ID

        '商品区分

        If Search_Data(0).PACKAGE_FLG = 0 Then
            ComboBox4.Text = "通常商品"
        Else
            ComboBox4.Text = "セット商品"
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Update_Data() As M_Item_List = Nothing
        Dim Update_Data_Count As Integer = 0

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim ChkStringAfter As String = Nothing
        Dim ChkDateAfter As String = Nothing
        Dim StringChkErrorMessage As String = Nothing
        Dim StringChkResult As Boolean = True

        Dim DataGridErrorMessage As String = Nothing
        Dim Error_Flg As Boolean = True
        Dim Count As Integer = 0

        ReDim Update_Data(0 To 0)

        '商品IDが入っていなかったら、検索ボタンを押していないことになるのでメッセージ表示
        If Label4.Text = "" Then
            MsgBox("検索をしてから修正ボタンを押してください。")
            Exit Sub
        End If

        '入力チェック
        '商品コードの入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox3.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox3.BackColor = Color.Salmon
            DataGridErrorMessage &= "商品コードが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox3.Text.Length) > 24 Then
                TextBox3.BackColor = Color.Salmon
                DataGridErrorMessage &= "商品コードが長すぎます。" & vbCr
                Error_Flg = False
            Else
                '商品コード格納
                Update_Data(0).I_CODE = Trim(TextBox3.Text)
            End If
        End If

        '商品コードがすでにマスタに存在するかチェックする。


        Result = GetItemDuplicationCheck(Update_Data(0).I_CODE, Count, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        If Count > 1 Then
            TextBox3.BackColor = Color.Salmon
            DataGridErrorMessage &= "商品コードがすでに登録されています。" & vbCr
            Error_Flg = False

        End If

        '商品名
        '商品名入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox4.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox4.BackColor = Color.Salmon
            DataGridErrorMessage &= "商品名が正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox4.Text.Length) > 100 Then
                TextBox4.BackColor = Color.Salmon
                DataGridErrorMessage &= "商品名が長すぎます。" & vbCr
                Error_Flg = False
            Else
                '商品名格納
                Update_Data(0).I_NAME = Trim(TextBox4.Text)
            End If
        End If

        'JANコード
        'JANコード入力妥当性チェック
        '***チェック項目***
        ' 未入力OK、シングルクォートOK
        If StringChkVal(Trim(TextBox6.Text), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox6.BackColor = Color.Salmon
            DataGridErrorMessage &= "JANコードが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox6.Text.Length) > 24 Then
                TextBox6.BackColor = Color.Salmon
                DataGridErrorMessage &= "JANコードが長すぎます。" & vbCr
                Error_Flg = False
            Else
                '商品名格納
                Update_Data(0).JAN = Trim(TextBox6.Text)
            End If
        End If

        '価格
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox5.Text), "^[0-9]+$", _
    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '価格格納
            Update_Data(0).PRICE = Trim(TextBox5.Text)
        Else
            TextBox5.BackColor = Color.Salmon
            DataGridErrorMessage &= "価格が正しくありません。" & vbCr
            Error_Flg = False
        End If

        '仕入金額
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox10.Text), "^[0-9]+$", _
    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '仕入金額格納
            Update_Data(0).PURCHASE_PRICE = Trim(TextBox10.Text)
        Else
            TextBox10.BackColor = Color.Salmon
            DataGridErrorMessage &= "仕入金額が正しくありません。" & vbCr
            Error_Flg = False
        End If

        '免責金額
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox11.Text), "^[0-9]+$", _
    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '免責金額格納
            Update_Data(0).IMMUNITY_PRICE = Trim(TextBox11.Text)
        Else
            TextBox11.BackColor = Color.Salmon
            DataGridErrorMessage &= "免責金額が正しくありません。" & vbCr
            Error_Flg = False
        End If

        'c
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox12.Text), "^[0-9]+$", _
    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '修理金額格納
            Update_Data(0).REPAIR_PRICE = Trim(TextBox12.Text)
        Else
            TextBox12.BackColor = Color.Salmon
            DataGridErrorMessage &= "修理金額が正しくありません。" & vbCr
            Error_Flg = False
        End If


        'ベンダーコード
        'ベンダーコード入力妥当性チェック

        'ベンダーコードはコード入力時にlabel3にベンダー名が表示されているかどうかで確認する。
        If Label3.Text = "" Then
            'TextBox7.BackColor = Color.Salmon
            'DataGridErrorMessage &= "ベンダーコードが正しくありません。" & vbCr
            'Error_Flg = False
        Else
            'ベンダーコード格納
            Update_Data(0).C_CODE = Trim(TextBox7.Text)
        End If


        'ロケーション
        'ロケーション入力妥当性チェック
        '***チェック項目***
        ' 未入力OK、シングルクォートOK
        If StringChkVal(Trim(TextBox8.Text), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox8.BackColor = Color.Salmon
            DataGridErrorMessage &= "ロケーションが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox8.Text.Length) > 100 Then
                TextBox8.BackColor = Color.Salmon
                DataGridErrorMessage &= "ロケーションが長すぎます。" & vbCr
                Error_Flg = False
            Else
                'ロケーション格納
                Update_Data(0).LOCATION = Trim(TextBox8.Text)
            End If
        End If

        '入数
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox2.Text), "^[0-9]+$", _
    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '入数格納
            Update_Data(0).IN_BOX_NUM = Trim(TextBox2.Text)
        Else
            TextBox2.BackColor = Color.Salmon
            DataGridErrorMessage &= "入数が正しくありません。" & vbCr
            Error_Flg = False
        End If

        'マスターカートンサイズ
        'マスターカートンサイズ入力妥当性チェック
        '***チェック項目***
        ' 未入力OK、シングルクォートOK
        If StringChkVal(Trim(TextBox9.Text), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox9.BackColor = Color.Salmon
            DataGridErrorMessage &= "マスターカートンサイズが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox9.Text.Length) > 200 Then
                TextBox9.BackColor = Color.Salmon
                DataGridErrorMessage &= "マスターカートンサイズが長すぎます。" & vbCr
                Error_Flg = False
            Else
                'ロケーション格納
                Update_Data(0).MASTER_CARTON_SIZE = Trim(TextBox9.Text)
            End If
        End If

        '商品区分が選択されていなかったらエラー。
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.BackColor = Color.Salmon
            DataGridErrorMessage &= "商品区分を選択してください。" & vbCr
            Error_Flg = False
        Else
            If Trim(ComboBox4.Text) = "通常商品" Then
                'セット商品格納
                Update_Data(0).PACKAGE_FLG = 0
            ElseIf Trim(ComboBox4.Text) = "セット商品" Then
                Update_Data(0).PACKAGE_FLG = 1
            Else
                ComboBox4.BackColor = Color.Salmon
                DataGridErrorMessage &= "商品区分が正しくありません。" & vbCr
                Error_Flg = False
            End If
        End If

        'プロダクトライン選択されていなかったらエラー。
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.BackColor = Color.Salmon
            DataGridErrorMessage &= "プロダクトラインを選択してください。" & vbCr
            Error_Flg = False
        Else

            PL_Id = ComboBox3.SelectedValue.ToString()
            If PL_Id = 0 Then
                ComboBox3.BackColor = Color.Salmon
                DataGridErrorMessage &= "プロダクトラインが正しくありません。" & vbCr
                Error_Flg = False
            End If

            Update_Data(0).PL_ID = PL_Id

        End If

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub

        End If

        '商品ID
        Update_Data(0).ID = Label4.Text

        Dim Message_Result As DialogResult = MessageBox.Show("商品マスタを登録してもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        Result = Upd_Mitem(Update_Data, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '確定処理が終了したら、画面を閉じて発注検索画面を表示。
        MsgBox("商品マスタの修正が完了しました。")


    End Sub


    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown


        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim CustomerName As String = Nothing
        Dim C_ID As Integer = 0
        Dim C_Code As String = Nothing
        Dim ChkCustomerCodeString As String = Nothing
        Dim Discount_rate As Decimal

        'ベンダーコードをクリアする。
        Label3.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '入力チェック
            If Trim(TextBox7.Text) <> "" Then
                ChkCustomerCodeString = Trim(TextBox7.Text)
                'アイテムコードに'が入力されていたらReplaceする。
                ChkCustomerCodeString = ChkCustomerCodeString.Replace("'", "''")

                Me.ProcessTabKey(Not e.Shift)
                e.Handled = True
                '入力された商品コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetCustomerName(ChkCustomerCodeString, 1, CustomerName, C_ID, C_Code, Discount_rate, Result, ErrorMessage)
                If Result = "True" Then
                    Label3.Text = CustomerName
                    TextBox7.BackColor = Color.White
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox7.Focus()
                    TextBox7.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    Label3.Text = ""
                End If
            End If
        End If

    End Sub
End Class