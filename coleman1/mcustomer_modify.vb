Public Class mcustomer_modify
    Private Sub mcustomer_modify_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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


    Private Sub mcustomer_modify_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "企業マスタ登録修正"

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

        TextBox1.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim C_Code As String = Nothing

        Dim Search_Data() As M_Customer_List = Nothing

        '入力エリア（検索結果）をクリアする。
        '企業コード
        TextBox3.Text = ""
        '企業名
        TextBox4.Text = ""
        '伝票タイプ
        ComboBox1.Text = "カンセキ"
        '納品先名
        TextBox5.Text = ""
        '納品先住所
        TextBox6.Text = ""
        '納品先郵便番号
        TextBox7.Text = ""
        '納品先TEL
        TextBox8.Text = ""
        '納品先FAX
        TextBox9.Text = ""
        '請求先コード
        TextBox2.Text = ""
        '掛け率
        TextBox10.Text = ""

        'カスタマー種別
        ComboBox1.Text = "ベンダー"

        '納品先別出荷リスト出力フラグ
        ComboBox3.Text = "出力する"

        If TextBox1.Text = "" Then
            MsgBox("企業コードを入力してください。")
            'カーソルを検索条件の商品コードへ。
            TextBox1.Focus()
            Exit Sub
        End If

        '商品コード
        C_Code = Trim(TextBox1.Text).Replace("'", "''")
        C_Code = Trim(C_Code).Replace("\", "\\")

        Result = Get_Mcustomer_Modify(C_Code, Search_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        If Search_Data(0).D_NAME = "" Then
            MsgBox("データがみつかりません")
        End If

        '入力欄に値をセットする。

        '企業コード
        TextBox3.Text = Search_Data(0).C_CODE
        '企業名
        TextBox4.Text = Search_Data(0).C_NAME
        '伝票タイプ
        If Search_Data(0).SHEET_TYPE = "chain1" Then
            ComboBox1.Text = "カンセキ"
        ElseIf Search_Data(0).SHEET_TYPE = "chain3" Then
            ComboBox1.Text = "T3"
        ElseIf Search_Data(0).SHEET_TYPE = "takamiya" Then
            ComboBox1.Text = "タカミヤ"
        End If

        '納品先名
        TextBox5.Text = Search_Data(0).D_NAME
        '納品先郵便番号
        TextBox6.Text = Search_Data(0).D_ZIP
        '納品先住所
        TextBox7.Text = Search_Data(0).D_ADDRESS
        '納品先TEL
        TextBox8.Text = Search_Data(0).D_TEL
        '納品先FAX
        TextBox9.Text = Search_Data(0).D_FAX

        'カスタマー種別
        ComboBox2.Text = Search_Data(0).CUSTOMER_TYPE

        '納品先別出荷リスト
        If Search_Data(0).DELIVERY_PRT_FLG = "1" Then
            ComboBox3.Text = "出力する"
        ElseIf Search_Data(0).DELIVERY_PRT_FLG = "0" Then
            ComboBox3.Text = "出力しない"
        End If

        '請求先コード
        TextBox2.Text = Search_Data(0).CLAIM_CODE
        '掛け率
        TextBox10.Text = Search_Data(0).DISCOUNT_RATE

        'ID
        Label12.Text = Search_Data(0).ID

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Update_Data() As M_Customer_List = Nothing
        Dim DataGridErrorMessage As String = Nothing
        Dim Error_Flg As Boolean = True

        Dim ChkStringAfter As String = Nothing
        Dim StringChkErrorMessage As String = Nothing
        Dim StringChkResult As Boolean = True

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        ReDim Update_Data(0 To 0)

        '商品IDが入っていなかったら、検索ボタンを押していないことになるのでメッセージ表示
        If Label12.Text = "" Then
            MsgBox("検索をしてから修正ボタンを押してください。")
            Exit Sub
        End If

        '入力チェック

        '企業コード２４
        '企業コードの入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox3.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox3.BackColor = Color.Salmon
            DataGridErrorMessage &= "企業コードが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox3.Text.Length) > 24 Then
                TextBox3.BackColor = Color.Salmon
                DataGridErrorMessage &= "企業コードが長すぎます。" & vbCr
                Error_Flg = False
            Else
                '企業コード格納
                Update_Data(0).C_CODE = Trim(TextBox3.Text)
            End If
        End If

        '請求先コード２４
        '請求先コードの入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox2.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox2.BackColor = Color.Salmon
            DataGridErrorMessage &= "請求先コードが正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox2.Text.Length) > 24 Then
                TextBox2.BackColor = Color.Salmon
                DataGridErrorMessage &= "請求先コードが長すぎます。" & vbCr
                Error_Flg = False
            Else
                '企業コード格納
                Update_Data(0).CLAIM_CODE = Trim(TextBox2.Text)
            End If
        End If

        '企業名１００
        '企業名の入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox4.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox4.BackColor = Color.Salmon
            DataGridErrorMessage &= "企業名が正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox4.Text.Length) > 100 Then
                TextBox4.BackColor = Color.Salmon
                DataGridErrorMessage &= "企業名が長すぎます。" & vbCr
                Error_Flg = False
            Else
                '企業コード格納
                Update_Data(0).C_NAME = Trim(TextBox4.Text)
            End If
        End If

        '納品先名の入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox5.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox5.BackColor = Color.Salmon
            DataGridErrorMessage &= "納品先名が正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox5.Text.Length) > 100 Then
                TextBox5.BackColor = Color.Salmon
                DataGridErrorMessage &= "納品先名が長すぎます。" & vbCr
                Error_Flg = False
            Else
                '企業コード格納
                Update_Data(0).D_NAME = Trim(TextBox5.Text)
            End If
        End If

        '郵便番号。文字列が8文字以内かチェックをする。
        If Trim(TextBox6.Text).Length > 8 Then
            'エラー
            TextBox6.BackColor = Color.Salmon
            DataGridErrorMessage &= "郵便番号が正しくありません。" & vbCr
            Error_Flg = False
        Else
            '郵便番号チェック
            If PostChkVal(Trim(TextBox6.Text), True, StringChkResult, StringChkErrorMessage) = False Then
                'エラー
                TextBox6.BackColor = Color.Salmon
                DataGridErrorMessage &= "郵便番号が正しくありません。" & vbCr
                Error_Flg = False
            Else
                'もし7桁の場合の場合（9999999）は、999-9999の形式に変換。
                Dim postdata As String = Nothing
                If Trim(TextBox6.Text).Length = 7 Then
                    '前後スペースを省い格納する。
                    TextBox6.Text = Trim(TextBox6.Text)
                    'ハイフンをつけて格納する。
                    TextBox6.Text = TextBox6.Text.Substring(0, 3) & "-" & TextBox6.Text.Substring(3, 4)

                End If
                '郵便番号コード格納
                Update_Data(0).D_ZIP = Trim(TextBox6.Text)

            End If
        End If

        '住所の入力妥当性チェック
        '***チェック項目***
        ' 未入力NG、シングルクォートOK
        If StringChkVal(Trim(TextBox7.Text), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            'MsgBox(NumChkErrorMessage)
            TextBox7.BackColor = Color.Salmon
            DataGridErrorMessage &= "住所が正しくありません。" & vbCr
            Error_Flg = False
        Else
            If Trim(TextBox7.Text.Length) > 200 Then
                TextBox7.BackColor = Color.Salmon
                DataGridErrorMessage &= "住所が長すぎます。" & vbCr
                Error_Flg = False
            Else
                '住所格納
                Update_Data(0).D_ADDRESS = Trim(TextBox7.Text)
            End If
        End If

        '電話番号の場合、文字列が30文字以内かチェックをする。
        If Trim(TextBox8.Text.Length) > 30 Then
            'エラー内容を格納する。
            TextBox8.BackColor = Color.Salmon
            DataGridErrorMessage &= "電話番号が正しくありません。" & vbCr
            Error_Flg = False
        Else

            '半角数値とハイフンのみかチェック
            If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox8.Text), "^[0-9-]+$", _
                System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                '電話番号格納
                TextBox8.BackColor = Color.White
                Update_Data(0).D_TEL = Trim(TextBox8.Text)
            Else
                'エラー内容を格納する。
                TextBox8.BackColor = Color.Salmon
                DataGridErrorMessage &= "電話番号が正しくありません。" & vbCr
                Error_Flg = False
            End If

        End If

        'IDを格納
        Update_Data(0).ID = Trim(Label12.Text)

        'FAXが未入力なら空白を入れて登録
        If Trim(TextBox9.Text) = "" Then
            Update_Data(0).D_FAX = ""
        Else
            'FAXの場合、文字列が30文字以内かチェックをする。
            If Trim(TextBox9.Text.Length) > 30 Then
                'エラー内容を格納する。
                TextBox9.BackColor = Color.Salmon
                DataGridErrorMessage &= "FAXが正しくありません。" & vbCr
                Error_Flg = False
            Else

                '半角数値とハイフンのみかチェック
                If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox9.Text), "^[0-9-]+$", _
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                    'FAX番号格納
                    TextBox9.BackColor = Color.White
                    Update_Data(0).D_FAX = Trim(TextBox9.Text)
                Else
                    'エラー内容を格納する。
                    TextBox9.BackColor = Color.Salmon
                    DataGridErrorMessage &= "FAXが正しくありません。" & vbCr
                    Error_Flg = False
                End If

            End If
        End If

        '伝票タイプが選択されていなかったらエラー。
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.BackColor = Color.Salmon
            DataGridErrorMessage &= "伝票タイプを選択してください。" & vbCr
            Error_Flg = False
        Else
            If Trim(ComboBox1.Text) = "カンセキ" Then
                Update_Data(0).SHEET_TYPE = "chain1"
            ElseIf Trim(ComboBox1.Text) = "T3" Then
                Update_Data(0).SHEET_TYPE = "chain3"
            ElseIf Trim(ComboBox1.Text) = "タカミヤ" Then
                Update_Data(0).SHEET_TYPE = "takamiya"
            Else
                ComboBox1.BackColor = Color.Salmon
                DataGridErrorMessage &= "伝票タイプが正しくありません。" & vbCr
                Error_Flg = False
            End If
        End If

        'カスタマー種別
        If Trim(ComboBox2.Text) = "" Then
            ComboBox1.BackColor = Color.Salmon
            DataGridErrorMessage &= "カスタマー種別を選択してください。" & vbCr
            Error_Flg = False
        Else
            If Trim(ComboBox2.Text) = "ベンダー" Or Trim(ComboBox2.Text) = "カスタマー" Or Trim(ComboBox2.Text) = "ウェアハウス" Then
                'カスタマー種別格納
                Update_Data(0).CUSTOMER_TYPE = Trim(ComboBox2.Text)

            Else
                ComboBox2.BackColor = Color.Salmon
                DataGridErrorMessage &= "カスタマー種別が正しくありません。" & vbCr
                Error_Flg = False
            End If
        End If

        '納品先別出荷リスト出力区分
        If Trim(ComboBox3.Text) = "" Then
            ComboBox1.BackColor = Color.Salmon
            DataGridErrorMessage &= "納品先別出荷リスト出力区分を選択してください。" & vbCr
            Error_Flg = False
        Else
            If Trim(ComboBox3.Text) = "出力する" Then
                'カスタマー種別格納
                Update_Data(0).DELIVERY_PRT_FLG = 1
            ElseIf Trim(ComboBox3.Text) = "出力しない" Then
                'カスタマー種別格納
                Update_Data(0).DELIVERY_PRT_FLG = 0
            End If
        End If

        '掛け率が半角数字のみかチェック
        If System.Text.RegularExpressions.Regex.IsMatch(Trim(TextBox10.Text), "^[0-9|.]+$", _
            System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '掛け率
            TextBox10.BackColor = Color.White
            Update_Data(0).DISCOUNT_RATE = Trim(TextBox10.Text)
        Else
            'エラー内容を格納する。
            TextBox10.BackColor = Color.Salmon
            DataGridErrorMessage &= "掛け率が正しくありません。" & vbCr
            Error_Flg = False
        End If

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("企業マスタを登録してもよろしいですか？", _
             "確認", _
             MessageBoxButtons.YesNo, _
             MessageBoxIcon.Exclamation, _
             MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        Result = Upd_Mcustomer(Update_Data, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '処理が終了
        MsgBox("企業マスタの修正が完了しました。")

    End Sub

End Class