Public Class PO_nyuukoyotei
    'False:検索ボタンを押したらFalse
    '入庫予定を行ったら、再度検索を行わせる為のFlg
    Public nSearchFLg As Boolean = False

    Dim Form_Load As Boolean = False

    Dim ListData() As Place_List = Nothing

    Private Sub PO_nyuukoyotei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "発注→納品確定（入庫予定）登録"

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

        '左9項目を固定
        DataGridView1.Columns(8).Frozen = True

        '倉庫の一覧を取得する
        Result = GetPLACEList(ListData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If


        Dim dtTbl As New DataTable
        dtTbl.Columns.Add("Display", GetType(String))
        dtTbl.Columns.Add("Value", GetType(Integer))
        For Count = 0 To ListData.Length - 1
            dtTbl.Rows.Add(ListData(Count).NAME, Count)
        Next

        Dim cbc As New DataGridViewComboBoxColumn
        cbc = CType(DataGridView1.Columns(8), DataGridViewComboBoxColumn)

        cbc.DataSource = dtTbl
        cbc.ValueMember = "Value"
        cbc.DisplayMember = "Display"

        Form_Load = True
    End Sub
    Private Sub PO_nyuukoyotei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        Me.Dispose()
        PO_kensaku.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '現在の行のロケーションが入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(Row).Cells(5).Value() = "" Then
            MsgBox(Row + 1 & "行目のインボイスNoが未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。

        For Count = Row To DataGridView1.Rows.Count - 1
            DataGridView1(5, Count).Value = DataGridView1(5, Row).Value
        Next

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '現在の行のロケーションが入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(Row).Cells(7).Value() = "" Then
            MsgBox(Row + 1 & "行目の入庫予定日が未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。
        For Count = Row To DataGridView1.Rows.Count - 1
            DataGridView1(7, Count).Value = DataGridView1(7, Row).Value
        Next

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim In_Schedule_Check() As PO_In_List = Nothing
        Dim In_Schedule_Data_Count As Integer = 0

        Dim StringChkErrorMessage As String = Nothing
        Dim StringChkResult As Boolean = True

        Dim ChkStringAfter As String = Nothing
        Dim ChkDateAfter As String = Nothing
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True
        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim DataGridErrorMessage As String = Nothing
        Dim Error_Flg As Boolean = True

        Dim Category As Integer
        Dim Defect_Type As Integer

        Dim Nyuuko_NUM_Check As Integer
        Dim Cancel_NUM_Check As Integer
        Dim Nyuuko_Yotei_NUM_Check As Integer
        Dim hacchuu_NUM As Integer

        Dim Check_Result As Boolean = False

        Dim PlaceData() As Place_List = Nothing

        Dim Place_Check_Flg As Boolean = True

        Dim SetPlace As Integer

        '倉庫が指定されていなかったらめっせーじ表示。
        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            ' If DataGridView1.Rows(Count).Cells(8).Value() = Nothing Then
            If DataGridView1.Item(8, Count).FormattedValue = "" Then
                Place_Check_Flg = False
            End If
        Next
        If Place_Check_Flg = False Then
            MsgBox("倉庫が設定されていないデータが存在します。倉庫は必ず選択してください。")
            Exit Sub
        End If

        '倉庫情報の取得
        Result = GetPLACEList(PlaceData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If


        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve In_Schedule_Check(0 To In_Schedule_Data_Count)
            '商品コード
            In_Schedule_Check(In_Schedule_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()

            'インボイスNoの入力妥当性チェック
            '***チェック項目***
            ' 未入力NG、シングルクォートOK
            If StringChkVal(Trim(DataGridView1.Rows(Count).Cells(5).Value()), False, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(5, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目のインボイスNoが正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(5, Count).Style.BackColor = Color.White
                'インボイスNo
                In_Schedule_Check(In_Schedule_Data_Count).INVOICE_NO = DataGridView1.Rows(Count).Cells(5).Value()
            End If

            '納品確定数の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(6).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(6, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の納品確定数の値が正しくありません。" & vbCr
                Error_Flg = False
            Else

                '入力値 <= 発注数 - 納品確定（入庫予定）数 - 発注キャンセル数

                '発注数
                hacchuu_NUM = Trim(Integer.Parse(DataGridView1.Rows(Count).Cells(3).Value()))
                '納品確定数（入力箇所）
                Nyuuko_NUM_Check = Trim(Integer.Parse(DataGridView1.Rows(Count).Cells(6).Value()))
                '納品確定（入庫予定）数
                Nyuuko_Yotei_NUM_Check = Trim(Integer.Parse(DataGridView1.Rows(Count).Cells(9).Value()))
                '発注キャンセル数
                Cancel_NUM_Check = Trim(Integer.Parse(DataGridView1.Rows(Count).Cells(11).Value()))

                If Nyuuko_NUM_Check <= hacchuu_NUM - Nyuuko_Yotei_NUM_Check - Cancel_NUM_Check Then
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView1(6, Count).Style.BackColor = Color.White
                    '納品確定数
                    In_Schedule_Check(In_Schedule_Data_Count).IN_NUM = DataGridView1.Rows(Count).Cells(6).Value()
                Else
                    DataGridView1(6, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の納品確定数が発注数を越えています" & vbCr
                    Error_Flg = False
                End If
            End If
                ChkDateAfter = Nothing

                '入庫予定日の妥当性チェック
            If DateChkVal(Trim(DataGridView1.Rows(Count).Cells(7).Value()), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(7, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の入庫予定日が正しくありません。" & vbCr
                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(7, Count).Style.BackColor = Color.White
                '入庫予定日
                In_Schedule_Check(In_Schedule_Data_Count).IN_DATE = ChkDateAfter
                'DataGridViewにも反映
                DataGridView1.Rows(Count).Cells(7).Value = ChkDateAfter
            End If

            '倉庫
            'If DataGridView1.Rows(Count).Cells(8).Value() = "平和島" Then
            '    In_Schedule_Check(In_Schedule_Data_Count).PLACE = 1
            'ElseIf DataGridView1.Rows(Count).Cells(8).Value() = "返品センター" Then
            '    In_Schedule_Check(In_Schedule_Data_Count).PLACE = 2
            'ElseIf DataGridView1.Rows(Count).Cells(8).Value() = "営業倉庫" Then
            '    In_Schedule_Check(In_Schedule_Data_Count).PLACE = 3
            'ElseIf DataGridView1.Rows(Count).Cells(8).Value() = "八潮" Then
            'End If

            SetPlace = 0
            For pcount = 0 To PlaceData.Length - 1
                If DataGridView1.Item(8, Count).FormattedValue = PlaceData(pcount).NAME Then
                    SetPlace = PlaceData(pcount).ID
                End If
            Next

            'In_Schedule_Check(In_Schedule_Data_Count).PLACE = PlaceData(DataGridView1.Rows(Count).Cells(8).Value()).ID
            In_Schedule_Check(In_Schedule_Data_Count).PLACE = SetPlace

            'ベンダー
            In_Schedule_Check(In_Schedule_Data_Count).VENDER_CODE = DataGridView1.Rows(Count).Cells(13).Value()

            'ID
            In_Schedule_Check(In_Schedule_Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
            'I_ID
            In_Schedule_Check(In_Schedule_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(18).Value()

            In_Schedule_Data_Count += 1
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        'インボイスNoがすでにIN_HEADERに登録されているかチェックする。
        Result = PO_Invoice_Check(In_Schedule_Check, Check_Result, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        If Check_Result = True Then



            Dim INVoice_Message_Result As DialogResult = MessageBox.Show("入力されたインボイスNoはすでに登録されていますが処理を続けてよろしいですか？", _
                         "確認", _
                         MessageBoxButtons.YesNo, _
                         MessageBoxIcon.Exclamation, _
                         MessageBoxDefaultButton.Button2)

            '何が選択されたか調べる 
            If INVoice_Message_Result = DialogResult.No Then
                'Noを選択した場合、処理終了
                Exit Sub
            End If

        End If

        Dim Message_Result As DialogResult = MessageBox.Show("入庫予定を登録してもよろしいですか？", _
                             "確認", _
                             MessageBoxButtons.YesNo, _
                             MessageBoxIcon.Exclamation, _
                             MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        Defect_Type = 1
        Category = 1
        Result = PO_InsItem(In_Schedule_Check, Defect_Type, Category, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '確定処理が終了したら、画面を閉じて発注検索画面を表示。
        MsgBox("入庫予定の登録が完了しました。")

        PO_kensaku.nSearchFLg = True
        Me.Dispose()
        PO_kensaku.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '件数分、発注数を納品確定数に入れる。
        For Count = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(Count).Cells(8).Value() = DataGridView1.Rows(Count).Cells(5).Value()
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '現在の行のロケーションが入れられてなかったらエラーメッセージ表示
        'If DataGridView1.Rows(Row).Cells(8).Value() = "" Then
        If DataGridView1.Item(8, Row).FormattedValue = "" Then
            MsgBox(Row + 1 & "行目の倉庫が未選択です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。
        For Count = Row To DataGridView1.Rows.Count - 1
            DataGridView1(8, Count).Value = DataGridView1(8, Row).Value
        Next
    End Sub

End Class