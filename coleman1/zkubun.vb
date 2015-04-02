Public Class zkubun

    Dim Form_Load As Boolean = False

    Private Sub zkubun_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        '登録ボタンが押されたら()

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim IStatus_Change_Check() As IStatus_Change_List = Nothing
        Dim IStatus_Change_Data_Count As Integer = 0

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        Dim ChkRemarksString As String = Nothing

        Dim DataGridErrorMessage As String = Nothing
        Dim ChkMemo As String = Nothing


        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve IStatus_Change_Check(0 To Count)
            '商品ID
            IStatus_Change_Check(Count).I_ID = DataGridView1.Rows(Count).Cells(12).Value()
            '商品コード
            IStatus_Change_Check(Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
            '変更数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                DataGridErrorMessage &= Count + 1 & "行目の変更数量は" & NumChkErrorMessage & vbCr

                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(4, Count).Style.BackColor = Color.White
                '変更後数量
                IStatus_Change_Check(Count).CHANGE_NUM = DataGridView1.Rows(Count).Cells(4).Value()
                '現数量
                IStatus_Change_Check(Count).NUM = DataGridView1.Rows(Count).Cells(3).Value()
            End If

            ' 現状不良区分
            If DataGridView1.Rows(Count).Cells(5).Value() = "良品" Then
                IStatus_Change_Check(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(5).Value() = "不良品" Then
                IStatus_Change_Check(Count).I_STATUS = 2
            ElseIf DataGridView1.Rows(Count).Cells(5).Value() = "保管品" Then
                IStatus_Change_Check(Count).I_STATUS = 3
            End If
            '変更後不良区分
            If DataGridView1.Rows(Count).Cells(6).Value() = "良品" Then
                IStatus_Change_Check(Count).CHANGE_I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(6).Value() = "不良品" Then
                IStatus_Change_Check(Count).CHANGE_I_STATUS = 2
            ElseIf DataGridView1.Rows(Count).Cells(6).Value() = "保管品" Then
                IStatus_Change_Check(Count).CHANGE_I_STATUS = 3
            End If

            '変更後ロケーション
            'メモに'があればPeplaceする。
            ChkRemarksString = DataGridView1.Rows(Count).Cells(8).Value()
            ChkRemarksString = ChkRemarksString.Replace("'", "''")
            IStatus_Change_Check(Count).NEW_LOCATION = ChkRemarksString

            'メモが101文字以上ならエラーメッセージ表示
            If ChkRemarksString.Length > 100 Then
                DataGridErrorMessage &= Count + 1 & "行目の変更後ロケーションの文字数が長すぎます。" & vbCr
            End If

            '倉庫
            IStatus_Change_Check(Count).PLACE = DataGridView1.Rows(Count).Cells(10).Value()

            'STOCK_ID
            IStatus_Change_Check(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(11).Value()

            'PLACE_ID
            IStatus_Change_Check(Count).PLACE_ID = DataGridView1.Rows(Count).Cells(13).Value()
        Next



        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("良品区分変更をしてもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        '良品変更Function
        Result = IStatus_Change(IStatus_Change_Check, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("良品変更が完了しました。")

        zkensaku.zSearchFLg = True


        Me.Dispose()
        zkensaku.Show()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        zkensaku.Show()
    End Sub
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If Form_Load = True Then
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            Dim Col As Integer = DataGridView1.CurrentCell.ColumnIndex
            Dim ChkMemo As String = Nothing
            Dim NumChkErrorMessage As String = Nothing
            Dim NumChkResult As Boolean = True

            If Col = 4 Then
                'NULLはNG、0もNG、マイナスの値NG
                If NumChkVal(Trim(DataGridView1.Rows(Row).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then

                    DataGridView1(4, Row).Style.BackColor = Color.Salmon
                    MsgBox(NumChkErrorMessage)
                    Exit Sub
                Else

                    '数量と変更後をチェック
                    '変更数量が０以下ならエラー。
                    If DataGridView1(4, Row).Value < 0 Then
                        MsgBox("変更数量をマイナスにはできません。")
                        DataGridView1(4, Row).Style.BackColor = Color.Salmon
                        Exit Sub
                    End If
                    '現在の数量より大きい数値が入力されるとエラー。
                    If DataGridView1(3, Row).Value < DataGridView1(4, Row).Value Then
                        MsgBox("現在庫数以上の数値は入力できません。")
                        DataGridView1(4, Row).Style.BackColor = Color.Salmon
                        Exit Sub
                    End If
                    DataGridView1(4, Row).Style.BackColor = Color.White
                End If
            ElseIf Col = 8 Then
                '変更後ロケーション
                ChkMemo = DataGridView1(8, Row).Value
                If ChkMemo.Length > 100 Then
                    MsgBox("入力した文字数が長すぎます。")
                    DataGridView1(8, Row).Style.BackColor = Color.Salmon
                    Exit Sub

                End If

            End If
        End If

    End Sub

    Private Sub zkubun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "良品変更"

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
        DataGridView1.CurrentCell = DataGridView1(4, 0)

        Form_Load = True

    End Sub
End Class