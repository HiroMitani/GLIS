Public Class ztanaoroshi

    Dim Form_Load As Boolean = False
    Private Sub ztanaoroshi_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        '棚卸ボタンが押されたら

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Tanaoroshi_Check() As Tanaoroshi_List = Nothing
        Dim Tanaoroshi_Data_Count As Integer = 0

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        Dim ChkRemarksString As String = Nothing

        Dim DataGridErrorMessage As String = Nothing

        Dim ChkMemo As String = Nothing


        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Tanaoroshi_Check(0 To Count)
            '商品ID
            Tanaoroshi_Check(Count).I_ID = DataGridView1.Rows(Count).Cells(12).Value()
            '商品コード
            Tanaoroshi_Check(Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
            '棚卸後数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の棚卸後数量は" & NumChkErrorMessage & vbCr
                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(4, Count).Style.BackColor = Color.White
                '棚卸後数量
                Tanaoroshi_Check(Count).NEW_NUM = DataGridView1.Rows(Count).Cells(4).Value()
                '現数量
                Tanaoroshi_Check(Count).NUM = DataGridView1.Rows(Count).Cells(3).Value()
            End If
            'メモ

            If DataGridView1.Rows(Count).Cells(6).Value() <> "" Then

                '501文字以上ならエラーメッセージを表示
                ChkMemo = DataGridView1.Rows(Count).Cells(6).Value()
                If ChkMemo.Length > 500 Then
                    DataGridErrorMessage &= Count + 1 & "行目のメモが長すぎます。" & vbCr
                    Error_Flg = False
                End If

                'メモに'があればPeplaceする。
                ChkRemarksString = DataGridView1.Rows(Count).Cells(6).Value()
                ChkRemarksString = ChkRemarksString.Replace("'", "''")
                Tanaoroshi_Check(Count).REMARKS = ChkRemarksString

            End If

            '倉庫
            Tanaoroshi_Check(Count).PLACE = DataGridView1.Rows(Count).Cells(10).Value()

            'STOCK_ID
            Tanaoroshi_Check(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(11).Value()

            '不良区分
            If DataGridView1.Rows(Count).Cells(9).Value() = "良品" Then
                Tanaoroshi_Check(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(9).Value() = "不良品" Then
                Tanaoroshi_Check(Count).I_STATUS = 2
            End If
        Next


        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("棚卸をしてもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        '棚卸Function
        Result = Tanaoroshi(Tanaoroshi_Check, 3, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("棚卸が完了しました。")

        zkensaku.zSearchFLg = True

        Me.Dispose()
        zkensaku.Show()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
        zkensaku.Show()
    End Sub

    Private Sub ztanaoroshi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'ウインドウを表示したとき、１行目の数量にフォーカスを移動。
        DataGridView1.Visible = True
        DataGridView1.Select()
        DataGridView1.CurrentCell = DataGridView1(4, 0)
        Form_Load = True
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If Form_Load = True Then
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            Dim Col As Integer = DataGridView1.CurrentCell.ColumnIndex

            Dim NumChkErrorMessage As String = Nothing
            Dim NumChkResult As Boolean = True
            Dim ChkMemo As String = Nothing


            '棚卸後数量なら
            If Col = 4 Then
                '数量から棚卸数量をひいて、差分に結果を表示する。
                'もし在庫数量欄にマイナスを入れたらメッセージを表示する。
                If NumChkVal(Trim(DataGridView1.Rows(Row).Cells(4).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView1(4, Row).Style.BackColor = Color.Salmon
                    MsgBox(NumChkErrorMessage)
                    Exit Sub
                Else
                    If DataGridView1(4, Row).Value < 0 Then
                        MsgBox("在庫数をマイナスにはできません。")
                        DataGridView1(4, Row).Style.BackColor = Color.Salmon

                        Exit Sub

                    End If
                    DataGridView1(4, Row).Style.BackColor = Color.White
                    DataGridView1(5, Row).Value = DataGridView1(4, Row).Value - DataGridView1(3, Row).Value
                End If
                'メモなら
            ElseIf Col = 6 Then
                ChkMemo = Trim(DataGridView1.Rows(Row).Cells(6).Value())
                If ChkMemo.Length > 500 Then
                    DataGridView1(6, Row).Style.BackColor = Color.Salmon
                    MsgBox("入力したメモが長すぎます。")
                    Exit Sub

                End If
            End If

        End If
    End Sub

End Class