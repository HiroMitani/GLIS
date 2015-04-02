Public Class StandardNum_modify

    Private Sub StandardNum_modify_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        PO_Prediction.Show()
        Me.Hide()
    End Sub

    Private Sub StandardNum_modify_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Dim Standart_modify_Data() As Standardnum_Import_List = Nothing
        'Dim ErrorMessage As String = Nothing
        'Dim Result As Boolean = True
        'Dim DataGridErrorMessage As String = Nothing

        'Dim Error_Flg As Boolean = True
        'Dim NumChkErrorMessage As String = Nothing
        'Dim NumChkResult As Boolean = True



        ''データ更新を行う為、予定数量、出荷予定日の妥当性チェックを行う。
        'For Count = 0 To DataGridView1.Rows.Count - 1
        '    ReDim Preserve Standart_modify_Data(0 To Count)

        '    '基準値の入力妥当性チェック
        '    '***チェック項目***
        '    ' 整数値、未入力NG、0の値NG、マイナス値NG
        '    If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(3).Value()), "INTEGER", False, True, False, NumChkResult, NumChkErrorMessage) = False Then
        '        'MsgBox(NumChkErrorMessage)
        '        DataGridView1(3, Count).Style.BackColor = Color.Salmon
        '        DataGridErrorMessage &= Count + 1 & "行目の基準値が正しくありません。" & vbCr

        '        Error_Flg = False
        '    Else
        '        'チェックに問題がなければ背景色を白に戻す。
        '        DataGridView1(3, Count).Style.BackColor = Color.White
        '        '基準値
        '        Standart_modify_Data(Count).STANDARD_NUM = DataGridView1.Rows(Count).Cells(3).Value
        '        '商品ID
        '        Standart_modify_Data(Count).I_ID = DataGridView1.Rows(Count).Cells(4).Value
        '    End If
        'Next

        ''エラーフラグがFalseならメッセージを表示し処理終了
        'If Error_Flg = False Then
        '    MsgBox(DataGridErrorMessage)
        '    Exit Sub
        'End If

        ''UpdateFunction
        'Result = Upd_Standard_Item(Standart_modify_Data, Result, ErrorMessage)
        'If Result = False Then
        '    MsgBox(ErrorMessage)
        '    Exit Sub
        'End If

        ''修正処理が終了したら、画面を閉じて発注予測検索画面を表示。
        'MsgBox("基準値の変更が完了しました。再度、発注予測検索を行ってください。")

        'PO_Prediction.Show()
        'Me.Dispose()

    End Sub
End Class