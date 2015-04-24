Public Class syukojyuchuhenkou

    Private Sub syukojyuchuhenkou_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "受注変更"

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

    End Sub

    Private Sub syukojyuchuhenkou_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        syukojyuchukensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'ダイアログ設定
        Dim result As DialogResult = MessageBox.Show("データ削除してもよろしいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)
        '何が選択されたか調べる()
        If result = DialogResult.No Then
            '「いいえ」が選択された時 
            Exit Sub
        End If

        MsgBox("削除が完了しました。")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim OutUpdate_Data() As OutShipping_Search_List = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Minus_Check As Boolean = False
        Dim DataGridErrorMessage As String = Nothing
        Dim Error_Flg As Boolean = True

        Dim NUM_Check As Boolean = True

        Dim Check_Num As Integer
        Dim Check_Plan_Num As Integer
        Dim Check_FIX_Num As Integer

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve OutUpdate_Data(0 To Count)
            'ID
            OutUpdate_Data(Count).ID = DataGridView1.Rows(Count).Cells(17).Value()

            '出荷希望数の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                Minus_Check = True
            Else
                Minus_Check = False
            End If

            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, False, Minus_Check, NumChkResult, NumChkErrorMessage) = False Then
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の受注数が正しくありません。" & vbCr

                Error_Flg = False
            Else
                '伝票のみ出力なら、出荷希望数がプラスだとエラー。
                If Trim(DataGridView1.Rows(Count).Cells(4).Value()) > 0 And DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                    DataGridView1(4, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の受注数が正しくありません。" & vbCr
                    Error_Flg = False
                Else
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView1(4, Count).Style.BackColor = Color.White

                    '出荷希望数
                    OutUpdate_Data(Count).NUM = Trim(DataGridView1.Rows(Count).Cells(4).Value())
                End If
            End If

            'もし、出荷指示数より出荷希望数を少なくした場合はエラー
            Check_Num = Trim(DataGridView1.Rows(Count).Cells(4).Value())
            Check_Plan_Num = Trim(DataGridView1.Rows(Count).Cells(5).Value())
            Check_FIX_Num = Trim(DataGridView1.Rows(Count).Cells(6).Value())
            If Check_Num < Check_FIX_Num Then
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の受注数は出荷指示済み数より小さい数値を入力できません。" & vbCr
                Error_Flg = False
            End If

            'もし出荷指示済み数が0の場合は出荷希望数の増減に合わせて出荷予定指示数も変更
            '出荷予定指示数にも希望数と同じ数値を設定する
            If Check_FIX_Num = 0 Then
                OutUpdate_Data(Count).PLAN_NUM = Check_Num
            Else
                '出荷指示済み数が0じゃなかったら出荷指示済み数と出荷希望数の差分を登録
                OutUpdate_Data(Count).PLAN_NUM = Check_Num - Check_FIX_Num
            End If
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        '予定変更 Function
        Result = Upd_OrderData(OutUpdate_Data, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("数量の変更が完了しました。")

        syukojyuchukensaku.SearchFLg = True
        syukojyuchukensaku.Show()
        Me.Dispose()

    End Sub
End Class