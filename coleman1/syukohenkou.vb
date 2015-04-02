Public Class syukohenkou

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        syukokensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '変更ボタン
        'データの変更が行えるのは、備考のみなので、Functionに渡す値はOUT_TBL.IDとRemarksのみ

        Dim Upd_Data() As Out_Upd_List = Nothing
        Dim Count As Integer = 0

        Dim Update_Result As Boolean = True
        Dim Update_Result_Message As String = Nothing

        'DataGridViewからチェックされたデータのIDとコメント情報を配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Upd_Data(0 To Count)
            'ID（OUT_TBL.ID）
            Upd_Data(Count).ID = DataGridView1.Rows(Count).Cells(20).Value()
            'コメント１
            Upd_Data(Count).COMMENT1 = DataGridView1.Rows(Count).Cells(17).Value()
            'コメント２
            Upd_Data(Count).COMMENT2 = DataGridView1.Rows(Count).Cells(18).Value()
        Next

        'データUpdate Function。IDを元にコメントを修正
        Update_Result = Out_Update(Upd_Data, Update_Result, Update_Result_Message)

        If Update_Result = False Then
            MsgBox(Update_Result_Message)
            Exit Sub
        End If

        MsgBox("コメントの修正が完了しました。")
        syukokensaku.sSearchFLg = True
        syukokensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub syukohenkou_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "コメント修正"

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

        '左3項目を固定(伝票番号、オーダー番号、商品コード)
        DataGridView1.Columns(2).Frozen = True

        'ウインドウを表示したとき、１行目の予定数量にフォーカスを移動。
        DataGridView1.Visible = True
        DataGridView1.Select()
        DataGridView1.CurrentCell = DataGridView1(17, 0)
    End Sub

    Private Sub syukohenkou_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

End Class