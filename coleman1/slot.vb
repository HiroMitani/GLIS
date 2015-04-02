Public Class slot

    Private Sub slot_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "ロット番号入力"

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
        syukokensaku.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Out_Lot_Data() As Lot_List = Nothing
        Dim Out_Data_Count As Integer = 0

        Dim Check_Count As Integer

        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Check_Count += 1
            End If
        Next

        If Check_Count <> 1 Then
            MsgBox("チェックは１つのみいれてください。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Out_Lot_Data(0 To Out_Data_Count)
            'OUT_ID
            Out_Lot_Data(Out_Data_Count).OUT_ID = Label6.Text
            'No
            Out_Lot_Data(Out_Data_Count).NO = DataGridView1.Rows(Count).Cells(0).Value()
            'ロット番号
            Out_Lot_Data(Out_Data_Count).LOT_NUMBER = DataGridView1.Rows(Count).Cells(1).Value()
            '保証書番号
            Out_Lot_Data(Out_Data_Count).WARRANTY_CARD_NUMBER = DataGridView1.Rows(Count).Cells(2).Value()
            Out_Data_Count += 1
        Next

        'ロット番号登録・更新 Function
        Result = Out_LotRegist(Out_Lot_Data, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("ロット番号・保証書番号登録が完了しました。")

        Me.Dispose()
        syukokensaku.Show()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '1行目のロット番号が入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(0).Cells(1).Value() = "" Then
            MsgBox("1行目のロット番号が未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。
        For Count = 0 To DataGridView1.Rows.Count - 1
            DataGridView1(1, Count).Value = DataGridView1(1, 0).Value
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '1行目の保証書番号が入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(Row).Cells(2).Value() = "" Then
            MsgBox("1行目の保証書番号が未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。
        For Count = 0 To DataGridView1.Rows.Count - 1
            DataGridView1(2, Count).Value = DataGridView1(2, 0).Value
        Next
    End Sub
End Class