Public Class zlocation

    Private Sub zlocation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "ロケーション管理"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label3.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Location_Change_Check() As Location_Change_List = Nothing
        Dim Location_Change_Data_Count As Integer = 0
        Dim Error_Flg As Boolean = True

        Dim DataGridErrorMessage As String = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim ChkLocationString As String = Nothing

        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Location_Change_Check(0 To Count)
            'ID
            Location_Change_Check(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(7).Value()
            '商品ID
            Location_Change_Check(Count).I_ID = DataGridView1.Rows(Count).Cells(8).Value()
            'P_ID
            Location_Change_Check(Count).PLACE_ID = DataGridView1.Rows(Count).Cells(9).Value()

            'ロケーション
            ChkLocationString = Nothing
            ChkLocationString = DataGridView1.Rows(Count).Cells(4).Value()
            ChkLocationString = ChkLocationString.Replace("'", "''")

            'ロケーションが101文字以上ならエラーメッセージ表示
            If ChkLocationString.Length > 100 Then
                DataGridErrorMessage &= Count + 1 & "行目の変更後ロケーションの文字数が長すぎます。" & vbCr
            End If
            'ロケーションを格納
            Location_Change_Check(Count).NEW_LOCATION = ChkLocationString
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("ロケーション変更をしてもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        'ロケーション変更Function
        Result = Location_Change(Location_Change_Check, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("良品変更が完了しました。")

        zkensaku.zSearchFLg = True
        Me.Dispose()
        zkensaku.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
        zkensaku.Show()
    End Sub
End Class