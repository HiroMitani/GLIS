Public Class PO_cancel

    Private Sub PO_cancel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '左2項目を固定
        DataGridView1.Columns(1).Frozen = True
    End Sub

    Private Sub PO_cancel_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
        PO_kensaku.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Cancel_Check() As PO_In_List = Nothing
        Dim Cancel_Data_Count As Integer = 0

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

        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Cancel_Check(0 To Cancel_Data_Count)
            '商品コード
            Cancel_Check(Cancel_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()

            '発注数
            Cancel_Check(Cancel_Data_Count).PO_NUM = DataGridView1.Rows(Count).Cells(3).Value()

            '納品確定（入庫予定）数
            Cancel_Check(Cancel_Data_Count).IN_NUM = DataGridView1.Rows(Count).Cells(6).Value()


            'キャンセル数
            Cancel_Check(Cancel_Data_Count).CANCEL_NUM = DataGridView1.Rows(Count).Cells(5).Value()

            'キャンセル数
            Cancel_Check(Cancel_Data_Count).CANCEL_NUM = DataGridView1.Rows(Count).Cells(5).Value()

            ChkDateAfter = Nothing

            'ID
            Cancel_Check(Cancel_Data_Count).ID = DataGridView1.Rows(Count).Cells(15).Value()
            'I_ID
            Cancel_Check(Cancel_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(16).Value()

            Cancel_Data_Count += 1
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If


        Dim Message_Result As DialogResult = MessageBox.Show("発注キャンセルをしてもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If


        Result = PO_Upd_Cancel(Cancel_Check, Result, ErrorMessage)

        If Result = False Then
            MsgBox("ErrorMessage")
            Exit Sub
        End If

        'キャンセルが終了したら、画面を閉じて発注検索画面を表示。
        MsgBox("発注のキャンセルが完了しました。")

        PO_kensaku.nSearchFLg = True
        Me.Dispose()
        PO_kensaku.Show()


    End Sub
End Class