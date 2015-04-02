Public Class nhenkou

    Private Sub nhenkou_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim Disp_Title As String = "入庫予定変更"

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
        DataGridView1.CurrentCell = DataGridView1(2, 0)

        '左4項目を固定(ドキュメント№、商品コード、予定数量、入荷予定日)
        DataGridView1.Columns(3).Frozen = True

    End Sub

    Private Sub nhenkou_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        nkensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Error_Flg As Boolean = True
        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim Upd_Data() As Upd_List = Nothing

        Dim Upd_Data_Count As Integer = 0

        Dim ChkDateAfter As String = Nothing
        Dim Count As Integer
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True

        Dim DataGridErrorMessage As String = Nothing

        Dim ChkDocNoString As String = Nothing

        '2012/02/10 変更画面でドキュメント№の修正が出来るようにする対応に伴う改修
        'ドキュメント№が空白ならメッセージ表示。
        If Trim(TextBox1.Text) = "" Then
            MsgBox("ドキュメント№は必須項目です。")
            Exit Sub
        Else
            ChkDocNoString = Trim(TextBox1.Text)
            'ドキュメント№に'が入力されていたらReplaceする。
            ChkDocNoString = ChkDocNoString.Replace("'", "''")
        End If

        'データ更新を行う為、予定数量、出荷予定日の妥当性チェックを行う。
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Upd_Data(0 To Count)

            '予定数量数の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(2).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(2, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の予定数量が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(2, Count).Style.BackColor = Color.White
                '入荷数量
                Upd_Data(Count).NUM = DataGridView1.Rows(Count).Cells(2).Value
            End If

            ChkDateAfter = Nothing
            '入荷予定日の妥当性チェック
            If DateChkVal(Trim(DataGridView1.Rows(Count).Cells(3).Value()), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(3, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の入荷予定日が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(3, Count).Style.BackColor = Color.White
                '入荷予定日
                Upd_Data(Count).N_DATE = ChkDateAfter

                'DataGridViewにも反映
                DataGridView1.Rows(Count).Cells(3).Value = ChkDateAfter
            End If

            'カテゴリー
            If DataGridView1.Rows(Count).Cells(9).Value = "通常入荷" Then
                Upd_Data(Count).CATEGORY = 1
            Else
                Upd_Data(Count).CATEGORY = 2
            End If

            '不良区分
            If DataGridView1.Rows(Count).Cells(10).Value = "良品" Then
                Upd_Data(Count).DEFECT_TYPE = 1
            Else
                Upd_Data(Count).DEFECT_TYPE = 2
            End If
            'ID
            Upd_Data(Count).ID = DataGridView1.Rows(Count).Cells(11).Value
            'Detail_ID
            Upd_Data(Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(12).Value
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        '予定変更 UpdateFunction
        Result = Upd_In_Item(Upd_Data, ChkDocNoString, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '変更処理が終了したら、画面を閉じて入庫予定検索画面を表示。
        MsgBox("予定変更が完了しました。")

        nkensaku.nSearchFLg = True
        Me.Dispose()
        nkensaku.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '現在の行の入荷予定日が入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(Row).Cells(3).Value() = "" Then
            MsgBox(Row + 1 & "行目の入荷予定日が未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。

        For Count = Row To DataGridView1.Rows.Count - 1
            DataGridView1(3, Count).Value = DataGridView1(3, Row).Value
        Next
    End Sub
End Class