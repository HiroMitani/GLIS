Public Class ztyousei

    Dim Form_Load As Boolean = False

    Private Sub ztyousei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        Me.Dispose()
        zkensaku.Show()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '在庫調整ボタンが押されたら

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Tyousei_Check() As Tanaoroshi_List = Nothing
        Dim TTyouseianaoroshi_Data_Count As Integer = 0

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        Dim DataGridErrorMessage As String = Nothing

        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve Tyousei_Check(0 To Count)
            '商品ID
            Tyousei_Check(Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
            '商品コード
            Tyousei_Check(Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
            '棚卸後数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, True, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の調整後実数量は" & NumChkErrorMessage & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(4, Count).Style.BackColor = Color.White
                '棚卸後数量
                Tyousei_Check(Count).NEW_NUM = DataGridView1.Rows(Count).Cells(4).Value()
                '現数量
                Tyousei_Check(Count).NUM = DataGridView1.Rows(Count).Cells(3).Value()
            End If
            '倉庫
            Tyousei_Check(Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()

            'STOCK_ID
            Tyousei_Check(Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
            'P_ID
            Tyousei_Check(Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()


            '不良区分
            If DataGridView1.Rows(Count).Cells(8).Value() = "良品" Then
                Tyousei_Check(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(8).Value() = "不良品" Then
                Tyousei_Check(Count).I_STATUS = 2
            End If
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("在庫調整をしてもよろしいですか？", _
                     "確認", _
                     MessageBoxButtons.YesNo, _
                     MessageBoxIcon.Exclamation, _
                     MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        '在庫調整Function
        Result = Tanaoroshi(Tyousei_Check, 4, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("在庫調整が完了しました。")

        zkensaku.zSearchFLg = True

        Me.Hide()
        zkensaku.Show()

    End Sub

    Private Sub ztyousei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Disp_Title As String = "在庫調整"

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

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If Form_Load = True Then
            Dim Row As Integer = DataGridView1.CurrentCell.RowIndex
            '数量から棚卸数量をひいて、差分に結果を表示する。
            'もし在庫数量欄にマイナスを入れたらメッセージを表示する。
            Dim NumChkErrorMessage As String = Nothing
            Dim NumChkResult As Boolean = True

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

        End If
    End Sub
End Class