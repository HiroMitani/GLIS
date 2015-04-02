Public Class syukokakutei

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        syukokensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '出荷確定ボタン、クリック
        Dim OutDefinition_Data() As OutDefinition_List = Nothing
        Dim Count As Integer = 0

        Dim OutDefinition_Result As Boolean = True
        Dim OutDefinition_Result_Message As String = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True
        Dim Error_Flg As Boolean = True

        Dim ChkDateAfter As String = Nothing

        Dim Stock_Message As String = Nothing
        Dim Stock_Minus_Check As Boolean = False

        Dim DataGridErrorMessage As String = Nothing


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve OutDefinition_Data(0 To Count)
            'ID
            OutDefinition_Data(Count).ID = DataGridView1.Rows(Count).Cells(21).Value()
            '商品ID
            OutDefinition_Data(Count).I_ID = DataGridView1.Rows(Count).Cells(22).Value()

            '倉庫IDを使用するため、コメントアウト2015/1/22 H.Mitani
            '倉庫
            'If DataGridView1.Rows(Count).Cells(20).Value() = "平和島" Then
            '    OutDefinition_Data(Count).PLACE = 1
            'ElseIf DataGridView1.Rows(Count).Cells(20).Value() = "返品センター" Then
            '    OutDefinition_Data(Count).PLACE = 2
            'ElseIf DataGridView1.Rows(Count).Cells(20).Value() = "営業倉庫" Then
            '    OutDefinition_Data(Count).PLACE = 3
            'ElseIf DataGridView1.Rows(Count).Cells(20).Value() = "八潮" Then
            'End If

            '倉庫
            OutDefinition_Data(Count).PLACE = DataGridView1.Rows(Count).Cells(23).Value()

            'ステータス（良品か不良品か）
            OutDefinition_Data(Count).I_STATUS = DataGridView1.Rows(Count).Cells(15).Value()
            If DataGridView1.Rows(Count).Cells(15).Value() = "良品" Then
                '良品なら
                OutDefinition_Data(Count).I_STATUS = 1
            ElseIf DataGridView1.Rows(Count).Cells(15).Value() = "不良品" Then
                '不良品なら
                OutDefinition_Data(Count).I_STATUS = 2
            ElseIf DataGridView1.Rows(Count).Cells(15).Value() = "保管品" Then
                '保管品なら
                OutDefinition_Data(Count).I_STATUS = 3
            End If


            '出荷数量の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(8).Value()), "INTEGER", False, True, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(8, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の出荷数量が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(8, Count).Style.BackColor = Color.White

                '出荷数量
                OutDefinition_Data(Count).FIX_NUM = Trim(DataGridView1.Rows(Count).Cells(8).Value())

            End If

            '出荷日の妥当性チェック
            If DateChkVal(Trim(DataGridView1.Rows(Count).Cells(9).Value()), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                DataGridView1(9, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の出荷日が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(5, Count).Style.BackColor = Color.White

                '出荷日
                OutDefinition_Data(Count).FIX_DATE = ChkDateAfter
                'DataGridViewにも反映
                DataGridView1.Rows(Count).Cells(9).Value() = ChkDateAfter
            End If
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        If Stock_Minus_Check = True Then

            Stock_Message &= "出庫すると在庫数がマイナスになりますが、出庫してもよろしいですか？"
            '取り込んだデータに重複がある場合、ダイアログを表示。
            Dim Stock_Message_Result As DialogResult = MessageBox.Show(Stock_Message, _
                                         "確認", _
                                         MessageBoxButtons.YesNo, _
                                         MessageBoxIcon.Exclamation, _
                                         MessageBoxDefaultButton.Button2)

            '何が選択されたか調べる 
            If Stock_Message_Result = DialogResult.No Then
                'Noを選択した場合、処理終了
                Exit Sub
            End If
        End If

        '出庫確定 Function
        OutDefinition_Result = OutDefinition(OutDefinition_Data, OutDefinition_Result, OutDefinition_Result_Message)

        If OutDefinition_Result = False Then
            MsgBox(OutDefinition_Result_Message)
            Exit Sub
        End If

        MsgBox("出荷確定が完了しました。")

        syukokensaku.sSearchFLg = True
        syukokensaku.Show()
        Me.Dispose()
    End Sub

    Private Sub syukokakutei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Disp_Title As String = "出荷確定"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label2.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '左3項目を固定(伝票番号、オーダー番号、商品コード)
        DataGridView1.Columns(2).Frozen = True

        '出庫確定画面が表示されたら
        '出庫数量にフォーカスを移動。
        DataGridView1.Visible = True
        DataGridView1.Select()
        DataGridView1.CurrentCell = DataGridView1(8, 0)

    End Sub

    Private Sub syukokakutei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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