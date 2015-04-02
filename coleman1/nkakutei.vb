Public Class nkakutei

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
        nkensaku.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim InDefinition_Check() As InDefinition_List = Nothing
        Dim InDefinition_Data_Count As Integer = 0
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Status As Integer

        Dim Error_Flg As Boolean = True

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True
        Dim StringChkErrorMessage As String = Nothing
        Dim StringChkResult As Boolean = True

        Dim ChkStringAfter As String = Nothing
        Dim ChkDateAfter As String = Nothing

        Dim DataGridErrorMessage As String = Nothing

        'DataGridViewのデータを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            ReDim Preserve InDefinition_Check(0 To InDefinition_Data_Count)
            '商品コード
            InDefinition_Check(InDefinition_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()

            '入庫数の入力妥当性チェック
            '***チェック項目***
            ' 整数値、未入力NG、0の値NG、マイナス値NG
            If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(4).Value()), "INTEGER", False, True, False, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の数量の値が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(4, Count).Style.BackColor = Color.White
                '入荷数量
                InDefinition_Check(InDefinition_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(4).Value()
            End If

            ChkDateAfter = Nothing

            '入庫日の妥当性チェック
            If DateChkVal(Trim(DataGridView1.Rows(Count).Cells(5).Value()), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(5, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の入庫日が正しくありません。" & vbCr

                Error_Flg = False
            Else
                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(5, Count).Style.BackColor = Color.White
                '入庫日
                InDefinition_Check(InDefinition_Data_Count).FIX_DATE = ChkDateAfter

                'DataGridViewにも反映
                DataGridView1.Rows(Count).Cells(5).Value = ChkDateAfter
            End If

            ChkStringAfter = Nothing

            'ロケーション（文字列）の妥当性チェック
            If StringChkVal(Trim(DataGridView1.Rows(Count).Cells(6).Value()), False, False, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataGridView1(6, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目のロケーションが正しくありません。" & vbCr
                Error_Flg = False
            Else
                If ChkStringAfter.Length > 100 Then
                    DataGridView1(6, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目のロケーションの文字数が長すぎます。" & vbCr
                    Error_Flg = False
                End If

                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(6, Count).Style.BackColor = Color.White


                'ロケーション
                InDefinition_Check(InDefinition_Data_Count).LOCATION = ChkStringAfter
            End If

            '入庫コメントを追加 2012/2/27
            If StringChkVal(Trim(DataGridView1.Rows(Count).Cells(7).Value()), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
                DataGridView1(7, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の入庫コメントが正しくありません。" & vbCr
                Error_Flg = False
            Else
                If ChkStringAfter.Length > 500 Then
                    DataGridView1(7, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の入庫コメントの文字数が長すぎます。" & vbCr
                    Error_Flg = False
                End If

                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(7, Count).Style.BackColor = Color.White

                'コメント
                InDefinition_Check(InDefinition_Data_Count).IN_COMMENT = ChkStringAfter
            End If

            'コメントを在庫コメントに修正 2012/2/27
            'コメント（文字列）の妥当性チェック
            If StringChkVal(Trim(DataGridView1.Rows(Count).Cells(8).Value()), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
                DataGridView1(8, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の在庫コメントが正しくありません。" & vbCr
                Error_Flg = False
            Else
                If ChkStringAfter.Length > 500 Then
                    DataGridView1(8, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の在庫コメントの文字数が長すぎます。" & vbCr
                    Error_Flg = False
                End If

                'チェックに問題がなければ背景色を白に戻す。
                DataGridView1(8, Count).Style.BackColor = Color.White

                'コメント
                InDefinition_Check(InDefinition_Data_Count).STOCK_COMMENT = ChkStringAfter
            End If

            'If StringChkVal(Trim(DataGridView1.Rows(Count).Cells(7).Value()), True, True, ChkStringAfter, StringChkResult, StringChkErrorMessage) = False Then
            '    DataGridView1(7, Count).Style.BackColor = Color.Salmon
            '    DataGridErrorMessage &= Count + 1 & "行目のコメントが正しくありません。" & vbCr
            '    Error_Flg = False
            'Else
            '    If ChkStringAfter.Length > 500 Then
            '        DataGridView1(7, Count).Style.BackColor = Color.Salmon
            '        DataGridErrorMessage &= Count + 1 & "行目のコメントの文字数が長すぎます。" & vbCr
            '        Error_Flg = False
            '    End If

            '    'チェックに問題がなければ背景色を白に戻す。
            '    DataGridView1(7, Count).Style.BackColor = Color.White

            '    'コメント
            '    InDefinition_Check(InDefinition_Data_Count).COMMENT = ChkStringAfter
            'End If

            '不良区分
            If DataGridView1.Rows(Count).Cells(13).Value() = "良品" Then
                Status = 1
            Else
                Status = 2
            End If

            '倉庫
            'If DataGridView1.Rows(Count).Cells(14).Value() = "平和島" Then
            '    InDefinition_Check(InDefinition_Data_Count).PLACE = 1
            'ElseIf DataGridView1.Rows(Count).Cells(14).Value() = "返品センター" Then
            '    InDefinition_Check(InDefinition_Data_Count).PLACE = 2
            'ElseIf DataGridView1.Rows(Count).Cells(14).Value() = "営業倉庫" Then
            '    InDefinition_Check(InDefinition_Data_Count).PLACE = 3
            'ElseIf DataGridView1.Rows(Count).Cells(14).Value() = "八潮" Then
            'End If
            InDefinition_Check(InDefinition_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(18).Value()

            InDefinition_Check(InDefinition_Data_Count).I_STATUS = Status
            'ID
            InDefinition_Check(InDefinition_Data_Count).ID = DataGridView1.Rows(Count).Cells(15).Value()
            'DetailID
            InDefinition_Check(InDefinition_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(16).Value()
            'I_ID
            InDefinition_Check(InDefinition_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(17).Value()

            InDefinition_Data_Count += 1
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("入庫確定してもよろしいですか？", _
                             "確認", _
                             MessageBoxButtons.YesNo, _
                             MessageBoxIcon.Exclamation, _
                             MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If

        '入庫確定Function
        Result = InDefinition(InDefinition_Check, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '確定処理が終了したら、画面を閉じて入庫予定検索画面を表示。
        MsgBox("入庫確定が完了しました。")

        nkensaku.nSearchFLg = True
        Me.Dispose()
        nkensaku.Show()

    End Sub

    Private Sub nkakutei_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - 入庫確定"
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

        '左7項目を固定(ドキュメント№、商品コード、予定数量、入荷予定日、入荷数量、入荷日、ロケーション,入庫コメント,在庫コメント)
        DataGridView1.Columns(8).Frozen = True

    End Sub

    Private Sub nkakutei_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

        Dim Row As Integer = DataGridView1.CurrentCell.RowIndex

        '現在の行のロケーションが入れられてなかったらエラーメッセージ表示
        If DataGridView1.Rows(Row).Cells(6).Value() = "" Then
            MsgBox(Row + 1 & "行目のロケーションが未入力です。")
            Exit Sub
        End If

        '問題なければ、以降全てにコピー。

        For Count = Row To DataGridView1.Rows.Count - 1
            DataGridView1(6, Count).Value = DataGridView1(6, Row).Value
        Next

    End Sub
End Class