Public Class zdisman

    Public I_Id As String

    Dim Form_Load As Boolean = False

    Public Structure Set_Item_Check_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim NUM As Integer
        Dim BEFORE_NUM As Integer
        Dim I_STATUS As Integer
    End Structure

    Private Sub zdisman_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub zdisman_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Disp_Title As String = "セットばらし"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label6.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemData() As Item_List = Nothing

        Dim Location As String = Nothing
        Dim LocationErrorMessage As String = Nothing
        Dim LocationResult As Boolean = True

        'ウインドウを表示したとき、数量にフォーカスを移動。
        TextBox3.Focus()

        Form_Load = True

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        Dim NumChkResult As Boolean = True
        Dim NumChkErrorMessage As String = Nothing

        Dim ErrorFlg As Boolean = True

        Dim DismantlingList() As Dismantling_List = Nothing

        Dim StringChkErrorMessage As String = Nothing
        Dim StringChkResult As Boolean = True

        Dim Set_I_ID As Integer = 0
        Dim Set_I_STATUS As Integer = 0
        Dim Set_NUM As Integer = 0
        Dim Set_UPD_NUM As Integer = 0
        Dim Set_STOCK_ID As Integer = 0

        Dim ChkLocation As String = Nothing
        Dim DataGridErrorMessage As String = Nothing

        '数量が未入力、０、マイナス値、整数値でないかを確認。
        '整数値、未入力NG、0の値NG、マイナス値NG
        If NumChkVal(Trim(TextBox3.Text), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
            MsgBox(NumChkErrorMessage)
            TextBox3.BackColor = Color.Salmon
            Exit Sub
        Else
            'チェックに問題がなければ背景色を白に戻す。
            TextBox3.BackColor = Color.White
        End If

        'もしセット数量より、大きい値を入力していたらメッセージを表示。
        If Trim(TextBox3.Text) > DataGridView1.Rows(0).Cells(3).Value Then
            MsgBox("在庫数量より、大きい値は入力できません。")
            TextBox3.BackColor = Color.Salmon
            Exit Sub
        Else
            TextBox3.BackColor = Color.White
        End If

        'DataGridView3のデータ入力チェックを行う。
        For Count = 0 To DataGridView3.Rows.Count - 2
            '商品コードが空なら
            If DataGridView3.Rows(Count).Cells(0).Value() = "" Then
                DataGridView3(0, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の商品コードを入力してください。" & vbCr
                ErrorFlg = False
            Else
                DataGridView3(0, Count).Style.BackColor = Color.White
            End If


            '数量が空なら
            If DataGridView3.Rows(Count).Cells(3).Value() = "" Then
                DataGridView3(3, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の数量を入力してください。" & vbCr
                ErrorFlg = False
            Else

                '数値の妥当性チェック
                If NumChkVal(Trim(DataGridView3.Rows(Count).Cells(3).Value()), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView3(4, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の数量は" & NumChkErrorMessage & vbCr

                    ErrorFlg = False
                Else
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView3(4, Count).Style.BackColor = Color.White
                End If

            End If

            '不良区分が空なら
            If DataGridView3.Rows(Count).Cells(4).Value() = "" Then
                DataGridView3(4, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目の不良区分を入力してください。" & vbCr
                ErrorFlg = False
            Else
                DataGridView3(4, Count).Style.BackColor = Color.White
            End If

            'ロケーションが空なら
            If DataGridView3.Rows(Count).Cells(5).Value() = "" Then
                DataGridView3(5, Count).Style.BackColor = Color.Salmon
                DataGridErrorMessage &= Count + 1 & "行目のロケーションを入力してください。" & vbCr
                ErrorFlg = False
            Else
                ChkLocation = DataGridView3.Rows(Count).Cells(5).Value()

                'ロケーションが101文字以上ならエラーメッセージを表示
                If ChkLocation.Length > 101 Then
                    DataGridErrorMessage &= Count + 1 & "行目のロケーションの文字数が長すぎます。" & vbCr

                    ErrorFlg = False
                End If

                DataGridView3(5, Count).Style.BackColor = Color.White
                End If

        Next

        If ErrorFlg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub

        End If

        'DataGridView3の値を配列に格納する。
        For Count = 0 To DataGridView3.Rows.Count - 2
            ReDim Preserve DismantlingList(0 To Count)

            '商品ID
            DismantlingList(Count).I_ID = DataGridView3.Rows(Count).Cells(7).Value()
            '商品コード
            DismantlingList(Count).I_CODE = DataGridView3.Rows(Count).Cells(0).Value()
            '商品名
            DismantlingList(Count).I_NAME = DataGridView3.Rows(Count).Cells(1).Value()
            '数量
            DismantlingList(Count).NUM = DataGridView3.Rows(Count).Cells(3).Value()
            '不良区分
            DismantlingList(Count).I_STATUS = DataGridView3.Rows(Count).Cells(4).Value()
            'ロケーション
            DismantlingList(Count).LOCATION = DataGridView3.Rows(Count).Cells(5).Value().Replace("'", "''")
            '倉庫
            DismantlingList(Count).PLACE = DataGridView3.Rows(Count).Cells(6).Value()
            '倉庫ID
            DismantlingList(Count).PLACE_ID = DataGridView3.Rows(Count).Cells(7).Value()

        Next

        'セット商品の情報を格納
        Set_I_ID = DataGridView1.Rows(0).Cells(8).Value()
        '不良区分
        If DataGridView1.Rows(0).Cells(6).Value() = "良品" Then
            Set_I_STATUS = 1
        ElseIf DataGridView1.Rows(0).Cells(6).Value() = "不良品" Then
            Set_I_STATUS = 2
        End If
        '現在数量
        Set_NUM = DataGridView1.Rows(0).Cells(3).Value()
        '更新後数量
        Set_UPD_NUM = TextBox3.Text
        'STOCK_ID
        Set_STOCK_ID = DataGridView1.Rows(0).Cells(8).Value()


        'ばらすセット商品情報を表示
        zdisman_conf.Label2.Text &= "商品名：" & Label11.Text & "  数量：" & Set_NUM & " → 数量：" & (Set_NUM - Set_UPD_NUM) & vbCr
        zdisman_conf.Label2.Text &= "を以下の商品にばらします。" & vbCr & vbCr

        'zdisman_confのLabel2に組み替え情報を表示
        For i = 0 To DismantlingList.Length - 1
            zdisman_conf.Label2.Text &= "商品名：" & DismantlingList(i).I_NAME & "　、数量：　" & DismantlingList(i).NUM & "　×　" & Set_UPD_NUM & "セット（数量合計：" & DismantlingList(i).NUM * Set_UPD_NUM & "）" & vbCr
        Next


        Me.Hide()
        zdisman_conf.Show()

    End Sub

    Private Sub DataGridView3_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellValueChanged

        Dim GetItemResult As Boolean = True
        Dim GetItemErrorMessage As String = Nothing

        Dim Item_Data() As Item_List = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim ChkStringAfter As String = Nothing
        Dim ChkMemo As String = Nothing

        If Form_Load = True Then

            Dim Row As Integer = DataGridView3.CurrentCell.RowIndex
            Dim Column As Integer = DataGridView3.CurrentCell.ColumnIndex

            'カラムが0なら商品コード
            If Column = 0 Then
                '入力された商品コードで商品名とJANコード、基本ロケーションを取得する。
                GetItemResult = GetItemName(Trim(DataGridView3(Column, Row).Value), 1, Item_Data, GetItemResult, GetItemErrorMessage)
                If GetItemResult = False Then
                    MsgBox(GetItemErrorMessage)
                    DataGridView3(0, Row).Style.BackColor = Color.Salmon
                    Exit Sub
                End If

                '取得できたら、商品名、JANコード、基本ロケーション、商品IDをセットする。
                DataGridView3(1, Row).Value = Item_Data(0).I_NAME
                DataGridView3(2, Row).Value = Item_Data(0).JAN
                DataGridView3(5, Row).Value = Item_Data(0).LOCATION
                DataGridView3(7, Row).Value = Item_Data(0).ID
                DataGridView3(0, Row).Style.BackColor = Color.White
                '倉庫情報をセット
                DataGridView3(6, Row).Value = DataGridView1(7, 0).Value
                DataGridView3(8, Row).Value = DataGridView1(10, 0).Value

            ElseIf Column = 3 Then
                ' 整数値、未入力NG、0の値NG、マイナス値NG
                If NumChkVal(Trim(DataGridView3(Column, Row).Value), "INTEGER", False, False, False, NumChkResult, NumChkErrorMessage) = False Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView3(3, Row).Style.BackColor = Color.Salmon
                    MsgBox(NumChkErrorMessage)
                Else
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView3(3, Row).Style.BackColor = Color.White
                End If
            ElseIf Column = 4 Then
                ' 良品区分が空白なら背景色を赤にする。
                If DataGridView3(Column, Row).Value = "" Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView3(Column, Row).Style.BackColor = Color.Salmon
                Else
                    'チェックに問題がなければ背景色を白に戻す。
                    DataGridView3(Column, Row).Style.BackColor = Color.White
                End If
            ElseIf Column = 5 Then
                'ロケーション
                '文字数の長さをチェック。101文字以上ならエラーメッセージ表示
                ChkMemo = DataGridView3(5, Row).Value

                If ChkMemo.Length > 100 Then
                    DataGridView3(5, Row).Style.BackColor = Color.Salmon
                    MsgBox("入力した文字が長すぎます。")
                    Exit Sub

                End If
            End If
        End If
    End Sub

End Class