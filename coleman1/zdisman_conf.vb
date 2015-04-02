Imports System.Windows.Forms

Public Class zdisman_conf

    Private Sub zdisman_conf_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        'キャンセルボタン
        Me.Dispose()
        zdisman.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '登録ボタン
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim DismantlingList() As Dismantling_List = Nothing
        Dim DismantlingResult As Boolean = True
        Dim DismantlingErrorMessage As String = Nothing

        Dim Set_I_ID As Integer = 0
        Dim Set_I_STATUS As Integer = 0
        Dim Set_NUM As Integer = 0
        Dim Set_UPD_NUM As Integer = 0
        Dim Set_STOCK_ID As Integer = 0
        Dim Set_Place As String = Nothing

        'DataGridView3の値を配列に格納する。
        For Count = 0 To zdisman.DataGridView3.Rows.Count - 2
            ReDim Preserve DismantlingList(0 To Count)

            '商品ID
            DismantlingList(Count).I_ID = zdisman.DataGridView3.Rows(Count).Cells(7).Value()
            '商品コード
            DismantlingList(Count).I_CODE = zdisman.DataGridView3.Rows(Count).Cells(0).Value()
            '商品名
            DismantlingList(Count).I_NAME = zdisman.DataGridView3.Rows(Count).Cells(1).Value()
            '数量
            DismantlingList(Count).NUM = zdisman.DataGridView3.Rows(Count).Cells(3).Value()
            '不良区分
            DismantlingList(Count).I_STATUS = zdisman.DataGridView3.Rows(Count).Cells(4).Value()
            'ロケーション
            DismantlingList(Count).LOCATION = zdisman.DataGridView3.Rows(Count).Cells(5).Value().Replace("'", "''")
            '倉庫
            DismantlingList(Count).PLACE = zdisman.DataGridView3.Rows(Count).Cells(6).Value()
            '倉庫ID
            DismantlingList(Count).PLACE_ID = zdisman.DataGridView3.Rows(Count).Cells(8).Value()
        Next

        'セット商品の情報を格納
        Set_I_ID = zdisman.DataGridView1.Rows(0).Cells(8).Value()
        '不良区分
        If zdisman.DataGridView1.Rows(0).Cells(6).Value() = "良品" Then
            Set_I_STATUS = 1
        ElseIf zdisman.DataGridView1.Rows(0).Cells(6).Value() = "不良品" Then
            Set_I_STATUS = 2
        End If
        '現在数量
        Set_NUM = zdisman.DataGridView1.Rows(0).Cells(3).Value()
        '更新後数量
        Set_UPD_NUM = zdisman.TextBox3.Text
        'STOCK_ID
        Set_STOCK_ID = zdisman.DataGridView1.Rows(0).Cells(8).Value()
        '倉庫
        'Set_Place = zdisman.DataGridView1.Rows(0).Cells(7).Value()
        Set_Place = zdisman.DataGridView1.Rows(0).Cells(10).Value()

        'Function
        DismantlingResult = Upd_Dismantling(DismantlingList, Set_I_ID, Set_I_STATUS, Set_NUM, Set_UPD_NUM, Set_STOCK_ID, Set_Place, DismantlingResult, DismantlingErrorMessage)
        If DismantlingResult = False Then
            MsgBox(DismantlingErrorMessage)
            Exit Sub
        End If

        MsgBox("セット商品から通常商品へのばらしが完了しました。")

        zkensaku.zSearchFLg = True

        Me.Dispose()
        zdisman.Dispose()
        zkensaku.Show()

    End Sub

    Private Sub zdisman_conf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "セットばらし確認"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title

        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub
End Class
