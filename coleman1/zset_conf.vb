Imports System.Windows.Forms

Public Class zset_conf

    Private Sub zset_conf_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        zset.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '登録ボタン
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim PLACE_ID As Integer = Nothing
        Dim SetItemList() As Set_Item_List = Nothing

            'セット組み換え画面のDataGridView2の値を配列に格納
            'DataGridView2の値を配列に格納
        For Count = 0 To zset.DataGridView1.Rows.Count - 2
            ReDim Preserve SetItemList(0 To Count)
            'STOCK_ID
            SetItemList(Count).STOCK_ID = zset.DataGridView1.Rows(Count).Cells(8).Value()
            '商品ID
            SetItemList(Count).I_ID = zset.DataGridView1.Rows(Count).Cells(9).Value()
            '商品コード
            SetItemList(Count).I_CODE = zset.DataGridView1.Rows(Count).Cells(0).Value()
            '商品名
            SetItemList(Count).I_NAME = zset.DataGridView1.Rows(Count).Cells(1).Value()
            '変更前数量
            SetItemList(Count).BEFORE_NUM = zset.DataGridView1.Rows(Count).Cells(3).Value()
            '1セットあたりの数量
            SetItemList(Count).SET_NUM = zset.DataGridView1.Rows(Count).Cells(3).Value()
            '組み換え後数量
            SetItemList(Count).STOCK_NUM = zset.DataGridView1.Rows(Count).Cells(3).Value() - (zset.DataGridView1.Rows(Count).Cells(4).Value() * Trim(zset.TextBox1.Text))
            '不良区分
            If zset.DataGridView1.Rows(Count).Cells(7).Value() = "良品" Then
                SetItemList(Count).I_STATUS = 1
            ElseIf zset.DataGridView1.Rows(Count).Cells(7).Value() = "不良品" Then
                SetItemList(Count).I_STATUS = 2
            End If
            '倉庫
            'If zset.RadioButton1.Checked = True Then
            '    SetItemList(Count).PLACE = 1
            '    Place = 1
            'ElseIf zset.RadioButton2.Checked = True Then
            '    SetItemList(Count).PLACE = 2
            '    Place = 2
            'End If
            SetItemList(Count).PLACE = Label3.Text
        Next

            'STOCKのデータを修正し、STOCK_LOGに履歴を書き込む
        Result = Upd_Set(SetItemList, zset.ComboBox1.SelectedValue.ToString, Trim(zset.TextBox1.Text), Trim(zset.TextBox2.Text), Label3.Text, Result, ErrorMessage)
            If Result = False Then
                MsgBox(ErrorMessage)
                Exit Sub
            End If

            MsgBox("通常商品からセット商品への組み換え完了しました。")

        zkensaku.zSearchFLg = True


        Me.Dispose()
        zset.Dispose()
        topmenu.Show()

    End Sub

    Private Sub zset_conf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Disp_Title As String = "セット組確認"

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
End Class
