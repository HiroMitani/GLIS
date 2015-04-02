Imports System.Windows.Forms
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class nkensaku

    Dim PLACE_ID As String

    '帳票ファイルの格納フォルダ指定
    Public PrtForm As String = System.Configuration.ConfigurationManager.AppSettings("PrtPath")

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Public Change_Data_list As Change_List = Nothing

    Dim AppPath As String

    Public FormLord As Boolean = False

    '入庫確定や変更など行ったら、再度検索を行わせる為のFlg
    Public nSearchFLg As Boolean = False

    Public Structure Change_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim DOC_NO As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim N_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim STATUS As String
        Dim CATEGORY As String
        Dim DEFECT_TYPE As String
        Dim LOCATION As String
        Dim I_ID As Integer
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Private Sub nkensaku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "入庫関連検索"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label12.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '初期表示時、ドキュメント№にフォーカス
        TextBox1.Focus()

        'DataGridViewのクリア
        DataGridView1.Rows.Clear()
        '検索条件のクリア
        'ドキュメント№
        TextBox1.Text = ""
        '入荷予定日From、Toのクリア
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        '入荷日のFrom、Toのクリア
        MaskedTextBox3.Clear()
        MaskedTextBox4.Clear()
        '商品コード
        TextBox6.Text = ""

        '左3項目を固定(チェック、ドキュメント№、商品コード)
        DataGridView1.Columns(2).Frozen = True

        '倉庫情報取得
        Dim PlaceTable As New DataTable()
        PlaceTable.Columns.Add("ID", GetType(String))
        PlaceTable.Columns.Add("NAME", GetType(String))
        '倉庫情報の取得
        Result = GetPLACEList(PlaceData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        For Count = 0 To PlaceData.Length - 1
            '新しい行を作成
            Dim row As DataRow = PlaceTable.NewRow()

            '各列に値をセット
            row("ID") = PlaceData(Count).ID
            row("NAME") = PlaceData(Count).NAME

            'DataTableに行を追加
            PlaceTable.Rows.Add(row)
        Next

        PlaceTable.AcceptChanges()
        'コンボボックスのDataSourceにDataTableを割り当てる
        ComboBox4.DataSource = PlaceTable

        '表示される値はDataTableのNAME列
        ComboBox4.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox4.ValueMember = "ID"
        '初期値をセット(倉庫データの1件目）
        ComboBox4.SelectedIndex = 0
        PLACE_ID = ComboBox4.SelectedValue.ToString()

        FormLord = True

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim InDefinition_Check_Flg As Boolean = False
        Dim InDefinition_Check_Message As String = Nothing
        Dim InDefinition_Data_Count As Integer = 0
        Dim InDefinition_Check() As Change_List = Nothing
        Dim Data_Num_Total As Integer = 0

        If nSearchFLg = True Then
            MsgBox("入荷確定、予定変更、入荷削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                InDefinition_Check_Flg = True
            End If
        Next
        If InDefinition_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        InDefinition_Check_Flg = True
        'チェックされた商品の中で入荷済みがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(9).Value() = "入荷済み" Then
                InDefinition_Check_Flg = False
            End If
        Next
        If InDefinition_Check_Flg = False Then
            MsgBox("入荷済みのデータがチェックされています。入荷済みのデータは入荷確定できません。")
            Exit Sub
        End If


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve InDefinition_Check(0 To InDefinition_Data_Count)
                'ドキュメント№
                InDefinition_Check(InDefinition_Data_Count).DOC_NO = DataGridView1.Rows(Count).Cells(1).Value()
                '商品コード
                InDefinition_Check(InDefinition_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(2).Value()
                '予定数量
                InDefinition_Check(InDefinition_Data_Count).NUM = DataGridView1.Rows(Count).Cells(5).Value()
                '入荷予定日
                InDefinition_Check(InDefinition_Data_Count).N_DATE = DataGridView1.Rows(Count).Cells(6).Value()
                '入荷数量（入庫確定に遷移するため、FIX_NUM(入庫数）は入庫予定数を入れて表示する。）
                InDefinition_Check(InDefinition_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(5).Value()
                '入荷日（入庫確定に遷移するため、FIX_DATE(入庫日）は今日の日付を入れて表示する。）
                InDefinition_Check(InDefinition_Data_Count).FIX_DATE = DateTime.Today
                '商品コード
                InDefinition_Check(InDefinition_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(9).Value()
                'JANコード
                InDefinition_Check(InDefinition_Data_Count).JAN = DataGridView1.Rows(Count).Cells(10).Value()
                'ステータス
                InDefinition_Check(InDefinition_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(11).Value()
                'カテゴリー
                InDefinition_Check(InDefinition_Data_Count).CATEGORY = DataGridView1.Rows(Count).Cells(12).Value()
                '不良区分
                InDefinition_Check(InDefinition_Data_Count).DEFECT_TYPE = DataGridView1.Rows(Count).Cells(13).Value()
                'ロケーション
                'InDefinition_Check(InDefinition_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(14).Value()
                'IN.ID
                InDefinition_Check(InDefinition_Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                'IN_DETAIL_DeiltaID
                InDefinition_Check(InDefinition_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(18).Value()
                'I_ID
                InDefinition_Check(InDefinition_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(19).Value()

                '倉庫
                InDefinition_Check(InDefinition_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(16).Value()
                'P_ID
                InDefinition_Check(InDefinition_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(20).Value()

                '配列に格納される予定数量を足していき、総数を表示する。
                Data_Num_Total += DataGridView1.Rows(Count).Cells(5).Value()
                InDefinition_Data_Count += 1
            End If
        Next

        'InDefinition_Data_Countが確定画面に表示する件数になるので、DataGridViewの上に表示する商品数はInDefinition_Data_Countを表示
        nkakutei.Label1.Text = "商品数： " & InDefinition_Data_Count
        '総数を表示
        nkakutei.Label3.Text = "総数： " & Data_Num_Total

        'nkakuteiのDataGridViewをクリア
        nkakutei.DataGridView1.Rows.Clear()

        '入荷確定の画面にチェックしたデータを表示する。
        For Count = 0 To InDefinition_Check.Length - 1
            Dim InDefinition_Data_List As New DataGridViewRow
            InDefinition_Data_List.CreateCells(nkakutei.DataGridView1)
            With InDefinition_Data_List
                'ドキュメント№
                .Cells(0).Value = InDefinition_Check(Count).DOC_NO
                '商品コード
                .Cells(1).Value = InDefinition_Check(Count).I_CODE
                '予定数量
                .Cells(2).Value = InDefinition_Check(Count).NUM
                '入荷予定日
                .Cells(3).Value = InDefinition_Check(Count).N_DATE
                '入荷数量
                .Cells(4).Value = InDefinition_Check(Count).FIX_NUM
                '入荷日
                .Cells(5).Value = InDefinition_Check(Count).FIX_DATE
                'ロケーション
                '.Cells(6).Value = InDefinition_Check(Count).LOCATION
                '商品名
                .Cells(9).Value = InDefinition_Check(Count).I_NAME
                'JANコード
                .Cells(10).Value = InDefinition_Check(Count).JAN
                'ステータス
                .Cells(11).Value = InDefinition_Check(Count).STATUS
                'カテゴリー
                .Cells(12).Value = InDefinition_Check(Count).CATEGORY
                '不良区分
                .Cells(13).Value = InDefinition_Check(Count).DEFECT_TYPE
                '倉庫
                .Cells(14).Value = InDefinition_Check(Count).PLACE

                'IN.ID
                .Cells(15).Value = InDefinition_Check(Count).ID
                'IN_DETAIL.DETAIL_ID
                .Cells(16).Value = InDefinition_Check(Count).DETAIL_ID
                'I_ID
                .Cells(17).Value = InDefinition_Check(Count).I_ID
                'P_ID
                .Cells(18).Value = InDefinition_Check(Count).PLACE_ID
                ''商品名
                '.Cells(8).Value = InDefinition_Check(Count).I_NAME
                ''JANコード
                '.Cells(9).Value = InDefinition_Check(Count).JAN
                ''ステータス
                '.Cells(10).Value = InDefinition_Check(Count).STATUS
                ''カテゴリー
                '.Cells(11).Value = InDefinition_Check(Count).CATEGORY
                ''不良区分
                '.Cells(12).Value = InDefinition_Check(Count).DEFECT_TYPE
                ''IN.ID
                '.Cells(13).Value = InDefinition_Check(Count).ID
                ''IN_DETAIL.DETAIL_ID
                '.Cells(14).Value = InDefinition_Check(Count).DETAIL_ID
                ''I__ID
                '.Cells(15).Value = InDefinition_Check(Count).I_ID

            End With
            nkakutei.DataGridView1.Rows.Add(InDefinition_Data_List)
        Next

        nkakutei.Show()
        Me.Hide()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Del_Check() As ItemID_List = Nothing
        Dim Count As Integer
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Del_Data_Count As Integer = 0
        Dim SearchResult() As Search_List = Nothing
        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Fix_Date_From As String = Nothing
        Dim Fix_Date_To As String = Nothing
        Dim Delete_Check_Flg As Boolean = False
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        If nSearchFLg = True Then
            MsgBox("入荷確定、予定変更、入荷削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Delete_Check_Flg = True
            End If
        Next
        If Delete_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Delete_Check_Flg = True
        'チェックされた商品の中で入荷済みがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(9).Value() = "入荷済み" Then
                Delete_Check_Flg = False
            End If
        Next
        If Delete_Check_Flg = False Then
            MsgBox("入荷済みのデータがチェックされています。入荷済みのデータは削除できません。")
            Exit Sub
        End If

        Dim Message_Result As DialogResult = MessageBox.Show("チェックされた商品を削除してもよろしいですか？", _
                                             "確認", _
                                             MessageBoxButtons.YesNo, _
                                             MessageBoxIcon.Exclamation, _
                                             MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If Message_Result = DialogResult.No Then
            'Noを選択した場合、処理終了
            Exit Sub
        End If
        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Del_Check(0 To Del_Data_Count)
                ''IDを格納
                'Del_Check(Del_Data_Count).ID = DataGridView1.Rows(Count).Cells(13).Value()
                ''Detail_IDを格納
                'Del_Check(Del_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(14).Value()

                'IDを格納
                Del_Check(Del_Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                'Detail_IDを格納
                Del_Check(Del_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(18).Value()
                Del_Data_Count += 1
            End If
        Next

        '削除用Function
        Result = Del_Item(Del_Check, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
        End If

        MsgBox("削除が完了しました。")
        nSearchFLg = True

    End Sub
    Private Sub nkensaku_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Public Class Win32
        Declare Auto Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    End Class

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount As Integer = 20

        'For i = 0 To loopcnt - 1
        '    If DataGridView1(0, i).Value = 1 Then
        '        DataGridView1(0, i).Value = 0
        '        For Count = 0 To IndexCount
        '            DataGridView1.Item(Count, i).Style.BackColor = Color.White
        '        Next
        '    ElseIf DataGridView1(0, i).Value = 0 Then
        '        DataGridView1(0, i).Value = 1
        '        For Count = 0 To IndexCount
        '            DataGridView1.Item(Count, i).Style.BackColor = Color.PaleGreen
        '        Next
        '    End If
        'Next

        Dim CheckDataCount As Integer = 0

        '2012/02/10 菊池様の依頼により、DataGridViewをドラッグで複数行指定し、
        'チェックボタンを押すことにより、指定された行のチェックをつけたりはずしたりする機能を実装

        Dim CheckData() As Integer = Nothing
        Dim SelectCount As Integer = 0
        Dim LoopCount As Integer = 0
        For Each Row As DataGridViewRow In DataGridView1.SelectedRows
            ReDim Preserve CheckData(0 To SelectCount)
            CheckData(SelectCount) = Row.Index
            SelectCount += 1
        Next Row

        If SelectCount = 0 Then
            For i = 0 To loopcnt - 1
                If DataGridView1(0, i).Value = 1 Then
                    DataGridView1(0, i).Value = 0
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.White
                    Next
                ElseIf DataGridView1(0, i).Value = 0 Then
                    DataGridView1(0, i).Value = 1
                    For Count = 0 To IndexCount
                        DataGridView1.Item(Count, i).Style.BackColor = Color.PaleGreen
                    Next
                End If
            Next
        Else
            For LoopCount = 0 To CheckData.Length - 1
                If DataGridView1(0, CheckData(LoopCount)).Value = 1 Then
                    If CheckBox7.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.White
                        Next
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox7.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = 1
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.PaleGreen
                        Next
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Update_Check_Flg As Boolean = False
        Dim Upd_Data_Count As Integer = 0
        Dim Upd_Check() As Change_List = Nothing

        If nSearchFLg = True Then
            MsgBox("入荷確定、予定変更、入荷削除を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Update_Check_Flg = True
            End If
        Next

        If Update_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Update_Check_Flg = True
        'チェックされた商品の中で入荷済みがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(9).Value() = "入荷済み" Then
                Update_Check_Flg = False
            End If
        Next
        If Update_Check_Flg = False Then
            MsgBox("入荷済みのデータがチェックされています。入荷済みのデータは予定変更できません。")
            Exit Sub
        End If

        '検索条件のドキュメント№と違うものがあれば、メッセージを表示。
        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox1.Text) <> DataGridView1.Rows(Count).Cells(1).Value() Then
                MsgBox("入荷予定変更を行うには" & vbCr & "ドキュメント№での検索が必須です。")
                Exit Sub
            End If
        Next

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Upd_Check(0 To Upd_Data_Count)
                'ドキュメント№
                Upd_Check(Upd_Data_Count).DOC_NO = DataGridView1.Rows(Count).Cells(1).Value()
                '商品コード
                Upd_Check(Upd_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(2).Value()
                '予定数量
                Upd_Check(Upd_Data_Count).NUM = DataGridView1.Rows(Count).Cells(5).Value()
                '入荷予定日
                Upd_Check(Upd_Data_Count).N_DATE = DataGridView1.Rows(Count).Cells(6).Value()
                '入荷数量
                '入荷数量はNULLなら、0に置き換える。
                If DataGridView1.Rows(Count).Cells(7).Value() = "" Then
                    Upd_Check(Upd_Data_Count).FIX_NUM = 0
                Else
                    Upd_Check(Upd_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(7).Value()
                End If
                '入荷日
                Upd_Check(Upd_Data_Count).FIX_DATE = DataGridView1.Rows(Count).Cells(8).Value()
                '商品名
                Upd_Check(Upd_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(9).Value()
                'JANコード
                Upd_Check(Upd_Data_Count).JAN = DataGridView1.Rows(Count).Cells(10).Value()
                'ステータス
                Upd_Check(Upd_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(11).Value()
                'カテゴリー
                Upd_Check(Upd_Data_Count).CATEGORY = DataGridView1.Rows(Count).Cells(12).Value()
                '不良区分
                Upd_Check(Upd_Data_Count).DEFECT_TYPE = DataGridView1.Rows(Count).Cells(13).Value()
                'ID
                Upd_Check(Upd_Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                'DETAIL_ID
                Upd_Check(Upd_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(18).Value()
                'I_ID
                Upd_Check(Upd_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(19).Value()

                Upd_Data_Count += 1
            End If
        Next

        'nhenkouのDataGridViewをクリア
        nhenkou.DataGridView1.Rows.Clear()

        '変更画面のComboBoxのリストに表示する項目を指定する
        Dim Category_Column As New DataGridViewComboBoxColumn()
        Category_Column.Items.Add("通常入荷")
        Category_Column.Items.Add("返品入荷")
        'データを表示する
        Category_Column.DataPropertyName = "種別"
        'ComboBox列を表示する
        nhenkou.DataGridView1.Columns.Insert(nhenkou.DataGridView1.Columns("Column10").Index, Category_Column)
        nhenkou.DataGridView1.Columns.Remove("Column10")
        Category_Column.Name = "種別"

        Dim Defect_Type_Column As New DataGridViewComboBoxColumn()
        Defect_Type_Column.Items.Add("良品")
        Defect_Type_Column.Items.Add("不良品")
        'バインドされているデータを表示する
        Defect_Type_Column.DataPropertyName = "不良区分"
        'ComboBox列を表示する
        nhenkou.DataGridView1.Columns.Insert(nhenkou.DataGridView1.Columns("Column11").Index, Defect_Type_Column)
        nhenkou.DataGridView1.Columns.Remove("Column11")
        Defect_Type_Column.Name = "不良区分"

        '2012/02/10 使用追加 変更画面でドキュメント№の修正が行えるよう対応。 
        '1行目のデータのドキュメント№を変更画面にも表示する。
        nhenkou.TextBox1.Text = DataGridView1.Rows(0).Cells(1).Value()

        For Count = 0 To Upd_Check.Length - 1
            Dim Change_Data_list As New DataGridViewRow
            Change_Data_list.CreateCells(nhenkou.DataGridView1)
            With Change_Data_list
                'ドキュメント№
                .Cells(0).Value = Upd_Check(Count).DOC_NO
                '商品コード
                .Cells(1).Value = Upd_Check(Count).I_CODE
                '予定数量
                .Cells(2).Value = Upd_Check(Count).NUM
                '入荷予定日
                .Cells(3).Value = Upd_Check(Count).N_DATE
                '入荷数量
                If Upd_Check(Count).FIX_NUM = 0 Then
                    .Cells(4).Value = ""
                Else
                    .Cells(4).Value = Upd_Check(Count).FIX_NUM
                End If
                '入荷日
                .Cells(5).Value = Upd_Check(Count).FIX_DATE
                '商品名
                .Cells(6).Value = Upd_Check(Count).I_NAME
                'JANコード
                .Cells(7).Value = Upd_Check(Count).JAN
                'ステータス
                .Cells(8).Value = Upd_Check(Count).STATUS
                '種別
                .Cells(9).Value = Upd_Check(Count).CATEGORY
                '不良区分
                .Cells(10).Value = Upd_Check(Count).DEFECT_TYPE
                'ロケーション
                '.Cells(11).Value = Upd_Check(Count).LOCATION
                'ID
                .Cells(11).Value = Upd_Check(Count).ID
                'DETAIL_ID
                .Cells(12).Value = Upd_Check(Count).DETAIL_ID
                'I_ID
                .Cells(13).Value = Upd_Check(Count).I_ID
            End With
            nhenkou.DataGridView1.Rows.Add(Change_Data_list)
        Next

        nhenkou.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemName() As Item_List = Nothing
        Dim ChkItemCodeString As String = Nothing

        '商品名欄をクリアする。
        Label7.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '入力チェック
            If Trim(TextBox6.Text) <> "" Then
                'あたかもTabキーが押されたかのようにする
                'Shiftが押されている時は前のコントロールのフォーカスを移動
                Me.ProcessTabKey(Not e.Shift)
                e.Handled = True
                '入力された商品コードを元に商品名を取得する。

                '商品コードに'が入力されていたらReplaceする。
                ChkItemCodeString = Trim(TextBox6.Text)
                ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

                '商品名取得Function
                Result = GetItemName(ChkItemCodeString, 1, ItemName, Result, ErrorMessage)
                If Result = "True" Then
                    Label7.Text = ItemName(0).I_NAME
                    CheckBox1.Focus()
                    TextBox6.BackColor = Color.White
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox6.Focus()
                    TextBox6.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    Label7.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MaskedTextBox1.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MaskedTextBox2.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True

            MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MaskedTextBox4.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True
            TextBox6.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MaskedTextBox3.KeyDown
        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            'あたかもTabキーが押されたかのようにする
            'Shiftが押されている時は前のコントロールのフォーカスを移動
            Me.ProcessTabKey(Not e.Shift)
            e.Handled = True
            MaskedTextBox4.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As Search_List = Nothing
        Dim Count As Integer = 0
        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Fix_Date_From As String = Nothing
        Dim Fix_Date_To As String = Nothing
        Dim Date_Check_Result As DateTime
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0
        Dim ChkItemCodeString As String = Nothing
        Dim ChkDocNoString As String = Nothing

        '一括チェックのチェックボックスをクリアする。
        CheckBox7.Checked = False

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        '検索ボタンクリックチェック
        nSearchFLg = False

        '入荷予定日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            Date_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '入荷予定日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '入荷日Fromのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            Fix_Date_From = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Fix_Date_From = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("入荷日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '入荷日Toのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            Fix_Date_To = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Fix_Date_To = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        ChkDocNoString = Trim(TextBox1.Text)
        ChkItemCodeString = Trim(TextBox6.Text)

        'ドキュメント№に'が入力されたいたらReplaceする。
        ChkDocNoString = ChkDocNoString.Replace("'", "''")

        '商品コードに'が入力されていたらReplaceする。
        ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        '検索Function
        Result = GetInSeach(ChkDocNoString, Date_From, Date_To, Fix_Date_From, Fix_Date_To, _
                            ChkItemCodeString, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, _
                            CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, PLACE_ID, SearchResult, Data_Total, _
                            Data_Num_Total, Result, ErrorMessage)
        If Result = False Then
            '商品数、総数をクリア
            Label11.Text = "商品数："
            Label13.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If
        '結果を元にDataGridViewに表示する。

        '商品数を表示
        Label11.Text = "商品数： " & Data_Total

        '総数を表示
        Label13.Text = "総数： " & Data_Num_Total
        If Result = True Then
            'DataGridへ入力したデータを挿入
            For Count = 0 To SearchResult.Length - 1
                Dim SR_List As New DataGridViewRow
                SR_List.CreateCells(DataGridView1)
                With SR_List
                    'ドキュメント№
                    .Cells(1).Value = SearchResult(Count).DOC_NO
                    '商品コード
                    .Cells(2).Value = SearchResult(Count).I_CODE
                    '入庫元コード
                    .Cells(3).Value = SearchResult(Count).C_CODE
                    '入庫元名
                    .Cells(4).Value = SearchResult(Count).C_NAME
                    '予定数量
                    .Cells(5).Value = SearchResult(Count).NUM
                    '入荷予定日
                    .Cells(6).Value = SearchResult(Count).N_DATE
                    '入荷数量
                    'ステータスが入荷済みのものは入荷数を表示し、入荷予定のものは空白表示
                    If SearchResult(Count).STATUS = "入荷済み" Then
                        .Cells(7).Value = SearchResult(Count).FIX_NUM
                    Else
                        .Cells(7).Value = ""
                    End If
                    '入荷日
                    .Cells(8).Value = SearchResult(Count).FIX_DATE
                    '商品名
                    .Cells(9).Value = SearchResult(Count).I_NAME
                    'JANコード
                    .Cells(10).Value = SearchResult(Count).JAN_CODE
                    'ステータス
                    .Cells(11).Value = SearchResult(Count).STATUS
                    'カテゴリー
                    .Cells(12).Value = SearchResult(Count).CATEGORY
                    '不良区分
                    .Cells(13).Value = SearchResult(Count).DEFECT_TYPE
                    '入庫コメント
                    '2012/2/27 改修に伴い入庫コメントを追加
                    .Cells(14).Value = SearchResult(Count).REMARKS
                    'ロケーション
                    .Cells(15).Value = SearchResult(Count).LOCATION
                    '倉庫
                    .Cells(16).Value = SearchResult(Count).PLACE

                    .Cells(17).Value = SearchResult(Count).ID
                    .Cells(18).Value = SearchResult(Count).DETAIL_ID
                    .Cells(19).Value = SearchResult(Count).I_ID
                    .Cells(20).Value = SearchResult(Count).PLACE_ID
                End With
                DataGridView1.Rows.Add(SR_List)
            Next
        Else
            MsgBox(ErrorMessage)
            Exit Sub
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '検品チェックシート出力
        Dim OutPut_Flg As Boolean = False

        Dim Output_List() As Output_List = Nothing
        Dim Output_PL_List() As PL_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim PL_Result As Boolean = True
        Dim InSheet_ErrorMessage As String = Nothing
        Dim InSheet_Result As Boolean = True
        Dim PLList_ErrorMessage As String = Nothing
        Dim PLList_Result As Boolean = True
        Dim PL_List() As PL_List = Nothing
        Dim Output_Data_Count As Integer
        Dim Output_Prt_List() As In_Check_List = Nothing
        Dim DataCount As Integer = 0
        Dim LoopCount As Integer = 0
        'すべてのデータで印字する帳票がなかった時、メッセージを表示するためのフラグ
        Dim Print_FLG As Boolean = False

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now
        Dim MAXPage As Integer = 0
        Dim MAXData As Integer = 0

        '検品チェックシートの１ページあたりの表示件数
        'ライン
        Dim MAXData_Line As Integer = 38
        'ベイト&ガルプ
        Dim MAXData_Bait As Integer = 34
        'サオ
        Dim MAXData_Rods As Integer = 18
        'リール
        '2012.09.13
        '帳票内容変更に伴い17件→13件へ変更
        'Dim MAXData_Reel As Integer = 17
        Dim MAXData_Reel As Integer = 13

        'アクセサリー
        Dim MAXData_Accessory As Integer = 27
        'バッグ＆アパレル
        Dim MAXData_Bag As Integer = 26

        Dim WhereSQL As String = Nothing
        Dim WherePL As String = Nothing
        Dim PL_Type As String = Nothing
        Dim ListDataCount As Integer = 0

        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                OutPut_Flg = True
            End If
        Next

        If OutPut_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'プロダクトラインの種類を取得
        PL_Result = GetPLList(Output_PL_List, PL_Result, ErrorMessage)
        If PL_Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If


        '検品チェックシートに出力するデータIDを格納。
        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Output_List(0 To Output_Data_Count)
                'Output_List(Output_Data_Count).ID = DataGridView1.Rows(Count).Cells(13).Value()
                'Output_List(Output_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(14).Value()
                'ID
                Output_List(Output_Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                'DETAIL_ID
                Output_List(Output_Data_Count).DETAIL_ID = DataGridView1.Rows(Count).Cells(18).Value()
                Output_Data_Count += 1
            End If
        Next

        For i = 0 To Output_List.Length - 1

            If i <> Output_List.Length - 1 Then
                WhereSQL &= Output_List(i).DETAIL_ID & ","
            Else
                WhereSQL &= Output_List(i).DETAIL_ID
            End If
        Next

        'ページ設定ダイアログを開く
        Call AxReport1.PrinterSetupDlg()
        '設定を変数に保存
        nSvOrientation = AxReport1.Orientation
        nSvPaperSize = AxReport1.PaperSize
        nSvPaperLength = AxReport1.PaperLength
        nSvPaperWidth = AxReport1.PaperWidth
        nSvDefaultSource = AxReport1.DefaultSource
        sSvPrinterName = AxReport1.PrinterName
        'レポートを閉じる
        AxReport1.ReportPath = ""

        PL_Type = "ライン"

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If

        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            '****************************************************
            ' ラインのデータ取得、帳票の作成
            '*****************************************************

            'ラインの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            If ListDataCount <> 0 Then

                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Line.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Line.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'ラインの１ページの最大表示件数を設定
                MAXData = MAXData_Line

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-6").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next


                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        '****************************************************
        ' ベイトのデータ取得、帳票の作成
        '*****************************************************

        PL_Type = "ベイト＆ガルプ"
        WherePL = Nothing
        PLList_Result = True
        PLList_ErrorMessage = Nothing

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If

        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            Output_Prt_List = Nothing
            ListDataCount = 0

            'ベイトの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            If ListDataCount <> 0 Then
                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Bait.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Bait.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'ラインの１ページの最大表示件数を設定
                MAXData = MAXData_Bait

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-6").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next

                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        '****************************************************
        ' サオ（Rods）のデータ取得、帳票の作成
        '*****************************************************

        PL_Type = "サオ"
        WherePL = Nothing
        PLList_Result = True
        PLList_ErrorMessage = Nothing

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If
        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            Output_Prt_List = Nothing
            ListDataCount = 0

            'サオの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            If ListDataCount <> 0 Then
                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Rod.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Rod.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'サオの１ページの最大表示件数を設定
                MAXData = MAXData_Rods

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-6").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next


                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        '****************************************************
        ' リールのデータ取得、帳票の作成
        '*****************************************************
        PL_Type = "リール"
        WherePL = Nothing
        PLList_Result = True
        PLList_ErrorMessage = Nothing

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If
        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            Output_Prt_List = Nothing
            ListDataCount = 0

            'リールの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            'データが0件ならプレビューも表示しない。
            If ListDataCount <> 0 Then

                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Reel.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Reel.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'リールの１ページの最大表示件数を設定
                MAXData = MAXData_Reel

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-3").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-3").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next

                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        '****************************************************
        ' アクセサリーのデータ取得、帳票の作成
        '*****************************************************
        PL_Type = "アクセサリー"
        WherePL = Nothing
        PLList_Result = True
        PLList_ErrorMessage = Nothing

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If
        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            Output_Prt_List = Nothing
            ListDataCount = 0

            'アクセサリーの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            If ListDataCount <> 0 Then
                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Accessory.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Accessory.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'アクセサリーの１ページの最大表示件数を設定
                MAXData = MAXData_Accessory

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-6").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next

                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        '****************************************************
        ' バッグ＆アパレルのデータ取得、帳票の作成
        '*****************************************************

        PL_Type = "バッグ＆アパレル"
        WherePL = Nothing
        PLList_Result = True
        PLList_ErrorMessage = Nothing

        'プロダクトラインコードの種類を取得。
        PLList_Result = GetPLSheetTypeList(PL_List, PL_Type, PLList_Result, PLList_ErrorMessage)

        'If PLList_Result = False Then
        '    MsgBox(PLList_ErrorMessage)
        '    Exit Sub
        'End If
        If PLList_Result = True Then

            For i = 0 To PL_List.Length - 1
                If i <> PL_List.Length - 1 Then
                    WherePL &= PL_List(i).ID & ","
                Else
                    WherePL &= PL_List(i).ID
                End If
            Next

            Output_Prt_List = Nothing
            ListDataCount = 0

            'バッグ＆アパレルの出力情報を取得
            InSheet_Result = GetOutputList(WherePL, WhereSQL, Output_Prt_List, ListDataCount, InSheet_Result, InSheet_ErrorMessage)

            If InSheet_Result = False Then
                MsgBox(InSheet_ErrorMessage)
                Exit Sub
            End If

            If ListDataCount <> 0 Then
                'レポートを開く
                AxReport1.ReportPath = PrtForm & "CheckSheet-Bag.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '3項目目を-1にするとプレビュー画面を表示。それ以外は直接印刷。
                If AxReport1.OpenPrintJob("CheckSheet-Bag.crp", 512, -1, "検品チェックシート　プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                'バッグ＆アパレルの１ページの最大表示件数を設定
                MAXData = MAXData_Bag

                DataCount = 0
                LoopCount = 1

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Output_Prt_List.Length <= MAXData Then
                    MAXPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Output_Prt_List.Length Mod MAXData = 0 Then
                        MAXPage = Output_Prt_List.Length \ MAXData
                    Else
                        MAXPage = Output_Prt_List.Length \ MAXData + 1
                    End If
                End If

                For Page = 1 To MAXPage
                    '検品報告書欄に今日の日付を設定
                    AxReport1.Item("", "date").Text = dtNow.ToString("yyyy/MM/dd")
                    If Output_Prt_List.Length <= MAXData Then
                        LoopCount = Output_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MAXPage Then
                            LoopCount = (Output_Prt_List.Length) - ((Page - 1) * MAXData)
                        Else
                            LoopCount = MAXData
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To MAXData
                        '商品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '商品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = ""
                        '数量の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = ""
                        'JANの表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Next

                    For j = 1 To LoopCount
                        '商品コードの表示
                        AxReport1.Item("", "Label" & j & "-1").Text = Output_Prt_List(DataCount).I_CODE
                        '商品名の表示
                        AxReport1.Item("", "Label" & j & "-2").Text = " " & Output_Prt_List(DataCount).I_NAME
                        '数量の表示
                        AxReport1.Item("", "Label" & j & "-4").Text = Output_Prt_List(DataCount).NUM
                        'JANの表示
                        AxReport1.Item("", "Label" & j & "-6").Text = Output_Prt_List(DataCount).JAN
                        DataCount += 1
                    Next

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next

                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""

                '印字対象データがあったのでフラグをTrueにする
                Print_FLG = True
            End If
        End If

        If Print_FLG = False Then
            MsgBox("印刷するデータがありません。")
        End If
    End Sub
    'CurrentCellDirtyStateChangedイベントハンドラ
    Private Sub DataGridView1_CurrentCellDirtyStateChanged( _
            ByVal sender As Object, ByVal e As EventArgs) _
            Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.CurrentCellAddress.X = 0 AndAlso _
            DataGridView1.IsCurrentCellDirty Then
            'コミットする
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    'CellValueChangedイベントハンドラ
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, _
            ByVal e As DataGridViewCellEventArgs) _
            Handles DataGridView1.CellValueChanged
        Dim count As Integer
        Dim IndexCount As Integer = 20
        '列のインデックスを確認する
        If e.ColumnIndex = 0 Then
            If FormLord = True Then
                If DataGridView1(e.ColumnIndex, e.RowIndex).Value = 1 Then
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.PaleGreen
                    Next
                Else
                    For count = 0 To IndexCount
                        DataGridView1.Item(count, e.RowIndex).Style.BackColor = Color.White
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker1.Value.ToShortDateString()

        MaskedTextBox1.Text = Date.ParseExact(DateTimePicker1.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker2.Value.ToShortDateString()

        MaskedTextBox2.Text = Date.ParseExact(DateTimePicker2.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker3.Value.ToShortDateString()

        MaskedTextBox3.Text = Date.ParseExact(DateTimePicker3.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub DateTimePicker4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker4.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker4.Value.ToShortDateString()

        MaskedTextBox4.Text = Date.ParseExact(DateTimePicker4.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As Search_List = Nothing
        Dim Count As Integer = 0
        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Fix_Date_From As String = Nothing
        Dim Fix_Date_To As String = Nothing
        Dim Date_Check_Result As DateTime
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0
        Dim ChkItemCodeString As String = Nothing
        Dim ChkDocNoString As String = Nothing
        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        '一括チェックのチェックボックスをクリアする。
        'CheckBox7.Checked = False

        'DataGridView（検索結果）をクリアする。
        'DataGridView1.Rows.Clear()

        '検索ボタンクリックチェック
        'nSearchFLg = False

        '入荷予定日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            Date_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '入荷予定日Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            Date_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Date_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '入荷日Fromのチェック
        If MaskedTextBox3.Text = "    /  /" Then
            '未入力なら""を格納
            Fix_Date_From = ""
            MaskedTextBox3.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox3.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Fix_Date_From = MaskedTextBox3.Text
                MaskedTextBox3.BackColor = Color.White
            Else
                MsgBox("入荷日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '入荷日Toのチェック
        If MaskedTextBox4.Text = "    /  /" Then
            '未入力なら""を格納
            Fix_Date_To = ""
            MaskedTextBox4.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox4.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                Fix_Date_To = MaskedTextBox4.Text
                MaskedTextBox4.BackColor = Color.White
            Else
                MsgBox("入荷予定日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        ChkDocNoString = Trim(TextBox1.Text)
        ChkItemCodeString = Trim(TextBox6.Text)

        'ドキュメント№に'が入力されたいたらReplaceする。
        ChkDocNoString = ChkDocNoString.Replace("'", "''")

        '商品コードに'が入力されていたらReplaceする。
        ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        '検索Function
        Result = GetInSeach(ChkDocNoString, Date_From, Date_To, Fix_Date_From, Fix_Date_To, _
                            ChkItemCodeString, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, _
                            CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, PLACE_ID, SearchResult, Data_Total, _
                            Data_Num_Total, Result, ErrorMessage)
        If Result = False Then
            '商品数、総数をクリア
            Label11.Text = "商品数："
            Label13.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "入庫関連データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "ドキュメント№,商品コード,入庫元コード,入庫元名,予定数量,入荷予定日,入荷数量,入荷日,商品名,JANコード,ステータス,種別,不良区分,入庫コメント,倉庫"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If Data_Total = 0 Then
            MsgBox("データがありません。")
            Exit Sub
        End If

        'ﾌｧｲﾙ名の設定
        strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)
        'headerを設定
        strStreamWriter.WriteLine(Sheet1_Header)

        'データ件数分ループ
        For i = 0 To SearchResult.Length - 1

            '取得するデータをLineataに格納する。
            'A列にドキュメント№
            LineData = """" & SearchResult(i).DOC_NO & ""","
            'B列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'C列に入庫元コード
            LineData &= """" & SearchResult(i).C_CODE & ""","
            'D列に入庫元名
            LineData &= """" & SearchResult(i).C_NAME & ""","
            'E列に予定数量
            LineData &= """" & SearchResult(i).NUM & ""","
            'F列に入荷予定日
            LineData &= """" & SearchResult(i).N_DATE & ""","
            'G列に入荷数量
            LineData &= """" & SearchResult(i).FIX_NUM & ""","
            'H列に入荷日
            LineData &= """" & SearchResult(i).FIX_DATE & ""","
            'I列に商品名
            LineData &= """" & SearchResult(i).I_NAME & ""","
            'J列にJANコード
            LineData &= """" & SearchResult(i).JAN_CODE & ""","
            'K列にステータス
            LineData &= """" & SearchResult(i).STATUS & ""","
            'L列に種別
            LineData &= """" & SearchResult(i).CATEGORY & ""","
            'M列に不良区分
            LineData &= """" & SearchResult(i).DEFECT_TYPE & ""","
            '2012/2/27 入庫コメント追加
            'N列に入庫コメント
            LineData &= """" & SearchResult(i).REMARKS & ""","
            'O列に倉庫
            LineData &= """" & SearchResult(i).PLACE & """"
            'O列にロケーション
            'LineData &= """" & SearchResult(i).LOCATION & ""","

            '2012/2/27 原様依頼により、入庫元コード、入庫元名追加に伴い改修
            ''C列に予定数量
            'LineData &= """" & SearchResult(i).NUM & ""","
            ''D列に入荷予定日
            'LineData &= """" & SearchResult(i).N_DATE & ""","
            ''E列に入荷数量
            'LineData &= """" & SearchResult(i).FIX_NUM & ""","
            ''F列に入荷日
            'LineData &= """" & SearchResult(i).FIX_DATE & ""","
            ''G列に商品名
            'LineData &= """" & SearchResult(i).I_NAME & ""","
            ''H列にJANコード
            'LineData &= """" & SearchResult(i).JAN_CODE & ""","
            ''I列にステータス
            'LineData &= """" & SearchResult(i).STATUS & ""","
            ''J列に種別
            'LineData &= """" & SearchResult(i).CATEGORY & ""","
            ''K列に不良区分
            'LineData &= """" & SearchResult(i).DEFECT_TYPE & ""","
            ''L列にロケーション
            'LineData &= """" & SearchResult(i).LOCATION & ""","


            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub
End Class