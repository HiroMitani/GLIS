Imports System.Windows.Forms
Imports System.Globalization

Public Class syukokensaku

    Dim PLACE_ID As String

    Public FormLord As Boolean = False

    'False:検索ボタンを押したらFalse
    '入庫確定や変更など行ったら、再度検索を行わせる為のFlg
    Public sSearchFLg As Boolean = False

    '帳票ファイルの格納フォルダ指定
    Public PrtForm As String = System.Configuration.ConfigurationManager.AppSettings("PrtPath")

    '出荷伝票の出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Public Structure Out_Change_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_ID As Integer
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim C_CODE As String
        Dim C_NAME As String
        Dim NUM As Integer
        Dim O_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim File_NAME As String
        Dim STATUS As String
        Dim CATEGORY As String
        Dim DEFECT_TYPE As String
        Dim COST As Integer
        Dim PRICE As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim REMARKS As String
        Dim PRT_DATE As String
        Dim P_ID As Integer
    End Structure

    Private Sub syukokensaku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "出庫関連検索"

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

        '左4項目を固定(チェック、伝票番号、オーダー番号、商品コード)
        DataGridView1.Columns(3).Frozen = True

        '出荷指示ファイル名にフォーカスを移動。
        TextBox3.Focus()

        ComboBox1.Text = "商品コード"

        FormLord = True

    End Sub

    Private Sub syukokensaku_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '出庫確定ボタンクリック

        Dim Check_Flg As Boolean = False

        Dim Out_Check_Flg As Boolean = False
        Dim OutDefinition_Check_Message As String = Nothing

        Dim OutDefinition_Data_Count As Integer = 0
        Dim OutDefinition_Check() As Out_Search_List = Nothing
        Dim Data_Num_Total As Integer = 0

        Dim Stock_Result As Boolean = True
        Dim Stock_Result_Message As String = Nothing
        Dim Stock_Num As Integer = 0

        If sSearchFLg = True Then
            MsgBox("出荷確定、出荷予定変更を行った後は再度検索を行って下さい。")
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
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Out_Check_Flg = True
        'チェックされた商品の中でピッキング済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷予定" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "出荷済み" OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("出荷予定、ピッキング戻し、出荷済み、伝票出力のみのデータにチェックされています。ピッキング済みのデータのみ出荷確定が行えます。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve OutDefinition_Check(0 To OutDefinition_Data_Count)
                '伝票番号
                OutDefinition_Check(OutDefinition_Data_Count).SHEET_NO = DataGridView1.Rows(Count).Cells(1).Value()
                'オーダー番号
                OutDefinition_Check(OutDefinition_Data_Count).ORDER_NO = DataGridView1.Rows(Count).Cells(2).Value()
                '商品コード
                OutDefinition_Check(OutDefinition_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(3).Value()
                '商品名
                OutDefinition_Check(OutDefinition_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(4).Value()
                '出荷指示ファイル名
                OutDefinition_Check(OutDefinition_Data_Count).FILE_NAME = DataGridView1.Rows(Count).Cells(5).Value()

                '納品先コード
                OutDefinition_Check(OutDefinition_Data_Count).C_CODE = DataGridView1.Rows(Count).Cells(6).Value()
                '納品先名
                OutDefinition_Check(OutDefinition_Data_Count).C_NAME = DataGridView1.Rows(Count).Cells(7).Value()
                '出荷予定数
                OutDefinition_Check(OutDefinition_Data_Count).NUM = DataGridView1.Rows(Count).Cells(8).Value()
                '出荷予定日
                OutDefinition_Check(OutDefinition_Data_Count).O_DATE = DataGridView1.Rows(Count).Cells(9).Value()
                '出荷数量は出荷予定数を入れる。
                OutDefinition_Check(OutDefinition_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(8).Value()
                '出荷日は今日の日付を入れる。
                OutDefinition_Check(OutDefinition_Data_Count).FIX_DATE = DateTime.Today


                'ステータス
                OutDefinition_Check(OutDefinition_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(12).Value()
                'カテゴリー
                OutDefinition_Check(OutDefinition_Data_Count).CATEGORY = DataGridView1.Rows(Count).Cells(13).Value()

                '不良区分
                OutDefinition_Check(OutDefinition_Data_Count).DEFECT_TYPE = DataGridView1.Rows(Count).Cells(14).Value()

                '印刷日
                OutDefinition_Check(OutDefinition_Data_Count).PRT_DATE = DataGridView1.Rows(Count).Cells(15).Value()

                '納入単価
                OutDefinition_Check(OutDefinition_Data_Count).COST = DataGridView1.Rows(Count).Cells(16).Value()
                '売単価
                OutDefinition_Check(OutDefinition_Data_Count).PRICE = DataGridView1.Rows(Count).Cells(17).Value()

                'コメント１
                OutDefinition_Check(OutDefinition_Data_Count).COMMENT1 = DataGridView1.Rows(Count).Cells(18).Value()

                'コメント２
                OutDefinition_Check(OutDefinition_Data_Count).COMMENT2 = DataGridView1.Rows(Count).Cells(19).Value()

                '備考
                OutDefinition_Check(OutDefinition_Data_Count).REMARKS = DataGridView1.Rows(Count).Cells(20).Value()
                '倉庫
                OutDefinition_Check(OutDefinition_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(21).Value()


                'OUT.ID
                OutDefinition_Check(OutDefinition_Data_Count).ID = DataGridView1.Rows(Count).Cells(22).Value()
                'I_ID
                OutDefinition_Check(OutDefinition_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(23).Value()
                'P_ID
                OutDefinition_Check(OutDefinition_Data_Count).P_ID = DataGridView1.Rows(Count).Cells(24).Value()

                '配列に格納される予定数量を足していき、総数を表示する。
                Data_Num_Total += DataGridView1.Rows(Count).Cells(8).Value()
                OutDefinition_Data_Count += 1
            End If
        Next

        'InDefinition_Data_Countが確定画面に表示する件数になるので、DataGridViewの上に表示する商品数はInDefinition_Data_Countを表示
        syukokakutei.Label1.Text = "商品数： " & OutDefinition_Data_Count
        '総数を表示
        syukokakutei.Label3.Text = "総数： " & Data_Num_Total

        'nkakuteiのDataGridViewをクリア
        nkakutei.DataGridView1.Rows.Clear()

        '出庫確定のDataGridViewに表示する。
        For Count = 0 To OutDefinition_Check.Length - 1
            Dim Out_Definition_Data_list As New DataGridViewRow
            Out_Definition_Data_list.CreateCells(syukokakutei.DataGridView1)
            With Out_Definition_Data_list
                '伝票番号
                .Cells(0).Value = OutDefinition_Check(Count).SHEET_NO
                'オーダー番号
                .Cells(1).Value = OutDefinition_Check(Count).ORDER_NO
                '商品コード
                .Cells(2).Value = OutDefinition_Check(Count).I_CODE
                '商品名
                .Cells(3).Value = OutDefinition_Check(Count).I_NAME
                '納品先コード
                .Cells(4).Value = OutDefinition_Check(Count).C_CODE
                '納品先名
                .Cells(5).Value = OutDefinition_Check(Count).C_NAME
                '予定数量
                .Cells(6).Value = OutDefinition_Check(Count).NUM
                '出荷予定日
                .Cells(7).Value = OutDefinition_Check(Count).O_DATE
                '出荷数量
                .Cells(8).Value = OutDefinition_Check(Count).FIX_NUM
                '出荷日
                .Cells(9).Value = OutDefinition_Check(Count).FIX_DATE
                '出荷指示ファイル名
                .Cells(10).Value = OutDefinition_Check(Count).FILE_NAME
                '納入単価
                .Cells(11).Value = OutDefinition_Check(Count).COST
                '売単価
                .Cells(12).Value = OutDefinition_Check(Count).PRICE
                'ステータス
                .Cells(13).Value = OutDefinition_Check(Count).STATUS
                'カテゴリー
                .Cells(14).Value = OutDefinition_Check(Count).CATEGORY
                '不良区分
                .Cells(15).Value = OutDefinition_Check(Count).DEFECT_TYPE
                '印刷日
                .Cells(16).Value = OutDefinition_Check(Count).PRT_DATE
                'コメント１
                .Cells(17).Value = OutDefinition_Check(Count).COMMENT1
                'コメント２
                .Cells(18).Value = OutDefinition_Check(Count).COMMENT2
                '備考
                .Cells(19).Value = OutDefinition_Check(Count).REMARKS
                '倉庫
                .Cells(20).Value = OutDefinition_Check(Count).PLACE
                'ID
                .Cells(21).Value = OutDefinition_Check(Count).ID
                'I_ID　
                .Cells(22).Value = OutDefinition_Check(Count).I_ID
                'P_ID　
                .Cells(23).Value = OutDefinition_Check(Count).P_ID
            End With
            syukokakutei.DataGridView1.Rows.Add(Out_Definition_Data_list)
        Next

        syukokakutei.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '旧出荷予定変更
        'コメント修正用に変更 2014/12/22

        Dim Upd_Check_Flg As Boolean = False
        Dim Upd_Data_Count As Integer = 0
        Dim Upd_Check() As Out_Change_List = Nothing

        Dim ALLCheck_Flg As Boolean = True

        If sSearchFLg = True Then
            MsgBox("再度検索を行って下さい。")
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
                Upd_Check_Flg = True
            End If
        Next

        If Upd_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("コメント修正するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        Upd_Check_Flg = True
        'チェックされた商品の中で出荷予定以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso _
            (DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" OrElse _
             DataGridView1.Rows(Count).Cells(12).Value() = "出荷済み" OrElse _
             DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Upd_Check_Flg = False
            End If
        Next
        If Upd_Check_Flg = False Then
            MsgBox("ピッキング戻し、出荷済み、伝票出力のみのデータがチェックされています。出荷予定、ピッキング済みのデータのみコメント修正が行えます。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("コメントを修正をするには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Upd_Check(0 To Upd_Data_Count)
                '伝票番号
                Upd_Check(Upd_Data_Count).SHEET_NO = DataGridView1.Rows(Count).Cells(1).Value()
                'オーダー番号
                Upd_Check(Upd_Data_Count).ORDER_NO = DataGridView1.Rows(Count).Cells(2).Value()
                '商品コード
                Upd_Check(Upd_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(3).Value()
                '商品名
                Upd_Check(Upd_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(4).Value()
                '出荷指示ファイル名
                Upd_Check(Upd_Data_Count).File_NAME = DataGridView1.Rows(Count).Cells(5).Value()

                '納品先コード
                Upd_Check(Upd_Data_Count).C_CODE = DataGridView1.Rows(Count).Cells(6).Value()
                '納品先名
                Upd_Check(Upd_Data_Count).C_NAME = DataGridView1.Rows(Count).Cells(7).Value()
                '予定数量
                Upd_Check(Upd_Data_Count).NUM = DataGridView1.Rows(Count).Cells(8).Value()
                '出荷予定日
                Upd_Check(Upd_Data_Count).O_DATE = DataGridView1.Rows(Count).Cells(9).Value()
                '出荷数量
                '出荷数量はNULLなら、0に置き換える。
                'If DataGridView1.Rows(Count).Cells(9).Value() = "" Then
                '    Upd_Check(Upd_Data_Count).FIX_NUM = 0
                'Else
                '    Upd_Check(Upd_Data_Count).FIX_NUM = DataGridView1.Rows(Count).Cells(9).Value()
                'End If
                ''出荷日
                'Upd_Check(Upd_Data_Count).FIX_DATE = DataGridView1.Rows(Count).Cells(10).Value()


                'ステータス
                Upd_Check(Upd_Data_Count).STATUS = DataGridView1.Rows(Count).Cells(12).Value()
                'カテゴリー
                Upd_Check(Upd_Data_Count).CATEGORY = DataGridView1.Rows(Count).Cells(13).Value()
                '不良区分
                Upd_Check(Upd_Data_Count).DEFECT_TYPE = DataGridView1.Rows(Count).Cells(14).Value()
                ''印刷日
                'Upd_Check(Upd_Data_Count).PRT_DATE = DataGridView1.Rows(Count).Cells(15).Value()
                '納入単価
                Upd_Check(Upd_Data_Count).COST = DataGridView1.Rows(Count).Cells(16).Value()
                '売単価
                Upd_Check(Upd_Data_Count).PRICE = DataGridView1.Rows(Count).Cells(17).Value()
                'コメント１
                Upd_Check(Upd_Data_Count).COMMENT1 = DataGridView1.Rows(Count).Cells(18).Value()
                'コメント２
                Upd_Check(Upd_Data_Count).COMMENT2 = DataGridView1.Rows(Count).Cells(19).Value()
                '備考
                Upd_Check(Upd_Data_Count).REMARKS = DataGridView1.Rows(Count).Cells(20).Value()
                'OUT_HEADER.ID
                Upd_Check(Upd_Data_Count).ID = DataGridView1.Rows(Count).Cells(22).Value()
                'I_ID
                Upd_Check(Upd_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(23).Value()
                'P_ID
                Upd_Check(Upd_Data_Count).P_ID = DataGridView1.Rows(Count).Cells(24).Value()
                Upd_Data_Count += 1
            End If
        Next

        For Count = 0 To Upd_Check.Length - 1
            Dim Change_Data_list As New DataGridViewRow
            Change_Data_list.CreateCells(syukohenkou.DataGridView1)
            With Change_Data_list
                '伝票番号
                .Cells(0).Value = Upd_Check(Count).SHEET_NO
                'オーダー番号
                .Cells(1).Value = Upd_Check(Count).ORDER_NO
                '商品コード
                .Cells(2).Value = Upd_Check(Count).I_CODE
                '商品名
                .Cells(3).Value = Upd_Check(Count).I_NAME
                '納品先コード
                .Cells(4).Value = Upd_Check(Count).C_CODE
                '納品先名
                .Cells(5).Value = Upd_Check(Count).C_NAME
                '予定数量
                .Cells(6).Value = Upd_Check(Count).NUM
                '出荷予定日
                .Cells(7).Value = Upd_Check(Count).O_DATE
                ''出荷数量
                'If Upd_Check(Count).FIX_NUM = 0 Then
                '    .Cells(8).Value = ""
                'Else
                '    .Cells(8).Value = Upd_Check(Count).FIX_NUM
                'End If
                ''出荷日
                '.Cells(9).Value = Upd_Check(Count).FIX_DATE


                '出荷指示ファイル名
                .Cells(10).Value = Upd_Check(Count).File_NAME
                '納入単価
                .Cells(11).Value = Upd_Check(Count).COST
                '売単価
                .Cells(12).Value = Upd_Check(Count).PRICE
                'ステータス
                .Cells(13).Value = Upd_Check(Count).STATUS
                'カテゴリー
                .Cells(14).Value = Upd_Check(Count).CATEGORY
                '不良区分
                .Cells(15).Value = Upd_Check(Count).DEFECT_TYPE
                ''印刷日
                '.Cells(16).Value = Upd_Check(Count).PRT_DATE
                'コメント１
                .Cells(17).Value = Upd_Check(Count).COMMENT1
                'コメント２
                .Cells(18).Value = Upd_Check(Count).COMMENT2
                '備考
                .Cells(19).Value = Upd_Check(Count).REMARKS
                'ID
                .Cells(20).Value = Upd_Check(Count).ID
                'I_ID
                .Cells(21).Value = Upd_Check(Count).I_ID
                'P_ID
                .Cells(22).Value = Upd_Check(Count).P_ID
            End With
            syukohenkou.DataGridView1.Rows.Add(Change_Data_list)
        Next

        syukohenkou.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click


        '削除
        'Dim Delete_Check_Flg As Boolean = False
        'Dim Del_Check() As Item_List = Nothing
        'Dim Del_Data_Count As Integer = 0
        'Dim ErrorMessage As String = Nothing
        'Dim Del_Result As Boolean = True

        ''データが０件ならエラー
        'If DataGridView1.Rows.Count = 0 Then
        '    MsgBox("データが選択されていません。")
        '    Exit Sub
        'End If

        ''チェックされた商品があるか確認
        'For Count = 0 To DataGridView1.Rows.Count - 1
        '    If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
        '        Delete_Check_Flg = True
        '    End If
        'Next

        'If Delete_Check_Flg = False Then
        '    MsgBox("チェックされた商品がありません。")
        '    Exit Sub
        'End If

        'Delete_Check_Flg = True
        ''チェックされた商品の中で入荷済みがチェックされていないか確認
        'For Count = 0 To DataGridView1.Rows.Count - 1
        '    If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(12).Value() = "入荷済み" Then
        '        Delete_Check_Flg = False
        '    End If
        'Next
        'If Delete_Check_Flg = False Then
        '    MsgBox("出荷済みのデータがチェックされています。出荷済みのデータは削除できません。")
        '    Exit Sub
        'End If

        ''ダイアログ設定
        'Dim result As DialogResult = MessageBox.Show("データ削除してもよろしいですか？", _
        '                                             "確認", _
        '                                             MessageBoxButtons.YesNo, _
        '                                             MessageBoxIcon.Question)
        ''何が選択されたか調べる()
        'If result = DialogResult.No Then
        '    '「いいえ」が選択された時 
        '    Exit Sub
        'End If

        ''DataGridViewからチェックされたデータのIDを配列に格納
        'For Count = 0 To DataGridView1.Rows.Count - 1
        '    If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
        '        ReDim Preserve Del_Check(0 To Del_Data_Count)
        '        'IDを格納
        '        Del_Check(Del_Data_Count).ID = DataGridView1.Rows(Count).Cells(21).Value()
        '        Del_Data_Count += 1
        '    End If
        'Next

        ''削除用Function
        'Del_Result = Out_Del_Item(Del_Check, Del_Result, ErrorMessage)

        'If Del_Result = False Then
        '    MsgBox(ErrorMessage)
        'End If

        'MsgBox("削除が完了しました。")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As Out_Search_List = Nothing
        Dim Count As Integer = 0

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Fix_Date_From As String = Nothing
        Dim Fix_Date_To As String = Nothing

        Dim Date_Check_Result As DateTime

        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim FileNameCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing

        Dim ItemJan_Flg As Integer = 0

        Dim ChkSheetNoString As String = Nothing
        Dim ChkOrderNoString As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ChkCommentString As String = Nothing
        Dim ChkCCodeString As String = Nothing

        Dim ChkClaimCodeString As String = Nothing

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        '検索ボタンクリックチェック
        sSearchFLg = False

        '一括チェックのチェックボックスのクリア
        CheckBox9.Checked = False

        '出荷指示ファイル名称の文字列妥当性チェック（SQLを実行するための'の対策）
        If Trim(TextBox3.Text) <> "" Then
            '文字の妥当性チェックを行う。
            'NullOK、
            If StringChkVal(Trim(TextBox3.Text), True, False, FileNameCheckAfter, StringChkResult, StringErrorMessage) = False Then
                TextBox3.BackColor = Color.Salmon
                MsgBox("出荷指示ファイル名に不正な文字が入力されています。")
            Else
                'チェックに問題がなければ背景色を白に戻す。
                TextBox3.BackColor = Color.White
            End If
        End If

        '出荷予定日Fromのチェック
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
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '出荷予定日Toのチェック
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
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '出荷日Fromのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '出荷日Toのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox1.Text = "商品コード" Then
            ItemJan_Flg = 1
            ComboBox1.BackColor = Color.White
        ElseIf ComboBox1.Text = "JANコード" Then
            ItemJan_Flg = 2
            ComboBox1.BackColor = Color.White
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox1.BackColor = Color.Salmon
            ComboBox1.Focus()
            Exit Sub
        End If

        ChkSheetNoString = Trim(TextBox1.Text)
        ChkOrderNoString = Trim(TextBox2.Text)
        ChkItemJanString = Trim(TextBox7.Text)
        ChkCommentString = Trim(TextBox8.Text)
        ChkCCodeString = Trim(TextBox4.Text)

        ChkClaimCodeString = Trim(TextBox5.Text)


        '伝票番号に'が入力されたいたらReplaceする。
        ChkSheetNoString = ChkSheetNoString.Replace("'", "''")

        'オーダー番号に'が入力されていたらReplaceする。
        ChkOrderNoString = ChkOrderNoString.Replace("'", "''")

        'アイテムコードorJANコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        'アイテムコードorJANコードに'が入力されていたらReplaceする。
        ChkCommentString = ChkCommentString.Replace("'", "''")

        '納品先コードに'が入っていたらReplaceする。
        ChkCCodeString = ChkCCodeString.Replace("'", "''")

        '出荷予定検索Function
        Result = GetOutSearch(FileNameCheckAfter, ChkSheetNoString, ChkOrderNoString, Date_From, Date_To, Fix_Date_From, Fix_Date_To, _
                            ItemJan_Flg, ChkItemJanString, ChkCommentString, ChkCCodeString, ChkClaimCodeString, CheckBox3.Checked, _
                            CheckBox11.Checked, CheckBox10.Checked, CheckBox4.Checked, CheckBox7.Checked, CheckBox5.Checked, _
                            CheckBox6.Checked, CheckBox1.Checked, CheckBox2.Checked, PLACE_ID, _
                            SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            Label14.Text = "商品数："
            Label15.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '件数を表示
        Label14.Text = "商品数：" & Data_Total
        '総商品数を表示
        Label15.Text = "総数：" & Data_Num_Total

        If Result = True Then
            'DataGridへ入力したデータを挿入
            For Count = 0 To SearchResult.Length - 1
                Dim SR_List As New DataGridViewRow
                SR_List.CreateCells(DataGridView1)
                With SR_List
                    '伝票番号
                    .Cells(1).Value = SearchResult(Count).SHEET_NO
                    'オーダー番号
                    .Cells(2).Value = SearchResult(Count).ORDER_NO
                    '商品コード
                    .Cells(3).Value = SearchResult(Count).I_CODE
                    '商品名
                    .Cells(4).Value = SearchResult(Count).I_NAME
                    '出荷指示ファイル名
                    .Cells(5).Value = SearchResult(Count).FILE_NAME

                    '納品先コード
                    .Cells(6).Value = SearchResult(Count).C_CODE
                    '納品先名
                    .Cells(7).Value = SearchResult(Count).C_NAME
                    '出荷予定数量
                    .Cells(8).Value = SearchResult(Count).NUM
                    '出荷予定日
                    .Cells(9).Value = SearchResult(Count).O_DATE
                    'ステータスが出荷済みのものは出荷数を表示し、入荷予定のものは空白表示
                    If SearchResult(Count).STATUS = "出荷済み" Then
                        .Cells(10).Value = SearchResult(Count).FIX_NUM
                    Else
                        .Cells(10).Value = ""
                    End If

                    '出荷日
                    .Cells(11).Value = SearchResult(Count).FIX_DATE


                    'ステータス
                    .Cells(12).Value = SearchResult(Count).STATUS
                    'カテゴリー
                    .Cells(13).Value = SearchResult(Count).CATEGORY
                    '不良区分
                    .Cells(14).Value = SearchResult(Count).DEFECT_TYPE
                    '印刷日
                    .Cells(15).Value = SearchResult(Count).PRT_DATE
                    '納入単価
                    .Cells(16).Value = SearchResult(Count).COST
                    '売単価
                    .Cells(17).Value = SearchResult(Count).PRICE
                    'コメント１
                    .Cells(18).Value = SearchResult(Count).COMMENT1
                    'コメント２
                    .Cells(19).Value = SearchResult(Count).COMMENT2
                    '備考
                    .Cells(20).Value = SearchResult(Count).REMARKS
                    '倉庫
                    .Cells(21).Value = SearchResult(Count).PLACE
                    'ID
                    .Cells(22).Value = SearchResult(Count).ID
                    'I_ID
                    .Cells(23).Value = SearchResult(Count).I_ID
                    '倉庫ID
                    .Cells(24).Value = SearchResult(Count).P_ID
                    '郵便番号
                    .Cells(25).Value = SearchResult(Count).D_ZIP
                    '住所
                    .Cells(26).Value = SearchResult(Count).D_ADDRESS

                End With
                DataGridView1.Rows.Add(SR_List)
            Next
        Else
            MsgBox(ErrorMessage)
            Exit Sub
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
        Dim IndexCount As Integer = 24
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


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Picking_Check_Flg As Boolean = False
        Dim Picking_Data_Count As Integer = 0

        Dim Count As Integer = 0

        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing

        Dim Picking_Prt_List_Result As Boolean = True
        Dim Picking_Prt_List_ErrorMessage As String = Nothing

        '出荷指示ファイル名
        Dim FILE_NAME As String = Nothing

        Dim ALLCheck_Flg As Boolean = True

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 32
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0

        '現在のデータ件数
        Dim NowCount As Integer = 0

        Dim Picking_Prt_List() As Picking_Prt_List = Nothing

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        Dim ReturnDataCount As Integer = 0

        Dim PL_Result As Boolean = True
        Dim PL_ErrorMessage As String = Nothing
        Dim PL_List() As PL_List = Nothing

        Dim ModCheck As Integer = 0

        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        Picking_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング済み" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" OrElse DataGridView1.Rows(Count).Cells(12).Value() = "出荷済み" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Picking_Check_Flg = False
            End If
        Next
        If Picking_Check_Flg = False Then
            MsgBox("ピッキング済み、ピッキング戻し、出荷済み、伝票出力のみのデータにチェックされています。出荷予定のデータのみピッキングリストの出力が行えます。")
            Exit Sub
        End If

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("ピッキングファイルを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("ピッキングファイルを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        'プロダクトライン情報を取得する。
        PL_Result = GetPLList(PL_List, PL_Result, PL_ErrorMessage)

        FILE_NAME = Trim(TextBox3.Text)

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

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "PickingList.crp"
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
        If AxReport1.OpenPrintJob("PickingList.crp", 512, -1, "ピッキングリスト　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            Exit Sub
        End If

        For PLCount = 0 To PL_List.Length - 1
            ReturnDataCount = 0
            Picking_Prt_List = Nothing

            Picking_Prt_List_Result = GetPickingPrtData(FILE_NAME, PL_List(PLCount).ID, Picking_Prt_List, ReturnDataCount, Picking_Prt_List_Result, Picking_Prt_List_ErrorMessage)
            If Picking_Prt_List_Result = False Then
                MsgBox(Picking_Prt_List_ErrorMessage)
                '印刷ＪＯＢの終了（ファイルを閉じる） 
                AxReport1.ClosePrintJob(True)

                'レポーを閉じる 
                AxReport1.ReportPath = ""
                Exit Sub
            End If

            '結果が0件なら印刷しない
            If ReturnDataCount <> 0 Then

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Picking_Prt_List.Length <= Max Then
                    MaxPage = 1
                Else
                    'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                    If Picking_Prt_List.Length Mod Max = 0 Then
                        MaxPage = Picking_Prt_List.Length \ Max
                    Else
                        MaxPage = Picking_Prt_List.Length \ Max + 1
                    End If
                End If

                DataCount = 0
                For Page = 1 To MaxPage
                    'Headerの設定（日付＋出荷指示ファイル名）
                    AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "ピッキングリスト（" & Picking_Prt_List(0).FILE_NAME & "）"

                    'カテゴリーの表示
                    AxReport1.Item("", "category").Text = Picking_Prt_List(0).PL_NAME
                    'カテゴリーの合計数量を表示
                    AxReport1.Item("", "num").Text = ReturnDataCount

                    LineCount = 1
                    'Max件数以下なら、データ件数分ループするように設定
                    If Picking_Prt_List.Length <= Max Then
                        LoopCount = Picking_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MaxPage Then
                            LoopCount = Picking_Prt_List.Length - ((Page - 1) * Max)
                        Else
                            LoopCount = Max
                        End If

                    End If
                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To Max
                        '製品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '製品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = " "
                        '出荷数の表示
                        AxReport1.Item("", "Label" & i & "-3").Text = " "
                        '在庫数の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = " "
                        '在庫数 - 出荷数の表示
                        AxReport1.Item("", "Label" & i & "-5").Text = " "
                    Next

                    For i = 1 To LoopCount

                        '製品コードの表示
                        AxReport1.Item("", "Label" & LineCount & "-1").Text = Picking_Prt_List(DataCount).I_CODE
                        '製品名の表示
                        AxReport1.Item("", "Label" & LineCount & "-2").Text = Picking_Prt_List(DataCount).I_NAME
                        '出荷数の表示
                        AxReport1.Item("", "Label" & LineCount & "-3").Text = Picking_Prt_List(DataCount).NUM
                        '在庫数の表示
                        AxReport1.Item("", "Label" & LineCount & "-4").Text = Picking_Prt_List(DataCount).STOCK_NUM
                        '在庫数 - 出荷数の表示
                        AxReport1.Item("", "Label" & LineCount & "-5").Text = Picking_Prt_List(DataCount).STOCK_NUM - Picking_Prt_List(DataCount).NUM
                        LineCount += 1
                        DataCount += 1
                    Next

                    'Footerの設定
                    AxReport1.Item("", "footer").Text = Page & "/ " & MaxPage & "　ページ"


                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next

            End If
        Next
        '印刷ＪＯＢの終了（ファイルを閉じる） 
        AxReport1.ClosePrintJob(True)

        'レポーを閉じる 
        AxReport1.ReportPath = ""

        'ピッキング済みにUpdateし、在庫から数量をひき、STOCKログにピッキング済み情報を追加するFunction
        '（本来はピッキングリスト出力後）倉庫情報は１件目の倉庫ＩＤから取得
        '
        Result = Out_Picking_Data(FILE_NAME, DataGridView1.Rows(0).Cells(24).Value(), Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("ステータスを出荷予定からピッキング済みに更新しました。")

    End Sub

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim ItemName() As Item_List = Nothing
        Dim Search_Flg As Integer

        If ComboBox1.Text = "商品コード" Then
            Search_Flg = 1
        ElseIf ComboBox1.Text = "JANコード" Then
            Search_Flg = 2

        Else
            '未選択なら処理をしない
        End If

        '商品名欄をクリアする。
        Label8.Text = ""

        If Trim(ComboBox1.Text) <> "" Then

            'Enterキーが押されているか確認
            If e.KeyCode = Keys.Enter Then
                '入力チェック
                If Trim(TextBox7.Text) <> "" Then
                    'あたかもTabキーが押されたかのようにする
                    'Shiftが押されている時は前のコントロールのフォーカスを移動
                    Me.ProcessTabKey(Not e.Shift)
                    e.Handled = True
                    '入力された商品コードを元に商品名を取得する。
                    'ログインチェックFunction
                    Result = GetItemName(Trim(TextBox7.Text), Search_Flg, ItemName, Result, ErrorMessage)
                    If Result = "True" Then
                        Label8.Text = ItemName(0).I_NAME
                        CheckBox1.Focus()
                        TextBox8.BackColor = Color.White
                    ElseIf Result = "False" Then
                        MsgBox(ErrorMessage)
                        TextBox7.Focus()
                        TextBox7.BackColor = Color.Salmon
                        'エラーの場合、商品名もクリア。
                        Label8.Text = ""
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label8.Text = ""
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim Delivery_Check_Flg As Boolean = False
        Dim Delivery_Data_Count As Integer = 0
        Dim Delivery_Check() As DeliveryID_List = Nothing
        Dim Delivery_Data_List() As Delivery_List = Nothing
        Dim Delivery_Prt_List() As Delivery_List = Nothing
        Dim FILE_NAME As String = Nothing
        Dim ALLCheck_Flg As Boolean = True
        Dim Out_Check_Flg As Boolean = True
        Dim Count As Integer = 0
        Dim Delivery_Result As Boolean = True
        Dim Delivery_ErrorMessage As String = Nothing

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 32
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0

        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0
        Dim TotalNum As Integer = 0
        Dim WhereSql As String = Nothing
        Dim PrtPage As Integer = 1
        Dim C_Code As String = Nothing
        Dim DataResultCount As Integer = 0

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        Out_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("伝票出力のみのデータにチェックされています。伝票出力のみのデータは納品先別出荷リストの出力は行えません。")
            Exit Sub
        End If

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("納品先別出荷リストを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("納品先別出荷リストを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        FILE_NAME = Trim(TextBox3.Text)
        C_Code = Trim(TextBox4.Text)
        C_Code = C_Code.Replace("'", "''")

        Delivery_Result = GetDeliveryList(FILE_NAME, C_Code, Delivery_Data_List, WhereSql, DataResultCount, Delivery_Result, Delivery_ErrorMessage)
        If Delivery_Result = False Then
            MsgBox(Delivery_ErrorMessage)
            Exit Sub
        End If

        If DataResultCount <> 0 Then

            Delivery_Result = True
            Delivery_ErrorMessage = Nothing

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

            'レポートを開く
            AxReport1.ReportPath = PrtForm & "DeliveryShippingList.crp"
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
            If AxReport1.OpenPrintJob("DeliveryShippingList.crp", 512, -1, "納品先別出荷リスト　プレビュー", 0) = False Then
                'エラー処理を記述します 
                MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                Exit Sub
            End If

            '納品先の一覧を取得できたので、件数分ループし、納品先別ごとの商品情報を取得。
            For Count = 0 To Delivery_Data_List.Length - 1
                TotalNum = 0
                Delivery_Result = GetDeliveryPrtList(Delivery_Data_List(Count).C_ID, FILE_NAME, Delivery_Prt_List, TotalNum, Delivery_Result, Delivery_ErrorMessage)
                If Delivery_Result = False Then
                    MsgBox(Delivery_ErrorMessage)
                    Exit Sub
                End If
                'ありえないが、もし0件だったらエラーメッセージを出して終了。
                If Delivery_Prt_List.Length = 0 Then
                    MsgBox("納品先商品データがみつかりませんでした。")
                    Exit Sub
                End If

                'データが1ページのMAX件数以下ならMAXPageに1を設定
                If Delivery_Prt_List.Length <= Max Then
                    MaxPage = 1
                Else
                    'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
                    If Delivery_Prt_List.Length Mod Max = 0 Then
                        MaxPage = Delivery_Prt_List.Length \ Max
                    Else
                        MaxPage = Delivery_Prt_List.Length \ Max + 1
                    End If
                End If

                DataCount = 0
                For Page = 1 To MaxPage
                    'Headerの設定
                    AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "納品先別出荷リスト（" & Delivery_Prt_List(0).FILE_NAME & "）"
                    '2ページ目以降は納品先名と数量は表示しない。
                    If Page = 1 Then
                        '納品先名の設定
                        AxReport1.Item("", "C_NAME").Text = "納品先：" & Delivery_Prt_List(0).D_NAME

                        '納品先計を設定
                        AxReport1.Item("", "total").Text = TotalNum

                        '御中 納品先計：の表示も削除するように設定
                        AxReport1.Item("", "Label1").Text = "御中　納品先計："
                    Else
                        '納品先名の設定
                        AxReport1.Item("", "C_NAME").Text = ""

                        '納品先計を設定
                        AxReport1.Item("", "total").Text = ""

                        '御中 納品先計：の表示も削除するように設定
                        AxReport1.Item("", "Label1").Text = ""
                    End If
                    LineCount = 1
                    'Max件数以下なら、データ件数分ループするように設定
                    If Delivery_Prt_List.Length <= Max Then
                        LoopCount = Delivery_Prt_List.Length
                    Else
                        'Max件数以上ならMax値までループを設定するが
                        '最終ページはデータ件数までの値を設定する。
                        If Page = MaxPage Then
                            LoopCount = Delivery_Prt_List.Length - ((Page - 1) * Max)
                        Else
                            LoopCount = Max
                        End If
                    End If

                    'ページが変わったら、明細部分をクリアする。
                    For i = 1 To Max
                        '製品コードの表示
                        AxReport1.Item("", "Label" & i & "-1").Text = ""
                        '製品名の表示
                        AxReport1.Item("", "Label" & i & "-2").Text = " "
                        '出荷数の表示
                        AxReport1.Item("", "Label" & i & "-4").Text = " "
                    Next

                    For i = 1 To LoopCount
                        '製品コードの表示
                        AxReport1.Item("", "Label" & LineCount & "-1").Text = Delivery_Prt_List(DataCount).I_CODE
                        '製品名の表示
                        AxReport1.Item("", "Label" & LineCount & "-2").Text = " " & Delivery_Prt_List(DataCount).I_NAME
                        '出荷数の表示
                        AxReport1.Item("", "Label" & LineCount & "-4").Text = Delivery_Prt_List(DataCount).NUM
                        LineCount += 1
                        DataCount += 1
                    Next

                    'Footerの設定
                    AxReport1.Item("", "footer").Text = PrtPage
                    PrtPage += 1

                    'レポートの印刷 
                    If AxReport1.PrintReport() = False Then
                        'エラー処理
                        MsgBox("印刷時にエラーが発生しました。")
                    End If
                Next
            Next

            '印刷ＪＯＢの終了（ファイルを閉じる） 
            AxReport1.ClosePrintJob(True)

            'レポーを閉じる 
            AxReport1.ReportPath = ""
        Else
            MsgBox("指定の納品先がない為、ファイルが作成されませんでした。")

        End If

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim Check_Flg As Boolean = False
        Dim Check_Data_Count As Integer = 0
        Dim CheckList() As CheckID_List = Nothing
        Dim Check_List() As CheckID_List = Nothing
        Dim Check_Data_List() As Check_List = Nothing

        Dim KansekiMaxData = 6
        Dim TakamiyaMaxData = 10
        Dim PfjMaxData = 10

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 44
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0

        Dim Count As Integer = 0
        Dim Check_Result As Boolean = True
        Dim Check_ErrorMessage As String = Nothing
        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0
        Dim Total As Integer = 0
        Dim FILE_NAME As String = Nothing
        Dim ALLCheck_Flg As Boolean = True
        Dim Out_Check_Flg As Boolean = True

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("納品書チェックリストを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        Out_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("伝票出力のみのデータにチェックされています。伝票出力のみのデータは納品書チェックリストの出力は行えません。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("納品書チェックリストを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        FILE_NAME = Trim(TextBox3.Text)

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

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "PackingSlipCheckList.crp"
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
        If AxReport1.OpenPrintJob("PackingSlipCheckList.crp", 512, -1, "納品書チェックリスト　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            Exit Sub
        End If

        'カンセキ伝票の納品先データを取得
        Check_Result = GetDliveryCheckData(FILE_NAME, 1, KansekiMaxData, Check_Data_List, Total, Check_Result, Check_ErrorMessage)
        If Check_Result = False Then
            MsgBox(Check_ErrorMessage)
            Exit Sub
        End If

        'カンセキ伝票分の納品書チェックリストの作成

        'データがなければ、印刷しない。
        If Total <> 0 Then

            'Headerの設定
            AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "納品先別出荷リスト（" & Check_Data_List(0).FILE_NAME & "）"

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Check_Data_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                If Check_Data_List.Length Mod Max = 0 Then
                    MaxPage = Check_Data_List.Length \ Max
                Else
                    MaxPage = Check_Data_List.Length \ Max + 1
                End If
            End If

            For Page = 1 To MaxPage
                '2ページ目以降は納品先名と数量は表示しない。
                If Page = 1 Then
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = "　伝票タイプ："
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = "チェーンストア統一伝票Ⅰ型"
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total
                Else
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = ""
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = ""
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total
                End If

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Check_Data_List.Length <= Max Then
                    LoopCount = Check_Data_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Check_Data_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & i & "-1").Text = ""
                    '納品先名の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = ""
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & i & "-3").Text = ""
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = ""
                Next

                For i = 1 To LoopCount
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & LineCount & "-1").Text = Check_Data_List(DataCount).C_CODE
                    '納品先名の表示
                    AxReport1.Item("", "Label" & LineCount & "-2").Text = " " & Check_Data_List(DataCount).D_NAME
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & LineCount & "-3").Text = Check_Data_List(DataCount).SHEET_NO
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & LineCount & "-4").Text = Check_Data_List(DataCount).DATA_NUM
                    LineCount += 1
                    DataCount += 1
                Next

                'Footerの設定
                AxReport1.Item("", "footer").Text = Page & "/ " & MaxPage & "　ページ"

                'レポートの印刷 
                If AxReport1.PrintReport() = False Then
                    'エラー処理
                    MsgBox("印刷時にエラーが発生しました。")
                End If
            Next
        End If
        DataCount = 0
        Total = 0

        Check_Data_List = Nothing

        'PFJ伝票の納品先データを取得
        Check_Result = GetDliveryCheckData(FILE_NAME, 2, PfjMaxData, Check_Data_List, Total, Check_Result, Check_ErrorMessage)
        If Check_Result = False Then
            MsgBox(Check_ErrorMessage)
            Exit Sub
        End If
        'データがなければ、印刷しない。
        If Total <> 0 Then

            'Headerの設定
            AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "納品先別出荷リスト（" & Check_Data_List(0).FILE_NAME & "）"

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Check_Data_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                If Check_Data_List.Length Mod Max = 0 Then
                    MaxPage = Check_Data_List.Length \ Max
                Else
                    MaxPage = Check_Data_List.Length \ Max + 1
                End If
            End If

            For Page = 1 To MaxPage
                '2ページ目以降は納品先名と数量は表示しない。
                If Page = 1 Then
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = "　伝票タイプ："
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = "チェーンストア統一伝票Ⅲ型"
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total
                Else
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = ""
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = ""
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total

                End If

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Check_Data_List.Length <= Max Then
                    LoopCount = Check_Data_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Check_Data_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & i & "-1").Text = ""
                    '納品先名の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = ""
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & i & "-3").Text = ""
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = ""
                Next

                For i = 1 To LoopCount
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & LineCount & "-1").Text = Check_Data_List(DataCount).C_CODE
                    '納品先名の表示
                    AxReport1.Item("", "Label" & LineCount & "-2").Text = " " & Check_Data_List(DataCount).D_NAME
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & LineCount & "-3").Text = Check_Data_List(DataCount).SHEET_NO
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & LineCount & "-4").Text = Check_Data_List(DataCount).DATA_NUM
                    LineCount += 1
                    DataCount += 1
                Next

                'Footerの設定
                AxReport1.Item("", "footer").Text = Page & "/ " & MaxPage & "　ページ"

                'レポートの印刷 
                If AxReport1.PrintReport() = False Then
                    'エラー処理
                    MsgBox("印刷時にエラーが発生しました。")
                End If
            Next
        End If

        DataCount = 0
        Total = 0
        Check_Data_List = Nothing

        'タカミヤ伝票の納品先データを取得
        Check_Result = GetDliveryCheckData(FILE_NAME, 3, TakamiyaMaxData, Check_Data_List, Total, Check_Result, Check_ErrorMessage)
        If Check_Result = False Then
            MsgBox(Check_ErrorMessage)
            Exit Sub
        End If

        'データがなければ、印刷しない。
        If Total <> 0 Then
            'Headerの設定
            AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "納品先別出荷リスト（" & Check_Data_List(0).FILE_NAME & "）"

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Check_Data_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、データ件数 \ MAX件数で商を求め、Modで余りが出るなら+1
                If Check_Data_List.Length Mod Max = 0 Then
                    MaxPage = Check_Data_List.Length \ Max
                Else
                    MaxPage = Check_Data_List.Length \ Max + 1
                End If
            End If

            For Page = 1 To MaxPage
                '2ページ目以降は納品先名と数量は表示しない。
                If Page = 1 Then
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = "　伝票タイプ："
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = "タカミヤ伝票"
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total
                Else
                    '伝票タイプのタイトル
                    AxReport1.Item("", "C_NAME").Text = ""
                    '伝票タイプの設定
                    AxReport1.Item("", "SHEET_TYPE").Text = ""
                    '合計の設定
                    AxReport1.Item("", "total").Text = Total
                End If

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Check_Data_List.Length <= Max Then
                    LoopCount = Check_Data_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Check_Data_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & i & "-1").Text = ""
                    '納品先名の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = ""
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & i & "-3").Text = ""
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = ""
                Next

                For i = 1 To LoopCount
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & LineCount & "-1").Text = Check_Data_List(DataCount).C_CODE
                    '納品先名の表示
                    AxReport1.Item("", "Label" & LineCount & "-2").Text = " " & Check_Data_List(DataCount).D_NAME
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & LineCount & "-3").Text = Check_Data_List(DataCount).SHEET_NO
                    '伝票総枚数の表示
                    AxReport1.Item("", "Label" & LineCount & "-4").Text = Check_Data_List(DataCount).DATA_NUM
                    LineCount += 1
                    DataCount += 1
                Next

                'Footerの設定
                AxReport1.Item("", "footer").Text = Page & "/ " & MaxPage & "　ページ"

                'レポートの印刷 
                If AxReport1.PrintReport() = False Then
                    'エラー処理
                    MsgBox("印刷時にエラーが発生しました。")
                End If

            Next
        End If

        '印刷ＪＯＢの終了（ファイルを閉じる） 
        AxReport1.ClosePrintJob(True)

        'レポーを閉じる 
        AxReport1.ReportPath = ""

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim CommentCheck_Flg As Boolean = False
        Dim CommentCheck_Data_Count As Integer = 0
        Dim CommentCheck_List() As CommentID_List = Nothing
        Dim Comment_List() As Comment_List = Nothing
        Dim Comment_Result As Boolean = True
        Dim Comment_ErrorMessage As String = Nothing
        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0
        Dim CommentDataCount As Integer = 0

        'データカウント用変数
        Dim DataCount As Integer = 0

        '1ページの最大行数
        Dim Max As Integer = 26
        '印刷枚数
        Dim MaxPage As Integer = 0

        '出荷指示ファイル名
        Dim FILE_NAME As String = Nothing

        Dim ALLCheck_Flg As Boolean = True
        Dim Out_Check_Flg As Boolean = True

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        Out_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("伝票出力のみのデータにチェックされています。伝票出力のみのデータはコメントリストリストの出力は行えません。")
            Exit Sub
        End If

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("コメントリストを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("コメントリストを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        FILE_NAME = Trim(TextBox3.Text)



        Comment_Result = GetCommentData(FILE_NAME, PLACE_ID, Comment_List, CommentDataCount, Comment_Result, Comment_ErrorMessage)
        If Comment_Result = False Then
            MsgBox("Comment_ErrorMessage")
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If CommentDataCount <> 0 Then

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

            'レポートを開く
            AxReport1.ReportPath = PrtForm & "CommentList.crp"
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
            If AxReport1.OpenPrintJob("CommentList.crp", 8, 512, "コメントリスト　プレビュー", 0) = False Then
                'エラー処理を記述します 
                MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                Exit Sub
            End If

            'Headerの設定
            AxReport1.Item("", "header").Text = dtNow.ToString("yyyy/MM/dd") & "  " & "コメントリスト（" & Comment_List(0).FILE_NAME & "）"

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Comment_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
                If Comment_List.Length Mod Max = 0 Then
                    MaxPage = Comment_List.Length \ Max
                Else
                    MaxPage = Comment_List.Length \ Max + 1
                End If
            End If

            'ページ数分ループ
            For Page = 1 To MaxPage

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Comment_List.Length <= Max Then
                    LoopCount = Comment_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Comment_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & i & "-1").Text = ""
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = ""
                    '製品コードの表示
                    AxReport1.Item("", "Label" & i & "-3").Text = ""
                    '出荷数の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = ""
                    'コメント１の表示
                    AxReport1.Item("", "Label" & i & "-5").Text = ""
                    'コメント２の表示
                    AxReport1.Item("", "Label" & i & "-6").Text = ""
                Next

                For i = 1 To LoopCount
                    '納品先IDの表示
                    AxReport1.Item("", "Label" & i & "-1").Text = Comment_List(DataCount).C_CODE
                    '伝票番号の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = Comment_List(DataCount).SHEET_NO
                    '製品コードの表示
                    AxReport1.Item("", "Label" & i & "-3").Text = Comment_List(DataCount).I_CODE
                    '出荷数の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = Comment_List(DataCount).NUM

                    If i <> 1 Then
                        If Comment_List(DataCount).C_CODE = Comment_List(DataCount - 1).C_CODE And Comment_List(DataCount).SHEET_NO = Comment_List(DataCount - 1).SHEET_NO Then
                            '二件目以降、１つ上のデータとチェックし、同じ納品先、同じ伝票番号ならコメント1,2は表示しない。
                            AxReport1.Item("", "Label" & i & "-5").Text = " "
                            AxReport1.Item("", "Label" & i & "-6").Text = " "
                        Else
                            'コメント１の表示
                            AxReport1.Item("", "Label" & i & "-5").Text = " " & Comment_List(DataCount).COMMENT1
                            'コメント２の表示
                            AxReport1.Item("", "Label" & i & "-6").Text = " " & Comment_List(DataCount).COMMENT2
                        End If
                    Else
                        'コメント１の表示
                        AxReport1.Item("", "Label" & i & "-5").Text = " " & Comment_List(DataCount).COMMENT1
                        'コメント２の表示
                        AxReport1.Item("", "Label" & i & "-6").Text = " " & Comment_List(DataCount).COMMENT2
                    End If
                    LineCount += 1
                    DataCount += 1
                Next

                'Footerの設定
                AxReport1.Item("", "footer").Text = Page & "/ " & MaxPage & "　ページ"

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
        Else
            MsgBox("コメントのあるデータが無い為、コメントリストの作成は行いませんでした。")
        End If

    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim CustomerName As String = Nothing
        Dim C_ID As Integer = 0
        Dim C_Code As String = Nothing
        Dim ChkCustomerCodeString As String = Nothing

        '商品名欄をクリアする。
        Label18.Text = ""

        'Enterキーが押されているか確認
        If e.KeyCode = Keys.Enter Then
            '入力チェック
            If Trim(TextBox4.Text) <> "" Then
                ChkCustomerCodeString = Trim(TextBox4.Text)
                'アイテムコードに'が入力されていたらReplaceする。
                ChkCustomerCodeString = ChkCustomerCodeString.Replace("'", "''")

                Me.ProcessTabKey(Not e.Shift)
                e.Handled = True
                '入力された商品コードを元に商品名を取得する。
                'ログインチェックFunction
                Result = GetCustomerName(ChkCustomerCodeString, 1, CustomerName, C_ID, C_Code, Result, ErrorMessage)
                If Result = "True" Then
                    Label18.Text = CustomerName
                    TextBox4.BackColor = Color.White
                ElseIf Result = "False" Then
                    MsgBox(ErrorMessage)
                    TextBox4.Focus()
                    TextBox4.BackColor = Color.Salmon
                    'エラーの場合、商品名もクリア。
                    Label7.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker2.Value.ToShortDateString()

        MaskedTextBox2.Text = Date.ParseExact(DateTimePicker2.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim ToDate As Date

        ToDate = DateTimePicker1.Value.ToShortDateString()

        MaskedTextBox1.Text = Date.ParseExact(DateTimePicker1.Value.ToShortDateString(), "yyyy/MM/dd", Nothing)

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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim IndexCount = 22

        Dim sample() As Integer = Nothing
        Dim SelectCount As Integer = 0
        Dim LoopCount As Integer = 0

        For Each Row As DataGridViewRow In DataGridView1.SelectedRows
            ReDim Preserve sample(0 To SelectCount)
            sample(SelectCount) = Row.Index
            SelectCount += 1
        Next Row

        For LoopCount = 0 To sample.Length - 1
            'MsgBox(sample(LoopCount))

            If DataGridView1(0, sample(LoopCount)).Value = 1 Then
                DataGridView1(0, sample(LoopCount)).Value = 0
                For Count = 0 To IndexCount
                    DataGridView1.Item(Count, sample(LoopCount)).Style.BackColor = Color.White
                Next
            ElseIf DataGridView1(0, sample(LoopCount)).Value = 0 Then
                DataGridView1(0, sample(LoopCount)).Value = 1
                For Count = 0 To IndexCount
                    DataGridView1.Item(Count, sample(LoopCount)).Style.BackColor = Color.PaleGreen
                Next
            End If
        Next
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount = 24

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
                    If CheckBox9.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.White
                        Next
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox9.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = 1
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.PaleGreen
                        Next
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As Out_Search_List = Nothing
        Dim Count As Integer = 0

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Fix_Date_From As String = Nothing
        Dim Fix_Date_To As String = Nothing

        Dim Date_Check_Result As DateTime

        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim FileNameCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing

        Dim ItemJan_Flg As Integer = 0

        Dim ChkSheetNoString As String = Nothing
        Dim ChkOrderNoString As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ChkCommentString As String = Nothing
        Dim ChkCCodeString As String = Nothing

        Dim ChkCLAIMCodeString As String = Nothing

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

        'DataGridViewをクリアする。
        'DataGridView1.Rows.Clear()
        '検索ボタンクリックチェック
        sSearchFLg = False

        '一括チェックのチェックボックスのクリア
        CheckBox9.Checked = False

        '出荷指示ファイル名称の文字列妥当性チェック（SQLを実行するための'の対策）
        If Trim(TextBox3.Text) <> "" Then
            '文字の妥当性チェックを行う。
            'NullOK、
            If StringChkVal(Trim(TextBox3.Text), True, False, FileNameCheckAfter, StringChkResult, StringErrorMessage) = False Then
                TextBox3.BackColor = Color.Salmon
                MsgBox("出荷指示ファイル名に不正な文字が入力されています。")
            Else
                'チェックに問題がなければ背景色を白に戻す。
                TextBox3.BackColor = Color.White
            End If
        End If

        '出荷予定日Fromのチェック
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
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '出荷予定日Toのチェック
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
                MsgBox("出荷予定日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '出荷日Fromのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox3.BackColor = Color.Salmon
                MaskedTextBox3.Focus()
                Exit Sub
            End If
        End If

        '出荷日Toのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox4.BackColor = Color.Salmon
                MaskedTextBox4.Focus()
                Exit Sub
            End If
        End If

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox1.Text = "商品コード" Then
            ItemJan_Flg = 1
            ComboBox1.BackColor = Color.White
        ElseIf ComboBox1.Text = "JANコード" Then
            ItemJan_Flg = 2
            ComboBox1.BackColor = Color.White
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox1.BackColor = Color.Salmon
            ComboBox1.Focus()
            Exit Sub
        End If

        ChkSheetNoString = Trim(TextBox1.Text)
        ChkOrderNoString = Trim(TextBox2.Text)
        ChkItemJanString = Trim(TextBox7.Text)
        ChkCommentString = Trim(TextBox8.Text)
        ChkCCodeString = Trim(TextBox4.Text)
        ChkCLAIMCodeString = Trim(TextBox5.Text)

        '伝票番号に'が入力されたいたらReplaceする。
        ChkSheetNoString = ChkSheetNoString.Replace("'", "''")

        'オーダー番号に'が入力されていたらReplaceする。
        ChkOrderNoString = ChkOrderNoString.Replace("'", "''")

        'アイテムコードorJANコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        'アイテムコードorJANコードに'が入力されていたらReplaceする。
        ChkCommentString = ChkCommentString.Replace("'", "''")

        '納品先コードに'が入っていたらReplaceする。
        ChkCCodeString = ChkCCodeString.Replace("'", "''")

        '出荷予定検索Function
        Result = GetOutSearch(FileNameCheckAfter, ChkSheetNoString, ChkOrderNoString, Date_From, Date_To, Fix_Date_From, Fix_Date_To, _
                            ItemJan_Flg, ChkItemJanString, ChkCommentString, ChkCCodeString, ChkCLAIMCodeString, CheckBox3.Checked, _
                            CheckBox11.Checked, CheckBox10.Checked, CheckBox4.Checked, CheckBox7.Checked, CheckBox5.Checked, _
                            CheckBox6.Checked, CheckBox1.Checked, CheckBox2.Checked, PLACE_ID, _
                            SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            Label14.Text = "商品数："
            Label15.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "出庫関連データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "伝票番号,オーダー番号,商品コード,商品名,出荷指示ファイル名,納品先コード,納品先名,予定数量,出荷予定日,出荷数量,出荷日,ステータス,出荷種別,商品ステータス,印刷日,納入単価,売単価,コメント１,コメント２,備考,倉庫,郵便番号,住所"

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

        Dim Item_Name As String = Nothing

        'データ件数分ループ
        For i = 0 To SearchResult.Length - 1

            '取得するデータをLineataに格納する。
            'A列に伝票番号
            LineData = """" & SearchResult(i).SHEET_NO & ""","
            'B列にオーダー番号
            LineData &= """" & SearchResult(i).ORDER_NO & ""","
            'C列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'D列に商品名
            'LineData &= """" & SearchResult(i).I_NAME & ""","
            Item_Name = SearchResult(i).I_NAME
            LineData &= """" & Item_Name.Replace("""", ChrW(34) & ChrW(34)) & ""","
            'E列に出荷指示ファイル名
            LineData &= """" & SearchResult(i).FILE_NAME & ""","
            'F列に納品先コード
            LineData &= """" & SearchResult(i).C_CODE & ""","
            'G列に納品先名
            LineData &= """" & SearchResult(i).C_NAME & ""","
            'H列に予定数量
            LineData &= """" & SearchResult(i).NUM & ""","
            'I列に出荷予定日
            LineData &= """" & SearchResult(i).O_DATE & ""","
            'J列に出荷数量
            LineData &= """" & SearchResult(i).FIX_NUM & ""","
            'K列に出荷日
            LineData &= """" & SearchResult(i).FIX_DATE & ""","

            'L列にステータス
            LineData &= """" & SearchResult(i).STATUS & ""","
            'M列に出荷種別
            LineData &= """" & SearchResult(i).CATEGORY & ""","
            'N列に商品ステータス
            LineData &= """" & SearchResult(i).DEFECT_TYPE & ""","
            'O列に印刷日
            LineData &= """" & SearchResult(i).PRT_DATE & ""","
            'P列に納入単価
            LineData &= """" & SearchResult(i).COST & ""","
            'Q列に売単価
            LineData &= """" & SearchResult(i).PRICE & ""","
            'R列にコメント１
            LineData &= """" & SearchResult(i).COMMENT1 & ""","
            'S列にコメント２
            LineData &= """" & SearchResult(i).COMMENT2 & ""","
            'T列に備考
            LineData &= """" & SearchResult(i).REMARKS & ""","
            'U列に倉庫
            LineData &= """" & SearchResult(i).PLACE & ""","
            'V列に備考
            LineData &= """" & SearchResult(i).D_ZIP & ""","
            'W列に倉庫
            LineData &= """" & SearchResult(i).D_ADDRESS & """"

            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing

        Dim Picking_Check_Flg As Boolean = False
        Dim Picking_Data_Count As Integer = 0

        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 84
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0

        '現在のデータ件数
        Dim NowCount As Integer = 0


        '出荷指示ファイル名
        Dim FILE_NAME As String = Nothing

        Dim ALLCheck_Flg As Boolean = True

        Dim Label_Prt_List() As Label_Prt_List = Nothing

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        Dim ReturnDataCount As Integer = 0
        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        Picking_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷予定" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" OrElse DataGridView1.Rows(Count).Cells(12).Value() = "出荷済み" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Picking_Check_Flg = False
            End If
        Next
        If Picking_Check_Flg = False Then
            MsgBox("出荷予定、ピッキング戻し、出荷済み、伝票出力のみのデータにチェックされています。ピッキング済みのデータのみラベルの出力が行えます。")
            Exit Sub
        End If

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("ラベルを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("ラベルを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        FILE_NAME = Trim(TextBox3.Text)


        Result = GetLabelPrtData(FILE_NAME, Label_Prt_List, ReturnDataCount, Result, ErrorMessage)

        If ReturnDataCount <> 0 Then

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

            'レポートを開く
            AxReport1.ReportPath = PrtForm & "LabelList.crp"
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
            If AxReport1.OpenPrintJob("LabelList.crp", 512, -1, "ラベルリスト　プレビュー", 0) = False Then
                'エラー処理を記述します 
                MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                Exit Sub
            End If

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Label_Prt_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
                If Label_Prt_List.Length Mod Max = 0 Then
                    MaxPage = Label_Prt_List.Length \ Max
                Else
                    MaxPage = Label_Prt_List.Length \ Max + 1
                End If
            End If

            'ページ数分ループ
            For Page = 1 To MaxPage

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Label_Prt_List.Length <= Max Then
                    LoopCount = Label_Prt_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Label_Prt_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '表示欄をクリア
                    AxReport1.Item("", "Label" & i).Text = ""

                Next

                For i = 1 To LoopCount
                    '表示欄に設定
                    'AxReport1.Item("", "Label" & i).Text = Label_Prt_List(DataCount).SHEET_NO & vbCr & Label_Prt_List(DataCount).I_CODE & " × " & Label_Prt_List(DataCount).NUM
                    AxReport1.Item("", "Label" & i).Text = Label_Prt_List(DataCount).I_CODE & " × " & Label_Prt_List(DataCount).NUM

                    LineCount += 1
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
        Else
            MsgBox("ラベルデータが無い為、ラベルリストの作成は行いませんでした。")
        End If

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim Delivery_Flg As Boolean = False
        Dim Delivery_Data_Count As Integer = 0
        Dim DeliveryList() As CheckID_List = Nothing
        Dim Delivery_List() As CheckID_List = Nothing
        Dim Delivery_Data_List() As Delivery_HeaderList = Nothing
        Dim Delivery_Data_Detail_List() As Delivery_DetailList = Nothing

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 15
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0

        Dim Count As Integer = 0
        Dim Check_Result As Boolean = True
        Dim Check_ErrorMessage As String = Nothing

        Dim Delivery_Result As Boolean = True
        Dim Delivery_ErrorMessage As String = Nothing

        Dim LineCount As Integer = 0
        Dim LoopCount As Integer = 0

        Dim TotalAmount As Integer = 0
        Dim C_ID As Integer
        Dim SHEET_NO As String = Nothing

        Dim FILE_NAME As String = Nothing
        Dim ALLCheck_Flg As Boolean = True
        Dim Out_Check_Flg As Boolean = True

        Dim DetailCount As Integer = 0
        Dim TotalCount As Integer = 0

        Dim SubTotal As Integer = 0

        Dim Delivery_Check_Flg As Boolean = False

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'DataGridViewに表示された全てのデータにチェックが入っているか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() <> 1 Then
                ALLCheck_Flg = False
                MsgBox("帳票を出力するには全てのデータにチェックが入っている必要があります。")
                Exit Sub
            End If
        Next

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox3.Text) = "" Then
            MsgBox("納品書を出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        Delivery_Check_Flg = True
        'チェックされた商品の中で出荷予定、ピッキング済以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso ( _
               DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Delivery_Check_Flg = False
            End If
        Next
        If Delivery_Check_Flg = False Then
            MsgBox("ステータスが出荷予定、ピッキング済み、出荷済みのみ、納品書の出力が行えます。")
            Exit Sub
        End If

        Out_Check_Flg = True
        'チェックされた商品の中で出荷予定以外のデータがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ" Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("伝票出力のみのデータにチェックされています。伝票出力のみのデータは納品書の出力は行えません。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(5).Value() Then
                MsgBox("納品書を出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
                Exit Sub
            End If
        Next

        FILE_NAME = Trim(TextBox3.Text)

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

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "DeliveryList.crp"
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
        If AxReport1.OpenPrintJob("DeliveryList.crp", 512, -1, "納品書　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            Exit Sub
        End If

        '納品先データを取得
        Check_Result = GetDliveryData(FILE_NAME, Delivery_Data_List, Check_Result, Check_ErrorMessage)
        If Check_Result = False Then
            MsgBox(Check_ErrorMessage)
            Exit Sub
        End If

        '納品先の一覧を取得できたので、件数分ループし、納品先別ごとの商品情報を取得。
        For Count = 0 To Delivery_Data_List.Length - 1
            TotalAmount = 0
            C_ID = Delivery_Data_List(Count).C_ID
            SHEET_NO = Delivery_Data_List(Count).SHEET_NO
            Delivery_Data_Detail_List = Nothing

            '詳細データを取得

            Delivery_Result = GetDliveryDetailData(FILE_NAME, C_ID, SHEET_NO, Delivery_Data_Detail_List, TotalAmount, Delivery_Result, Delivery_ErrorMessage)
            If Delivery_Result = False Then
                MsgBox(Delivery_ErrorMessage)
                Exit Sub
            End If

            'データが1ページのMAX件数以下ならMAXPageに1を設定
            If Delivery_Data_Detail_List.Length <= Max Then
                MaxPage = 1
            Else
                'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
                If Delivery_Data_Detail_List.Length Mod Max = 0 Then
                    MaxPage = Delivery_Data_Detail_List.Length \ Max
                Else
                    MaxPage = Delivery_Data_Detail_List.Length \ Max + 1
                End If
            End If

            DataCount = 0
            For Page = 1 To MaxPage
                'Headerの設定

                'Page No
                AxReport1.Item("", "PageNo").Text = Page & " "
                '発送日
                AxReport1.Item("", "Shipping_Date").Text = DateTime.Today & " "
                '納品番号
                AxReport1.Item("", "Delivery_No").Text = Delivery_Data_List(Count).SHEET_NO & " "

                'LEGITデザイン企業情報
                '企業名
                AxReport1.Item("", "CompanyName").Text = Com_NAME
                '企業郵便番号
                AxReport1.Item("", "CompanyPost").Text = Com_POST
                '企業住所
                AxReport1.Item("", "CompanyAddress").Text = Com_ADDRESS
                '企業TEL
                AxReport1.Item("", "CompanyTel").Text = Com_TEL
                '企業FAX
                AxReport1.Item("", "CompanyFax").Text = Com_FAX


                '企業名
                AxReport1.Item("", "CustomerName").Text = Delivery_Data_List(Count).C_NAME
                '納品番号
                AxReport1.Item("", "C_Delivery_No").Text = Delivery_Data_List(Count).C_CODE & " "
                '納品先名
                AxReport1.Item("", "Delivery_Name").Text = Delivery_Data_List(Count).D_NAME

                '合計金額
                AxReport1.Item("", "Header_TotalAmount").Text = Format(TotalAmount, "#,#")

                'コメントは空の場合があるのでクリア
                AxReport1.Item("", "Comment1").Text = ""
                AxReport1.Item("", "Comment2").Text = ""

                LineCount = 1
                'Max件数以下なら、データ件数分ループするように設定
                If Delivery_Data_Detail_List.Length <= Max Then
                    LoopCount = Delivery_Data_Detail_List.Length
                Else
                    'Max件数以上ならMax値までループを設定するが
                    '最終ページはデータ件数までの値を設定する。
                    If Page = MaxPage Then
                        LoopCount = Delivery_Data_Detail_List.Length - ((Page - 1) * Max)
                    Else
                        LoopCount = Max
                    End If
                End If

                SubTotal = 0
                'ページが変わったら、明細部分をクリアする。
                For i = 1 To Max
                    '品番の表示
                    AxReport1.Item("", "Label" & i & "-1").Text = " "
                    '品名の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = " "
                    'JANの表示
                    AxReport1.Item("", "Label" & i & "-3").Text = " "
                    '数量の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = " "
                    '納入単価の表示
                    AxReport1.Item("", "Label" & i & "-5").Text = " "
                    '金額の表示
                    AxReport1.Item("", "Label" & i & "-6").Text = " "
                    '参考上代の表示
                    AxReport1.Item("", "Label" & i & "-7").Text = " "
                Next

                For i = 1 To LoopCount
                    '品番の表示
                    AxReport1.Item("", "Label" & i & "-1").Text = " " & Delivery_Data_Detail_List(DataCount).I_CODE
                    '品名の表示
                    AxReport1.Item("", "Label" & i & "-2").Text = " " & Delivery_Data_Detail_List(DataCount).I_NAME
                    'JANの表示
                    AxReport1.Item("", "Label" & i & "-3").Text = " " & Delivery_Data_Detail_List(DataCount).JAN
                    '数量の表示
                    AxReport1.Item("", "Label" & i & "-4").Text = Format(Delivery_Data_Detail_List(DataCount).NUM, "#,#") & " "
                    '納入単価の表示（納入単価が０だったら空白を表示する。）
                    If Delivery_Data_Detail_List(DataCount).UNIT_PRICE = 0 Then
                        AxReport1.Item("", "Label" & i & "-5").Text = " "
                    Else
                        AxReport1.Item("", "Label" & i & "-5").Text = Format(Delivery_Data_Detail_List(DataCount).UNIT_PRICE, "#,#") & " "
                    End If
                    '金額の表示（納入単価が０だったら空白を表示する。）
                    If Delivery_Data_Detail_List(DataCount).UNIT_PRICE = 0 Then
                        AxReport1.Item("", "Label" & i & "-6").Text = " "
                    Else
                        AxReport1.Item("", "Label" & i & "-6").Text = Format(Delivery_Data_Detail_List(DataCount).NUM * Delivery_Data_Detail_List(DataCount).UNIT_PRICE, "#,#") & " "
                    End If
                    '参考上代の表示。参考上代が0だったらオープンと表示する。
                    If Delivery_Data_Detail_List(DataCount).REFERENCE_PRICE = 0 Then
                        '0だったらオープンと表示
                        AxReport1.Item("", "Label" & i & "-7").Text = "オープン"
                    Else
                        AxReport1.Item("", "Label" & i & "-7").Text = Format(Delivery_Data_Detail_List(DataCount).REFERENCE_PRICE, "#,#") & " "
                    End If

                    '小計を算出
                    SubTotal += Delivery_Data_Detail_List(DataCount).NUM * Delivery_Data_Detail_List(DataCount).UNIT_PRICE

                    LineCount += 1
                    DataCount += 1
                Next

                'Footerの設定
                '小計
                AxReport1.Item("", "Sub_Total").Text = Format(SubTotal, "#,#")
                'Footerの合計金額
                AxReport1.Item("", "Footer_TotalAmount").Text = Format(TotalAmount, "#,#")
                'コメント１
                AxReport1.Item("", "Comment1").Text = Delivery_Data_List(Count).COMMENT1
                'コメント２
                AxReport1.Item("", "Comment2").Text = Delivery_Data_List(Count).COMMENT2
                'レポートの印刷 
                If AxReport1.PrintReport() = False Then
                    'エラー処理
                    MsgBox("印刷時にエラーが発生しました。")
                End If
            Next
        Next

        '印刷ＪＯＢの終了（ファイルを閉じる） 
        AxReport1.ClosePrintJob(True)

        'レポーを閉じる 
        AxReport1.ReportPath = ""

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim I_CODE As String = Nothing
        Dim I_NAME As String = Nothing
        Dim FIX_NUM As Integer
        Dim OUT_ID As Integer
        Dim FIX_DATE As String = Nothing
        Dim D_NAME As String = Nothing

        Dim Out_Lot_Data() As Lot_List = Nothing
        Dim Out_Data_Count As Integer = 0

        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing
        Dim DataCount As Integer

        Dim Check_Count As Integer = 0

        Dim Check_Row As Integer = 0
        'データが０件ならエラー
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Check_Count += 1
                Check_Row = Count
            End If
        Next

        If Check_Count <> 1 Then
            MsgBox("１つだけチェックをして、ボタンを押してください。")
            Exit Sub
        End If

        Dim Out_Lot_Check As Boolean = True

        Out_Lot_Check = True
        'チェックされた商品の中で出庫済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷予定" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング済み" OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Out_Lot_Check = False
            End If
        Next
        If Out_Lot_Check = False Then
            MsgBox("出荷予定、ピッキング戻し、ピッキング済み、伝票出力のみのデータにチェックされています。出荷済みのデータのみロット番号入力が行えます。")
            Exit Sub
        End If

        '商品コード
        I_CODE = DataGridView1.Rows(Check_Row).Cells(3).Value()
        '商品名
        I_NAME = DataGridView1.Rows(Check_Row).Cells(4).Value()
        '数量
        FIX_NUM = DataGridView1.Rows(Check_Row).Cells(10).Value()
        'OUT.ID
        OUT_ID = DataGridView1.Rows(Check_Row).Cells(22).Value()
        '出荷日
        FIX_DATE = DataGridView1.Rows(Check_Row).Cells(9).Value()
        '納品先
        D_NAME = DataGridView1.Rows(Check_Row).Cells(7).Value()

        'slotのDataGridViewをクリア
        slot.DataGridView1.Rows.Clear()

        slot.Label4.Text = I_CODE
        slot.Label5.Text = I_NAME
        slot.Label6.Text = OUT_ID
        slot.Label8.Text = FIX_DATE
        slot.Label10.Text = D_NAME

        'もし、out_lot_managementにすでにデータがあれば取得し表示する。
        'プロダクトライン情報を取得する。
        Result = GetLOTList(OUT_ID, Out_Lot_Data, DataCount, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '数量分のDataGridViewに表示する。
        For Count = 0 To FIX_NUM - 1
            Dim Out_Lot_Data_list As New DataGridViewRow
            Out_Lot_Data_list.CreateCells(slot.DataGridView1)
            With Out_Lot_Data_list
                'No
                .Cells(0).Value = Count + 1
                'ロット番号
                If DataCount <> 0 Then
                    .Cells(1).Value = Out_Lot_Data(Count).LOT_NUMBER
                Else
                    .Cells(1).Value = ""
                End If

                '保証書番号
                If DataCount <> 0 Then
                    .Cells(2).Value = Out_Lot_Data(Count).WARRANTY_CARD_NUMBER
                Else
                    .Cells(2).Value = ""
                End If

            End With
            slot.DataGridView1.Rows.Add(Out_Lot_Data_list)
        Next

        slot.Show()
        Me.Hide()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim SearchResult() As Out_Search_List = Nothing
        Dim Count As Integer = 0
        Dim OutLOT_Check() As Out_Search_List = Nothing
        Dim OutLOT_Count As Integer = 0

        Dim Data_Count As Integer = 0

        Dim FileNameCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing

        Dim ItemJan_Flg As Integer = 0

        Dim ChkSheetNoString As String = Nothing
        Dim ChkOrderNoString As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ChkCommentString As String = Nothing
        Dim ChkCCodeString As String = Nothing

        Dim ChkCLAIMCodeString As String = Nothing

        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        Dim Check_Flg As Boolean = False
        Dim Out_Check_Flg As Boolean = False


        If sSearchFLg = True Then
            MsgBox("出荷確定、出荷予定変更を行った後は再度検索を行って下さい。")
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
                Check_Flg = True
            End If
        Next

        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Out_Check_Flg = True
        'チェックされた商品の中でピッキング済み以外がチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 AndAlso (DataGridView1.Rows(Count).Cells(12).Value() = "出荷予定" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング済み" OrElse DataGridView1.Rows(Count).Cells(12).Value() = "ピッキング戻し" _
               OrElse DataGridView1.Rows(Count).Cells(12).Value() = "伝票出力のみ") Then
                Out_Check_Flg = False
            End If
        Next
        If Out_Check_Flg = False Then
            MsgBox("出荷予定、ピッキング戻し、ピッキング済み、伝票出力のみのデータにチェックされています。出荷済みのデータのみCSV出力が行えます。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve OutLOT_Check(0 To OutLOT_Count)
                'OUT.ID
                OutLOT_Check(OutLOT_Count).ID = DataGridView1.Rows(Count).Cells(22).Value()

                OutLOT_Count += 1
            End If
        Next

        'LOTデータ取得Function
        Result = GetOutLOTSearch(OutLOT_Check, SearchResult, Data_Count, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "ロットデータ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "伝票番号,オーダー番号,商品コード,商品名,出荷指示ファイル名,納品先コード,納品先名,予定数量,出荷予定日,出荷数量,出荷日,ステータス,出荷種別,商品ステータス,印刷日,納入単価,売単価,コメント１,コメント２,備考,倉庫,郵便番号,住所,NO,ロット番号,保証書番号"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If Data_Count = 0 Then
            MsgBox("データがありません。")
            Exit Sub
        End If

        'ﾌｧｲﾙ名の設定
        strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)
        'headerを設定
        strStreamWriter.WriteLine(Sheet1_Header)

        Dim Item_Name As String = Nothing

        'データ件数分ループ
        For i = 0 To SearchResult.Length - 1

            '取得するデータをLineataに格納する。
            'A列に伝票番号
            LineData = """" & SearchResult(i).SHEET_NO & ""","
            'B列にオーダー番号
            LineData &= """" & SearchResult(i).ORDER_NO & ""","
            'C列に商品コード
            LineData &= """" & SearchResult(i).I_CODE & ""","
            'D列に商品名
            'LineData &= """" & SearchResult(i).I_NAME & ""","
            Item_Name = SearchResult(i).I_NAME
            LineData &= """" & Item_Name.Replace("""", ChrW(34) & ChrW(34)) & ""","
            'E列に出荷指示ファイル名
            LineData &= """" & SearchResult(i).FILE_NAME & ""","
            'F列に納品先コード
            LineData &= """" & SearchResult(i).C_CODE & ""","
            'G列に納品先名
            LineData &= """" & SearchResult(i).C_NAME & ""","
            'H列に予定数量
            LineData &= """" & SearchResult(i).NUM & ""","
            'I列に出荷予定日
            LineData &= """" & SearchResult(i).O_DATE & ""","
            'J列に出荷数量
            LineData &= """" & SearchResult(i).FIX_NUM & ""","
            'K列に出荷日
            LineData &= """" & SearchResult(i).FIX_DATE & ""","

            'L列にステータス
            LineData &= """" & SearchResult(i).STATUS & ""","
            'M列に出荷種別
            LineData &= """" & SearchResult(i).CATEGORY & ""","
            'N列に商品ステータス
            LineData &= """" & SearchResult(i).DEFECT_TYPE & ""","
            'O列に印刷日
            LineData &= """" & SearchResult(i).PRT_DATE & ""","
            'P列に納入単価
            LineData &= """" & SearchResult(i).COST & ""","
            'Q列に売単価
            LineData &= """" & SearchResult(i).PRICE & ""","
            'R列にコメント１
            LineData &= """" & SearchResult(i).COMMENT1 & ""","
            'S列にコメント２
            LineData &= """" & SearchResult(i).COMMENT2 & ""","
            'T列に備考
            LineData &= """" & SearchResult(i).REMARKS & ""","
            'U列に倉庫
            LineData &= """" & SearchResult(i).PLACE & ""","
            'V列に備考
            LineData &= """" & SearchResult(i).D_ZIP & ""","
            'W列に倉庫
            LineData &= """" & SearchResult(i).D_ADDRESS & ""","
            'X列にNO
            LineData &= """" & SearchResult(i).NO & ""","
            'Y列にロット番号
            LineData &= """" & SearchResult(i).LOT_NUMBER & ""","
            'Z列に保証書番号
            LineData &= """" & SearchResult(i).WARRANTY_CARD_NUMBER & """"
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)













    End Sub
End Class