Public Class zkensaku

    Dim PL_Id As Integer
    Dim PLACE_ID As String

    Public FormLord As Boolean = False

    'False:検索ボタンを押したらFalse
    '棚卸等行ったら、再度検索を行わせる為のFlg
    Public zSearchFLg As Boolean = False

    '帳票ファイルの格納フォルダ指定
    Public PrtForm As String = System.Configuration.ConfigurationManager.AppSettings("PrtPath")

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '棚卸ボタンが押されたら
        Dim Stock_Check_Flg As Boolean = False
        Dim Stock_Data_Count As Integer = 0
        Dim Stock_Check() As Stock_Tanaoroshi_List = Nothing

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                Stock_Check_Flg = True
            End If
        Next

        If Stock_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Stock_Check(0 To Stock_Data_Count)
                '商品コード
                Stock_Check(Stock_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                Stock_Check(Stock_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                Stock_Check(Stock_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '数量
                Stock_Check(Stock_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                Stock_Check(Stock_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()

                'プロダクトコード名
                Stock_Check(Stock_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                Stock_Check(Stock_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                Stock_Check(Stock_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                Stock_Check(Stock_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                Stock_Check(Stock_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()

                Stock_Data_Count += 1
            End If
        Next

        'ztanaoroshiのDataGridViewをクリア
        ztanaoroshi.DataGridView1.Rows.Clear()

        'ztanaoroshiのDataGridViewに値を入れていく。
        For Count = 0 To Stock_Check.Length - 1
            Dim Tanaoroshi_Data_list As New DataGridViewRow
            Tanaoroshi_Data_list.CreateCells(ztanaoroshi.DataGridView1)
            With Tanaoroshi_Data_list
                '商品コード
                .Cells(0).Value = Stock_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = Stock_Check(Count).I_NAME
                'JAN
                .Cells(2).Value = Stock_Check(Count).JAN
                '数量
                .Cells(3).Value = Stock_Check(Count).NUM
                '棚卸後数量
                .Cells(4).Value = Stock_Check(Count).NUM
                '差分
                .Cells(5).Value = 0
                'ロケーション
                .Cells(7).Value = Stock_Check(Count).LOCATION
                'プロダクトラインコード名
                .Cells(8).Value = Stock_Check(Count).PL_NAME
                '不良区分
                .Cells(9).Value = Stock_Check(Count).I_STATUS
                '倉庫
                .Cells(10).Value = Stock_Check(Count).PLACE

                'STOCK_ID
                .Cells(11).Value = Stock_Check(Count).STOCK_ID
                'I_ID
                .Cells(12).Value = Stock_Check(Count).I_ID

            End With
            ztanaoroshi.DataGridView1.Rows.Add(Tanaoroshi_Data_list)
        Next

        ztanaoroshi.Show()
        Me.Hide()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '在庫調整ボタンが押されたら

        Dim Stock_Check_Flg As Boolean = False
        Dim Stock_Data_Count As Integer = 0
        Dim Stock_Check() As Stock_Tanaoroshi_List = Nothing

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                Stock_Check_Flg = True
            End If
        Next

        If Stock_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Stock_Check(0 To Stock_Data_Count)
                '商品コード
                Stock_Check(Stock_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                Stock_Check(Stock_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                Stock_Check(Stock_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '数量
                Stock_Check(Stock_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                Stock_Check(Stock_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()
                'プロダクトコード名
                Stock_Check(Stock_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                Stock_Check(Stock_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                Stock_Check(Stock_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                Stock_Check(Stock_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                Stock_Check(Stock_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
                'P_ID
                Stock_Check(Stock_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
                Stock_Data_Count += 1
            End If
        Next

        'ztyouseiのDataGridViewをクリア
        ztyousei.DataGridView1.Rows.Clear()

        'ztyouseiのDataGridViewに値を入れていく。
        For Count = 0 To Stock_Check.Length - 1
            Dim Tyousei_Data_list As New DataGridViewRow
            Tyousei_Data_list.CreateCells(ztyousei.DataGridView1)
            With Tyousei_Data_list
                '商品コード
                .Cells(0).Value = Stock_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = Stock_Check(Count).I_NAME
                'JAN
                .Cells(2).Value = Stock_Check(Count).JAN
                '数量
                .Cells(3).Value = Stock_Check(Count).NUM
                'ロケーション
                .Cells(6).Value = Stock_Check(Count).LOCATION
                'プロダクトラインコード名
                .Cells(7).Value = Stock_Check(Count).PL_NAME
                '不良区分
                .Cells(8).Value = Stock_Check(Count).I_STATUS
                '倉庫
                .Cells(9).Value = Stock_Check(Count).PLACE
                'STOCK_ID
                .Cells(10).Value = Stock_Check(Count).STOCK_ID
                'I_ID
                .Cells(11).Value = Stock_Check(Count).I_ID
                'P_ID
                .Cells(12).Value = Stock_Check(Count).PLACE_ID
            End With
            ztyousei.DataGridView1.Rows.Add(Tyousei_Data_list)
        Next

        ztyousei.Show()
        Me.Hide()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'セットばらしボタンが押されたら

        Dim Set_Check_Flg As Boolean = False
        Dim Set_Data_Count As Integer = 0
        Dim Set_Check() As Stock_Tanaoroshi_List = Nothing

        Dim Check_Flg As Boolean = True

        Dim SetCount As Integer = 0
        Dim SingleItem_Check_Flg As Boolean = False

        Dim CheckCount As Integer = 0

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                Set_Check_Flg = True
            End If
        Next

        If Set_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        '不良品がチェックされているか確認。チェックされていたらエラー。
        Check_Flg = True
        'チェックされた商品の中で入荷済みがチェックされていないか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(8).Value() = "不良品" Then
                Check_Flg = False
            End If
        Next
        If Check_Flg = False Then
            MsgBox("不良品にチェックが入っています。不良品はセット組み換えが行えません。")
            Exit Sub
        End If

        SingleItem_Check_Flg = False
        '単体商品がチェックされているか確認。
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(7).Value() = "通常商品" Then
                SingleItem_Check_Flg = True
            End If

            'セット商品が複数チェックされているか確認。
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(7).Value() = "セット商品" Then
                CheckCount += 1
            End If
        Next
        If SingleItem_Check_Flg = True Then
            MsgBox("通常商品にチェックが入っています。" & vbCr & "通常商品は商品をばらすことができません。")
            Exit Sub
        End If
        If CheckCount > 1 Then
            MsgBox("セット商品が複数選択されています。" & vbCr & "セット商品は1件しかチェックできません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Set_Check(0 To Set_Data_Count)
                '商品コード
                Set_Check(Set_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                Set_Check(Set_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                Set_Check(Set_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '数量
                Set_Check(Set_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                Set_Check(Set_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()
                'プロダクトコード名
                Set_Check(Set_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                Set_Check(Set_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                Set_Check(Set_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                Set_Check(Set_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                Set_Check(Set_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
                'P_ID
                Set_Check(Set_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
                Set_Data_Count += 1
            End If
        Next

        'zdismanのDataGridView3をクリア
        zdisman.DataGridView3.Rows.Clear()

        'セット商品->通常商品への組み換えの場合

        'zsetのDataGridViewに値を入れていく。
        For Count = 0 To Set_Check.Length - 1
            Dim Set_Data_list As New DataGridViewRow
            Set_Data_list.CreateCells(zdisman.DataGridView1)
            With Set_Data_list
                '商品コード
                .Cells(0).Value = Set_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = Set_Check(Count).I_NAME
                'JAN
                .Cells(2).Value = Set_Check(Count).JAN
                '数量
                .Cells(3).Value = Set_Check(Count).NUM
                'ロケーション
                .Cells(4).Value = Set_Check(Count).LOCATION
                'プロダクトラインコード名
                .Cells(5).Value = Set_Check(Count).PL_NAME
                '不良区分
                .Cells(6).Value = Set_Check(Count).I_STATUS
                '倉庫
                .Cells(7).Value = Set_Check(Count).PLACE
                'STOCK_ID
                .Cells(8).Value = Set_Check(Count).STOCK_ID
                'I_ID
                .Cells(9).Value = Set_Check(Count).I_ID
                'P_ID
                .Cells(10).Value = Set_Check(Count).PLACE_ID
            End With
            zdisman.DataGridView1.Rows.Add(Set_Data_list)
        Next

        'セットばらし用の商品名欄に選択された商品を表示する。
        zdisman.Label11.Text = Set_Check(0).I_NAME

        Me.Hide()
        zdisman.Show()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '返品出荷ボタンが押されたら

        Dim Return_Check_Flg As Boolean = False
        Dim Return_Data_Count As Integer = 0
        Dim Return_Check() As Stock_Tanaoroshi_List = Nothing

        Dim Check_Flg = True

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                Return_Check_Flg = True
            End If
        Next

        If Return_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'もし良品がチェックされていたらエラーメッセージを表示
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 And DataGridView1.Rows(Count).Cells(8).Value() = "良品" Then
                Check_Flg = False
            End If
        Next
        If Check_Flg = False Then
            MsgBox("良品にチェックが入っています。良品は返品出荷が行えません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Return_Check(0 To Return_Data_Count)
                '商品コード
                Return_Check(Return_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                Return_Check(Return_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                Return_Check(Return_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '数量
                Return_Check(Return_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                Return_Check(Return_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()
                'プロダクトコード名
                Return_Check(Return_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                Return_Check(Return_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                Return_Check(Return_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                Return_Check(Return_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                Return_Check(Return_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
                'P_ID
                Return_Check(Return_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
                Return_Data_Count += 1
            End If
        Next

        'zhenpinのDataGridViewをクリア
        zhenpin.DataGridView1.Rows.Clear()

        'zhenpinのDataGridViewに値を入れていく。
        For Count = 0 To Return_Check.Length - 1
            Dim Return_Data_list As New DataGridViewRow
            Return_Data_list.CreateCells(zhenpin.DataGridView1)
            With Return_Data_list
                '商品コード
                .Cells(0).Value = Return_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = Return_Check(Count).I_NAME
                'JAN
                .Cells(2).Value = Return_Check(Count).JAN
                '数量
                .Cells(3).Value = Return_Check(Count).NUM
                'ロケーション
                .Cells(6).Value = Return_Check(Count).LOCATION
                'プロダクトラインコード名
                .Cells(7).Value = Return_Check(Count).PL_NAME
                '不良区分
                .Cells(8).Value = Return_Check(Count).I_STATUS
                '倉庫
                .Cells(9).Value = Return_Check(Count).PLACE
                'STOCK_ID
                .Cells(10).Value = Return_Check(Count).STOCK_ID
                'I_ID
                .Cells(11).Value = Return_Check(Count).I_ID
                'P_ID
                .Cells(12).Value = Return_Check(Count).PLACE_ID
            End With
            zhenpin.DataGridView1.Rows.Add(Return_Data_list)
        Next

        Me.Hide()
        zhenpin.Show()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '良品変更ボタンが押されたら

        Dim I_Status_Check_Flg As Boolean = False
        Dim I_Status_Data_Count As Integer = 0
        Dim I_Status_Check() As Stock_Tanaoroshi_List = Nothing

        Dim Check_Flg = True

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                I_Status_Check_Flg = True
            End If
        Next

        If I_Status_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve I_Status_Check(0 To I_Status_Data_Count)
                '商品コード
                I_Status_Check(I_Status_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                I_Status_Check(I_Status_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                I_Status_Check(I_Status_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '現在庫数量
                I_Status_Check(I_Status_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                I_Status_Check(I_Status_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()
                'プロダクトコード名
                I_Status_Check(I_Status_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                I_Status_Check(I_Status_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                I_Status_Check(I_Status_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                I_Status_Check(I_Status_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                I_Status_Check(I_Status_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
                'P_ID
                I_Status_Check(I_Status_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
                I_Status_Data_Count += 1
            End If
        Next

        'zkubunのDataGridViewをクリア
        zkubun.DataGridView1.Rows.Clear()

        'zhenpinのDataGridViewに値を入れていく。
        For Count = 0 To I_Status_Check.Length - 1
            Dim Return_Data_list As New DataGridViewRow
            Return_Data_list.CreateCells(zkubun.DataGridView1)
            With Return_Data_list
                '商品コード
                .Cells(0).Value = I_Status_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = I_Status_Check(Count).I_NAME
                'JAN
                .Cells(2).Value = I_Status_Check(Count).JAN
                '数量
                .Cells(3).Value = I_Status_Check(Count).NUM
                '不良区分
                .Cells(5).Value = I_Status_Check(Count).I_STATUS
                '変更後不良区分は不良区分が良品なら不良品を、不良品なら良品を表示。
                '保管品が追加されるので、不良区分と同じものを入れるよう変更。 2015.02.02 H.Mitani
                'If I_Status_Check(Count).I_STATUS = "良品" Then
                '    .Cells(6).Value = "不良品"
                'ElseIf I_Status_Check(Count).I_STATUS = "不良品" Then
                '    .Cells(6).Value = "良品"
                'End If
                .Cells(6).Value = I_Status_Check(Count).I_STATUS

                'ロケーション
                .Cells(7).Value = I_Status_Check(Count).LOCATION
                '変更後ロケーションは、同じものを入れる。
                .Cells(8).Value = I_Status_Check(Count).LOCATION

                'プロダクトラインコード名
                .Cells(9).Value = I_Status_Check(Count).PL_NAME
                '倉庫
                .Cells(10).Value = I_Status_Check(Count).PLACE
                'STOCK_ID
                .Cells(11).Value = I_Status_Check(Count).STOCK_ID
                'I_ID
                .Cells(12).Value = I_Status_Check(Count).I_ID
                'P_ID
                .Cells(13).Value = I_Status_Check(Count).PLACE_ID
            End With
            zkubun.DataGridView1.Rows.Add(Return_Data_list)
        Next
        Me.Hide()
        zkubun.Show()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Dispose()
        topmenu.Show()

    End Sub

    Private Sub zkensaku_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub zkensaku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim ItemCodeResult As Boolean = True
        Dim ItemCodeErrorMessage As String = Nothing
        Dim ItemCodeList() As String = Nothing

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "在庫検索"

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - " & Disp_Title
        Label3.Text = Disp_Title
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        '左3項目を固定(チェック、商品コード、商品名)
        DataGridView1.Columns(2).Frozen = True

        'プロダクトラインコード情報取得
        Dim PLTable As New DataTable()
        PLTable.Columns.Add("ID", GetType(String))
        PLTable.Columns.Add("NAME", GetType(String))
        'プロダクトラインの情報を取得する。
        PLResult = GetPLList(PLList, PLResult, PLErrorMessage)
        If PLResult = False Then
            MsgBox(PLErrorMessage)
            Exit Sub
        End If

        For Count = 0 To PLList.Length - 1
            '新しい行を作成
            Dim row As DataRow = PLTable.NewRow()

            '各列に値をセット
            row("ID") = PLList(Count).ID
            row("NAME") = PLList(Count).NAME

            'DataTableに行を追加
            PLTable.Rows.Add(row)
        Next

        PLTable.AcceptChanges()
        'コンボボックスのDataSourceにDataTableを割り当てる
        ComboBox1.DataSource = PLTable

        '表示される値はDataTableのNAME列
        ComboBox1.DisplayMember = "NAME"
        '対応する値はDataTableのID列
        ComboBox1.ValueMember = "ID"

        '初期値をセット
        ComboBox1.SelectedIndex = -1
        'PL_Id = ComboBox1.SelectedValue.ToString()
        PL_Id = ComboBox1.SelectedValue

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

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        'Dim ErrorMessage As String = Nothing
        'Dim Result As Boolean = True
        'Dim ItemName() As Item_List = Nothing

        'Dim ChkItemCodeString As String = Nothing

        'ChkItemCodeString = Trim(TextBox6.Text)

        ''アイテムコードに'が入力されていたらReplaceする。
        'ChkItemCodeString = ChkItemCodeString.Replace("'", "''")

        ''商品名欄をクリアする。
        'Label7.Text = ""

        ''Enterキーが押されているか確認
        'If e.KeyCode = Keys.Enter Then
        '    '入力チェック
        '    If Trim(TextBox6.Text) <> "" Then
        '        'あたかもTabキーが押されたかのようにする
        '        'Shiftが押されている時は前のコントロールのフォーカスを移動
        '        Me.ProcessTabKey(Not e.Shift)
        '        e.Handled = True
        '        '入力された商品コードを元に商品名を取得する。
        '        'ログインチェックFunction
        '        Result = GetItemName(ChkItemCodeString, 1, ItemName, Result, ErrorMessage)
        '        If Result = "True" Then
        '            Label7.Text = ItemName(0).I_NAME
        '            ComboBox1.Focus()
        '            TextBox6.BackColor = Color.White
        '        ElseIf Result = "False" Then
        '            MsgBox(ErrorMessage)
        '            TextBox6.Focus()
        '            TextBox6.BackColor = Color.Salmon
        '            'エラーの場合、商品名もクリア。
        '            Label7.Text = ""
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim ItemName As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ItemJan_Flg As Integer = 0

        Dim SearchResult() As Stock_Search_List = Nothing

        Dim ZeroDataFlg As Integer = 0

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        zSearchFLg = False

        Dim PL_ID As String = Nothing

        If ComboBox1.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox1.SelectedValue.ToString
        End If

        ChkItemJanString = Trim(TextBox6.Text)
        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        ItemName = Trim(TextBox1.Text)
        ItemName = ItemName.Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox2.Text = "商品コード" Then
            ItemJan_Flg = 1
        ElseIf ComboBox2.Text = "JANコード" Then
            ItemJan_Flg = 2
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox2.BackColor = Color.Salmon
            ComboBox2.Focus()
            Exit Sub
        End If

        '数量0のデータを表示する or しない。
        If RadioButton1.Checked = True Then
            ZeroDataFlg = 1
        ElseIf RadioButton2.Checked = True Then
            ZeroDataFlg = 2
        End If

        '在庫検索Function
        Result = GetStockSearch(ChkItemJanString, ItemJan_Flg, PL_ID, CheckBox6.Checked, CheckBox5.Checked, CheckBox2.Checked, _
                            ItemName, ZeroDataFlg, PLACE_ID, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            Label11.Text = "商品数："
            Label4.Text = "総数："
            MsgBox(ErrorMessage)

            Exit Sub
        End If

        '取得結果を表示
        '件数を表示
        Label11.Text = "商品数：" & Data_Total
        '総商品数を表示
        Label4.Text = "総数：" & Data_Num_Total


        'DataGridへ入力したデータを挿入
        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '商品コード
                .Cells(1).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(2).Value = SearchResult(Count).I_NAME
                'JANコード
                .Cells(3).Value = SearchResult(Count).JAN
                '数量
                .Cells(4).Value = SearchResult(Count).NUM
                'ロケーション
                .Cells(5).Value = SearchResult(Count).LOCATION

                'プロダクトラインコード名
                .Cells(6).Value = SearchResult(Count).PL_NAME
                'セット商品ステータス
                If SearchResult(Count).PACKAGE_FLG = False Then
                    .Cells(7).Value = "通常商品"
                ElseIf SearchResult(Count).PACKAGE_FLG = True Then
                    .Cells(7).Value = "セット商品"
                End If
                '不良区分
                .Cells(8).Value = SearchResult(Count).I_STATUS
                '倉庫
                .Cells(9).Value = SearchResult(Count).PLACE
                'STOCK.ID
                .Cells(10).Value = SearchResult(Count).ID
                'I_ID
                .Cells(11).Value = SearchResult(Count).I_ID
                'P_ID
                .Cells(12).Value = SearchResult(Count).P_ID
            End With
            DataGridView1.Rows.Add(SR_List)
        Next



    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount As Integer = 12

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
                    If CheckBox1.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.White
                        Next
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox1.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = 1
                        For Count = 0 To IndexCount
                            DataGridView1.Item(Count, CheckData(LoopCount)).Style.BackColor = Color.PaleGreen
                        Next
                    End If
                End If
            Next
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
        Dim IndexCount As Integer = 12
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

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '棚カードボタンが押されたら

        Dim Data_Check_Flg As Boolean = False
        Dim RackCard_Count As Integer = 0

        Dim WhereSQL As String = Nothing

        Dim Stock_Check() As Stock_Tanaoroshi_List = Nothing
        Dim RackCardResultCount As Integer = 0

        Dim RackCard_List() As RackCard_List = Nothing
        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing

        '印刷用設定変数
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
                Data_Check_Flg = True
            End If
        Next

        If Data_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータの商品IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Stock_Check(0 To RackCard_Count)
                'I_ID
                Stock_Check(RackCard_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()

                RackCard_Count += 1
            End If
        Next

        For i = 0 To Stock_Check.Length - 1

            If i <> Stock_Check.Length - 1 Then
                WhereSQL &= Stock_Check(i).I_ID & ","
            Else
                WhereSQL &= Stock_Check(i).I_ID
            End If
        Next

        '商品コード、商品名、在庫数量を取得
        Result = GetRackCardList(RackCard_List, WhereSQL, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
        End If

        'データがなければ何もしない。
        If RackCard_List.Length <> 0 Then

            'レポートを開く
            AxReport1.ReportPath = PrtForm & "ShelfCard.crp"
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
            If AxReport1.OpenPrintJob("ShelfCard.crp", 512, -1, "棚カード　プレビュー", 0) = False Then
                'エラー処理を記述します 
                MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                Exit Sub
            End If

            For i = 0 To RackCard_List.Length - 1
                '製品コードの表示
                AxReport1.Item("", "I_CODE").Text = RackCard_List(i).I_CODE
                '製品名の表示
                AxReport1.Item("", "I_NAME").Text = " " & RackCard_List(i).I_NAME
                '在庫数の表示
                AxReport1.Item("", "STOCK").Text = RackCard_List(i).STOCK_NUM

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

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '棚卸表出力ボタンが押されたら

        Dim Tanaoroshi_List() As Tanaoroshi_PrtList = Nothing
        Dim Result As Boolean = True
        Dim ErrorMessage As String = Nothing
        Dim Page As Integer = 0
        Dim MaxPage As Integer = 0
        Dim LastPage As Integer = 0
        Dim DataCount As Integer = 0
        Dim SQL As String = Nothing


        '印字用に今日の日付を取得
        Dim dtNow As DateTime = DateTime.Now
        '1回のプリンターにジョブを投げるページ数を設定。
        Dim Unit As Integer = 50
        '1ジョブで印字するデータ数(1ページ2件なので Unit×2の値を格納）
        Dim LoopData As Integer = Unit * 2

        Dim Stock_Data_Count As Integer = 0
        Dim Stock_Data() As StockID_List = Nothing

        Dim Check_Flg = True
        Dim Check_Data_Flg As Boolean = False

        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        'チェックを入れたデータを格納。
        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Stock_Data(0 To Stock_Data_Count)
                '在庫ID
                'Stock_Data(Stock_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(9).Value()
                Stock_Data(Stock_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()

                Stock_Data_Count += 1
                Check_Data_Flg = True

            End If
        Next

        '一件もチェックされていなかったらエラーメッセージ表示
        If Check_Data_Flg = False Then
            MsgBox("一件もデータがチェックされていません。")
            Exit Sub

        End If

        For count = 0 To Stock_Data.Length - 1
            If count = Stock_Data.Length - 1 Then
                SQL &= Stock_Data(count).STOCK_ID
            Else
                SQL &= Stock_Data(count).STOCK_ID & ","
            End If

        Next

        'チェックしたデータを渡し、データを取得する。
        Result = Tanaoroshi_Prt(SQL, Tanaoroshi_List, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

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

        'データが0件なら処理を行わない。
        If Tanaoroshi_List.Length <> 0 Then

            'ジョブ数を求める
            MaxPage = (Tanaoroshi_List.Length / 2) \ Unit

            '剰余を格納（最後のページのループ数に使用）
            LastPage = (Tanaoroshi_List.Length / 2) Mod Unit
            '求めたページ数分ループ
            For Page = 0 To MaxPage

                'レポートを開く
                AxReport1.ReportPath = PrtForm & "InventorySheet.crp"
                AxReport1.Copies = 1

                '用紙・プリンタを設定
                AxReport1.Orientation = nSvOrientation
                AxReport1.PaperSize = nSvPaperSize
                AxReport1.PaperLength = nSvPaperLength
                AxReport1.PaperWidth = nSvPaperWidth
                AxReport1.DefaultSource = nSvDefaultSource
                AxReport1.PrinterName = sSvPrinterName

                '印刷JOBの開始 プレビューを表示する。印刷設定ダイアログを表示する
                '2項目目は8(プリンタドライバのデフォルトの用紙を使用して印刷）を設定
                '3項目目を-1にするとプレビュー、それ以外は印刷。
                If AxReport1.OpenPrintJob("InventorySheet.crp", 512, -1, "棚卸表プレビュー", 0) = False Then
                    'エラー処理を記述します 
                    MsgBox("印刷ジョブ開始時にエラーが発生しました。")
                    Exit Sub
                End If

                '1ページ2データで設定件数分出力
                '最終ページのみ、上部で求めた剰余の件数分ループする。
                If Page = MaxPage Then
                    If MaxPage = 0 Then
                        LoopData = LastPage
                    Else
                        LoopData = LastPage - 1
                    End If
                    LoopData = LastPage - 1
                    If LoopData < 0 Then
                        LoopData = 0
                    End If
                Else

                    LoopData = 49
                End If

                For i = 0 To LoopData
                    'ページ上部の棚卸実施日を設定 
                    AxReport1.Item("", "Label1").Text = dtNow.ToString("yyyy/MM/dd")
                    'ページ上部のロケーションを設定
                    AxReport1.Item("", "Label2").Text = Tanaoroshi_List(DataCount).LOCATION
                    'ページ上部のプロダクトラインを設定
                    AxReport1.Item("", "Label3").Text = Tanaoroshi_List(DataCount).PL_NAME
                    'ページ上部の商品コードを設定
                    AxReport1.Item("", "Label4").Text = Tanaoroshi_List(DataCount).I_CODE
                    'ページ上部の商品名を設定
                    AxReport1.Item("", "Label5").Text = Tanaoroshi_List(DataCount).I_NAME
                    'ページ上部の在庫数を設定
                    AxReport1.Item("", "Label6").Text = Tanaoroshi_List(DataCount).NUM
                    If Tanaoroshi_List.Length > DataCount + 1 Then
                        'ページ下部の棚卸実施日を設定 
                        AxReport1.Item("", "Label7").Text = dtNow.ToString("yyyy/MM/dd")
                        'ページ下部のロケーションを設定
                        AxReport1.Item("", "Label8").Text = Tanaoroshi_List(DataCount + 1).LOCATION
                        'ページ下部のプロダクトラインを設定
                        AxReport1.Item("", "Label9").Text = Tanaoroshi_List(DataCount + 1).PL_NAME
                        'ページ下部の商品コードを設定
                        AxReport1.Item("", "Label10").Text = Tanaoroshi_List(DataCount + 1).I_CODE
                        'ページ下部の商品名を設定
                        AxReport1.Item("", "Label11").Text = Tanaoroshi_List(DataCount + 1).I_NAME
                        'ページ下部の在庫数を設定
                        AxReport1.Item("", "Label12").Text = Tanaoroshi_List(DataCount + 1).NUM

                        DataCount = DataCount + 2
                    Else
                        'データのない場合は、空白を設定
                        'ページ下部の棚卸実施日を設定 
                        AxReport1.Item("", "Label7").Text = ""
                        'ページ下部のロケーションを設定
                        AxReport1.Item("", "Label8").Text = ""
                        'ページ下部のプロダクトラインを設定
                        AxReport1.Item("", "Label9").Text = ""
                        'ページ下部の商品コードを設定
                        AxReport1.Item("", "Label10").Text = ""
                        'ページ下部の商品名を設定
                        AxReport1.Item("", "Label11").Text = ""
                        'ページ下部の在庫数を設定
                        AxReport1.Item("", "Label12").Text = ""
                    End If
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
            Next
        Else
            MsgBox("出力データがない為、ファイルは作成されませんでした。")
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        zset.Show()
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        Dim ItemName As String = Nothing
        Dim ChkItemJanString As String = Nothing
        Dim ItemJan_Flg As Integer = 0

        Dim SearchResult() As Stock_Search_List = Nothing

        zSearchFLg = False

        Dim PL_ID As String = Nothing

        Dim dtNow As DateTime = DateTime.Now

        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim LineData As String = Nothing

        Dim Csv_Complete_Message As String = Nothing

        Dim ZeroDataFlg As Integer = 0

        If ComboBox1.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox1.SelectedValue.ToString
        End If

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        ChkItemJanString = Trim(TextBox6.Text)
        'アイテムコードに'が入力されていたらReplaceする。
        ChkItemJanString = ChkItemJanString.Replace("'", "''")

        ItemName = Trim(TextBox1.Text)
        ItemName = ItemName.Replace("'", "''")

        '商品コードとJANコードのどちらが選択されているか（商品コード:1、JANコード:2）
        If ComboBox2.Text = "商品コード" Then
            ItemJan_Flg = 1
        ElseIf ComboBox2.Text = "JANコード" Then
            ItemJan_Flg = 2
        Else
            MsgBox("商品コードかJANコードを選択してください。")
            ComboBox2.BackColor = Color.Salmon
            ComboBox2.Focus()
            Exit Sub
        End If

        '数量0のデータを表示する or しない。
        If RadioButton1.Checked = True Then
            ZeroDataFlg = 1
        ElseIf RadioButton2.Checked = True Then
            ZeroDataFlg = 2
        End If

        '在庫検索Function
        Result = GetStockSearch(ChkItemJanString, ItemJan_Flg, PL_ID, CheckBox6.Checked, CheckBox5.Checked, CheckBox2.Checked, _
                            ItemName, ZeroDataFlg, PLACE_ID, SearchResult, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            '商品数、総数をクリア
            Label11.Text = "商品数："
            Label4.Text = "総数："
            MsgBox(ErrorMessage)

            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "在庫データ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "商品コード,商品名,JANコード,数量,ロケーション,プロダクトライン名,セット商品ステータス,不良区分,倉庫"

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
            'A列に商品コード
            LineData = """" & SearchResult(i).I_CODE & ""","
            'B列に商品名
            LineData &= """" & SearchResult(i).I_NAME.Replace("""", """""") & ""","
            'C列にJANコード
            LineData &= """" & SearchResult(i).JAN & ""","
            'D列に数量
            LineData &= SearchResult(i).NUM & ","
            'E列にロケーション
            LineData &= """" & SearchResult(i).LOCATION & ""","
            'F列にプロダクトライン名
            LineData &= """" & SearchResult(i).PL_NAME & ""","
            'G列にセット商品ステータス
            'セット商品ステータス
            If SearchResult(i).PACKAGE_FLG = False Then
                LineData &= """通常商品"","
            ElseIf SearchResult(i).PACKAGE_FLG = True Then
                LineData &= """セット商品"","
            End If
            'H列に不良区分
            LineData &= """" & SearchResult(i).I_STATUS & ""","
            'I列に倉庫
            LineData &= """" & SearchResult(i).PLACE & """"
            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "在庫データCSVの作成が完了しました。"

        MsgBox(Csv_Complete_Message)

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        Dim Stock_Check_Flg As Boolean = False
        Dim Stock_Data_Count As Integer = 0
        Dim Stock_Check() As Stock_Tanaoroshi_List = Nothing

        If zSearchFLg = True Then
            MsgBox("棚卸、在庫調整、セット組み換え、返品出荷、良品変更を行った後は再度検索を行って下さい。")
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
                zSearchFLg = True
            End If
        Next

        If zSearchFLg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If


        'DataGridViewからチェックされたデータのIDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Stock_Check(0 To Stock_Data_Count)
                '商品コード
                Stock_Check(Stock_Data_Count).I_CODE = DataGridView1.Rows(Count).Cells(1).Value()
                '商品名
                Stock_Check(Stock_Data_Count).I_NAME = DataGridView1.Rows(Count).Cells(2).Value()
                'JANコード
                Stock_Check(Stock_Data_Count).JAN = DataGridView1.Rows(Count).Cells(3).Value()
                '数量
                Stock_Check(Stock_Data_Count).NUM = DataGridView1.Rows(Count).Cells(4).Value()
                'ロケーション
                Stock_Check(Stock_Data_Count).LOCATION = DataGridView1.Rows(Count).Cells(5).Value()

                'プロダクトコード名
                Stock_Check(Stock_Data_Count).PL_NAME = DataGridView1.Rows(Count).Cells(6).Value()
                '不良区分
                Stock_Check(Stock_Data_Count).I_STATUS = DataGridView1.Rows(Count).Cells(8).Value()
                '倉庫
                Stock_Check(Stock_Data_Count).PLACE = DataGridView1.Rows(Count).Cells(9).Value()
                'Stock.ID
                Stock_Check(Stock_Data_Count).STOCK_ID = DataGridView1.Rows(Count).Cells(10).Value()
                'I_ID
                Stock_Check(Stock_Data_Count).I_ID = DataGridView1.Rows(Count).Cells(11).Value()
                'P_ID
                Stock_Check(Stock_Data_Count).PLACE_ID = DataGridView1.Rows(Count).Cells(12).Value()
                Stock_Data_Count += 1
            End If
        Next

        'zlocationのDataGridViewをクリア
        zlocation.DataGridView1.Rows.Clear()

        'zlocationのDataGridViewに値を入れていく。
        For Count = 0 To Stock_Check.Length - 1
            Dim LocationChange_Data_list As New DataGridViewRow
            LocationChange_Data_list.CreateCells(zlocation.DataGridView1)
            With LocationChange_Data_list
                '商品コード
                .Cells(0).Value = Stock_Check(Count).I_CODE
                '商品名
                .Cells(1).Value = Stock_Check(Count).I_NAME
                '数量
                .Cells(2).Value = Stock_Check(Count).NUM
                'ロケーション
                .Cells(3).Value = Stock_Check(Count).LOCATION
                '移動ロケーション
                .Cells(4).Value = Stock_Check(Count).LOCATION
                '倉庫
                .Cells(5).Value = Stock_Check(Count).PLACE
                '不良区分
                .Cells(6).Value = Stock_Check(Count).I_STATUS
                'STOCK_ID
                .Cells(7).Value = Stock_Check(Count).STOCK_ID
                'I_ID
                .Cells(8).Value = Stock_Check(Count).I_ID
                'P_ID
                .Cells(9).Value = Stock_Check(Count).PLACE_ID
            End With
            zlocation.DataGridView1.Rows.Add(LocationChange_Data_list)
        Next

        zlocation.Show()
        Me.Hide()
    End Sub
End Class