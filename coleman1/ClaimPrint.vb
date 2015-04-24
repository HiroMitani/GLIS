Imports System.Windows.Forms
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class ClaimPrint

    Public FormLord As Boolean = False

    '更新を行ったら、再度検索を行わせる為のFlg
    Public ClaimSearchFLg As Boolean = False

    '帳票ファイルの格納フォルダ指定
    Public PrtForm As String = System.Configuration.ConfigurationManager.AppSettings("PrtPath")

    Private Sub ClaimPrint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "請求書印刷"

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

        '左3項目を固定(チェック、商品コード、商品名)
        DataGridView1.Columns(2).Frozen = True

        FormLord = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim Date_Check_Result As DateTime

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim ChkClaimCodeString As String = Nothing

        Dim SearchResult() As Claim_List = Nothing

        '一括チェックのチェックボックスをクリアする。
        CheckBox5.Checked = False

        'DataGridView（検索結果）をクリアする。
        DataGridView1.Rows.Clear()

        '検索ボタンクリックチェック
        ClaimSearchFLg = False

        '出荷日のFrom、Toともに入力されていなければエラー
        If MaskedTextBox1.Text = "    /  /" And MaskedTextBox2.Text = "    /  /" Then
            MsgBox("出荷日は必須項目です。")
            Exit Sub
        End If

        '出荷日Fromのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '出荷日Toのチェック
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
                MsgBox("出荷日の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        '請求先コードに'が入力されていたらReplaceする。
        ChkClaimCodeString = Trim(TextBox1.Text)
        ChkClaimCodeString = ChkClaimCodeString.Replace("'", "''")

        '検索Function
        Result = GetClaimSeach(Date_From, Date_To, ChkClaimCodeString, CheckBox1.Checked, CheckBox2.Checked, _
                            CheckBox3.Checked, CheckBox4.Checked, SearchResult, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        For Count = 0 To SearchResult.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '請求先コード
                .Cells(1).Value = SearchResult(Count).CLAIM_CODE
                'ステータス
                .Cells(2).Value = SearchResult(Count).STATUS
                '伝票番号
                .Cells(3).Value = SearchResult(Count).SHEET_NO
                '納品先コード
                .Cells(4).Value = SearchResult(Count).C_CODE
                '納品先名
                .Cells(5).Value = SearchResult(Count).D_NAME
                '商品コード
                .Cells(6).Value = SearchResult(Count).I_CODE
                '商品名
                .Cells(7).Value = SearchResult(Count).I_NAME
                '出荷数量
                .Cells(8).Value = SearchResult(Count).FIX_NUM

                '納品単価
                .Cells(9).Value = SearchResult(Count).UNIT_COST
                '請求書印刷日
                .Cells(10).Value = SearchResult(Count).CLAIM_PRT_DATE
                '出荷指示ファイル
                .Cells(11).Value = SearchResult(Count).FILE_NAME
                'オーダー番号
                .Cells(12).Value = SearchResult(Count).ORDER_NO
                '出荷日
                .Cells(13).Value = SearchResult(Count).FIX_DATE
                '出荷倉庫
                .Cells(14).Value = SearchResult(Count).P_NAME
                'コメント１
                .Cells(15).Value = SearchResult(Count).COMMENT1
                'コメント２
                .Cells(16).Value = SearchResult(Count).COMMENT2
                'OUT_TBL.ID
                .Cells(17).Value = SearchResult(Count).ID
                'OU_TBL.I_ID
                .Cells(18).Value = SearchResult(Count).I_ID
                'OUT_TBL.C_ID
                .Cells(19).Value = SearchResult(Count).C_ID
                '請求書番号
                .Cells(20).Value = SearchResult(Count).CLAIM_NO

            End With
            DataGridView1.Rows.Add(SR_List)
        Next

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Data_Count As Integer = 0
        Dim Upd_Data() As Claim_List = Nothing

        Dim UpdResult As Boolean = True
        Dim UpdErrorMessage As String = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim DataGridErrorMessage As String = Nothing

        Dim Error_Flg As Boolean = True

        Dim Check_Flg As Boolean = False

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
            MsgBox("チェックされたデータがありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータの納品単価、IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = True Then
                ReDim Preserve Upd_Data(0 To Data_Count)

                '納品単価の入力妥当性チェック
                '***チェック項目***
                ' 整数値、未入力NG、0の値OK、マイナス値NG
                If NumChkVal(Trim(DataGridView1.Rows(Count).Cells(9).Value()), "INTEGER", False, True, False, NumChkResult, NumChkErrorMessage) = False Then
                    'MsgBox(NumChkErrorMessage)
                    DataGridView1(9, Count).Style.BackColor = Color.Salmon
                    DataGridErrorMessage &= Count + 1 & "行目の納品単価が正しくありません。" & vbCr

                    Error_Flg = False
                Else
                    '納品単価
                    Upd_Data(Data_Count).UNIT_COST = DataGridView1.Rows(Count).Cells(9).Value()
                End If
                'ID
                Upd_Data(Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                Data_Count += 1
            End If
        Next

        If Error_Flg = False Then
            MsgBox(DataGridErrorMessage)
            Exit Sub
        End If
        'ダイアログ設定
        Dim result As DialogResult = MessageBox.Show("データを更新してもよろしいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)
        '何が選択されたか調べる()
        If result = DialogResult.No Then
            '「いいえ」が選択された時 
            Exit Sub
        End If

        '更新Function
        UpdResult = UpdOutTbl_UnitCost(Upd_Data, UpdResult, UpdErrorMessage)

        If Result = False Then
            MsgBox(UpdErrorMessage)
            Exit Sub
        End If

        MsgBox("修正が完了しました。")

        ClaimSearchFLg = True

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Check_Flg As Boolean = False

        Dim Prt_Data() As Claim_List = Nothing
        Dim ClaimSearchResult() As Claim_List = Nothing
        Dim ClaimDetailIDList() As Claim_List = Nothing
        Dim ClaimDetailSearchResult() As Claim_List = Nothing
        Dim Data_Count As Integer
        Dim DataCheckFlg As Boolean = True

        Dim UpdResult As Boolean = True
        Dim UpdErrorMessage As String = Nothing

        Dim DataGridErrorMessage As String = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim ResultDataCount As Integer

        Dim PageDataCount As Integer
        Dim LineCount As Integer
        Dim LoopCount As Integer

        '印字用に日時を取得
        Dim dtNow As DateTime = DateTime.Now

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 21
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0
        '合計を格納（税抜）
        Dim Total As Integer
        '合計を格納（税込）
        Dim TaxTotal As Integer
        'ページごとの小計を格納
        Dim SubTotal As Integer

        Dim ClaimDate As String = DateTimePicker3.Value.ToString("yyyy/MM/dd")
        '請求書番号（最後にOUT_TBLにInsert用）
        Dim Claim_No As String = Nothing


        If ClaimSearchFLg = True Then
            MsgBox("納品単価の更新を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        '請求書コードで検索が行われていなければエラー
        If Trim(TextBox1.Text) = "" Then
            MsgBox("請求書を印刷するには請求書コードでの検索が必須となります。")
            Exit Sub
        End If

        '検索結果の請求先コードが検索条件の請求先コードと全て同じかチェック
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(1).Value() <> Trim(TextBox1.Text) Then
                DataCheckFlg = False
            End If
        Next

        If DataCheckFlg = False Then
            MsgBox("検索条件の請求先コードと異なるデータが存在しています。再度検索を行ってください。")
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
            MsgBox("チェックされたデータがありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータの納品単価、IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Prt_Data(0 To Data_Count)
                'ID
                Prt_Data(Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                Data_Count += 1
            End If
        Next

        '請求先一覧を取得
        Result = GetClaimPrtData(Trim(TextBox1.Text), Prt_Data, ClaimSearchResult, ResultDataCount, Total, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '請求書番号を格納
        Claim_No = ClaimSearchResult(0).CLAIM_CODE & ClaimDate.Replace("/", "")

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "ClaimList.crp"

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
        AxReport1.ReportPath = PrtForm & "ClaimList.crp"
        '用紙・プリンタを設定
        AxReport1.Orientation = nSvOrientation
        AxReport1.PaperSize = nSvPaperSize
        AxReport1.PaperLength = nSvPaperLength
        AxReport1.PaperWidth = nSvPaperWidth
        AxReport1.DefaultSource = nSvDefaultSource
        AxReport1.PrinterName = sSvPrinterName
        AxReport1.Copies = 1

        If AxReport1.OpenPrintJob("ClaimList.crp", 512, -1, "請求書　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            Exit Sub
        End If

        'データが1ページのMAX件数以下ならMAXPageに1を設定
        If ResultDataCount <= Max Then
            MaxPage = 1
        Else
            'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
            If ResultDataCount Mod Max = 0 Then
                MaxPage = ResultDataCount \ Max
            Else
                MaxPage = ResultDataCount \ Max + 1
            End If
        End If

        PageDataCount = 0
        For Page = 1 To MaxPage

            'ヘッダー部分のデータを反映
            'Page No
            AxReport1.Item("", "PageNo").Text = Page & " "
            '発行日
            AxReport1.Item("", "Claim_Date").Text = ClaimDate & " "
            '請求書番号（請求書コード+発行日（yyyymmdd））
            AxReport1.Item("", "Claim_No").Text = ClaimSearchResult(0).CLAIM_CODE & ClaimDate.Replace("/", "") & " "

            '請求先名
            AxReport1.Item("", "Claim_Name").Text = ClaimSearchResult(0).C_NAME & " "


            TaxTotal = Math.Round(Total * (1 + Tax / 100))
            '合計金額（税込）
            AxReport1.Item("", "TotalTax").Text = Format(TaxTotal, "#,#") & " "
            '合計金額（税抜）
            AxReport1.Item("", "Total").Text = Format(Total, "#,#") & " "
            '消費税額
            AxReport1.Item("", "Tax").Text = Format(TaxTotal - Total, "#,#") & " "

            '企業情報
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

            '企業振込先銀行名
            AxReport1.Item("", "Bank_Name").Text = Com_BANKNAME
            '企業振込先口座名
            AxReport1.Item("", "Account_Info").Text = Com_ACCOUNTINFO

            'メモ欄１
            AxReport1.Item("", "Memo1").Text = Memo1
            'メモ欄２
            AxReport1.Item("", "Memo2").Text = Memo2

            LineCount = 1
            'Max件数以下なら、データ件数分ループするように設定
            If ResultDataCount <= Max Then
                LoopCount = ResultDataCount
            Else
                'Max件数以上ならMax値までループを設定するが
                '最終ページはデータ件数までの値を設定する。
                If Page = MaxPage Then
                    LoopCount = ResultDataCount - ((Page - 1) * Max)
                Else
                    LoopCount = Max
                End If
            End If

            '一度明細行をクリア
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
            Next

            'ページごとの小計変数もクリア
            SubTotal = 0

            For i = 1 To LoopCount
                '発送日
                AxReport1.Item("", "Label" & i & "-1").Text = " " & ClaimSearchResult(DataCount).FIX_DATE
                '納品番号
                AxReport1.Item("", "Label" & i & "-2").Text = " " & ClaimSearchResult(DataCount).SHEET_NO
                '納品先コード
                AxReport1.Item("", "Label" & i & "-3").Text = " " & ClaimSearchResult(DataCount).C_CODE
                '納品先名
                AxReport1.Item("", "Label" & i & "-4").Text = " " & ClaimSearchResult(DataCount).D_NAME
                '納品数量
                AxReport1.Item("", "Label" & i & "-5").Text = ClaimSearchResult(DataCount).FIX_NUM & " "
                '納入金額
                AxReport1.Item("", "Label" & i & "-6").Text = Format(ClaimSearchResult(DataCount).FIX_NUM * ClaimSearchResult(DataCount).UNIT_COST, "#,#") & " "

                SubTotal += ClaimSearchResult(DataCount).FIX_NUM * ClaimSearchResult(DataCount).UNIT_COST

                LineCount += 1
                DataCount += 1
            Next

            'Footerの設定
            '小計
            AxReport1.Item("", "SubTotal").Text = Format(SubTotal, "#,#") & " "
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

        '出力が終了したら請求書番号と請求書日付を更新
        '明細のOUT_TBL.IDを全て渡し、請求書印刷日に日付をアップデートする。
        '更新Function
        UpdResult = UpdOutTbl_ClaimDate(Claim_No, ClaimDate, ClaimSearchResult, UpdResult, UpdErrorMessage)

        If Result = False Then
            MsgBox(UpdErrorMessage)
            Exit Sub
        End If

        MsgBox("請求書の印刷が完了しました。")

        ClaimSearchFLg = True

    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount = 20

        Dim CheckDataCount As Integer = 0

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
                    If CheckBox5.Checked = False Then
                        DataGridView1(0, CheckData(LoopCount)).Value = 0
                    End If
                ElseIf DataGridView1(0, CheckData(LoopCount)).Value = 0 Then
                    If CheckBox5.Checked = True Then

                        DataGridView1(0, CheckData(LoopCount)).Value = 1
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Prt_Data() As Claim_List = Nothing
        Dim Check_Flg As Boolean = False
        Dim DataCheckFlg As Boolean = True
        Dim Data_Count As Integer
        Dim ClaimSearchResult() As Claim_List = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim ResultDataCount As Integer

        Dim PageDataCount As Integer
        Dim LineCount As Integer
        Dim LoopCount As Integer

        '帳票の1ページあたりのデータ件数の設定
        Dim Max As Integer = 21
        '印刷するページ数を格納
        Dim MaxPage As Integer = 1
        '現在のページ数
        Dim Page As Integer = 1
        'データカウント用変数
        Dim DataCount As Integer = 0
        '合計を格納（税抜）
        Dim Total As Integer
        '合計を格納（税込）
        Dim TaxTotal As Integer
        'ページごとの小計を格納
        Dim SubTotal As Integer

        If ClaimSearchFLg = True Then
            MsgBox("納品単価の更新を行った後は再度検索を行って下さい。")
            Exit Sub
        End If

        '請求書コードで検索が行われていなければエラー
        If Trim(TextBox2.Text) = "" Then
            MsgBox("請求書を再印刷するには請求書番号での検索が必須となります。")
            Exit Sub
        End If

        '検索結果の請求書番号が検索条件の請求番号と全て同じかチェック
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(20).Value() <> Trim(TextBox2.Text) Then
                DataCheckFlg = False
            End If
        Next

        If DataCheckFlg = False Then
            MsgBox("検索条件の請求先番号と異なるデータが存在しています。再度検索を行ってください。")
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
            MsgBox("チェックされたデータがありません。")
            Exit Sub
        End If

        'DataGridViewからチェックされたデータの納品単価、IDを配列に格納
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve Prt_Data(0 To Data_Count)
                'ID
                Prt_Data(Data_Count).ID = DataGridView1.Rows(Count).Cells(17).Value()
                Data_Count += 1
            End If
        Next

        '再印刷用請求先一覧を取得
        Result = GetReClaimPrtData(Trim(TextBox2.Text), ClaimSearchResult, ResultDataCount, Total, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '印刷用設定変数
        Dim nSvOrientation As Long
        Dim nSvPaperSize As Long
        Dim nSvPaperLength As Long
        Dim nSvPaperWidth As Long
        Dim nSvDefaultSource As Long
        Dim sSvPrinterName As String

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "ClaimList.crp"

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
        AxReport1.ReportPath = PrtForm & "ClaimList.crp"
        '用紙・プリンタを設定
        AxReport1.Orientation = nSvOrientation
        AxReport1.PaperSize = nSvPaperSize
        AxReport1.PaperLength = nSvPaperLength
        AxReport1.PaperWidth = nSvPaperWidth
        AxReport1.DefaultSource = nSvDefaultSource
        AxReport1.PrinterName = sSvPrinterName
        AxReport1.Copies = 1

        If AxReport1.OpenPrintJob("ClaimList.crp", 512, -1, "請求書　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            Exit Sub
        End If

        'データが1ページのMAX件数以下ならMAXPageに1を設定
        If ResultDataCount <= Max Then
            MaxPage = 1
        Else
            'MAX件以上なら、余りと商を求め、Modで余りが出るなら+1
            If ResultDataCount Mod Max = 0 Then
                MaxPage = ResultDataCount \ Max
            Else
                MaxPage = ResultDataCount \ Max + 1
            End If
        End If

        PageDataCount = 0
        For Page = 1 To MaxPage

            'ヘッダー部分のデータを反映
            'Page No
            AxReport1.Item("", "PageNo").Text = Page & " "
            '発行日
            AxReport1.Item("", "Claim_Date").Text = ClaimSearchResult(0).CLAIM_PRT_DATE & " "
            '請求書番号（請求書コード+発行日（yyyymmdd））
            AxReport1.Item("", "Claim_No").Text = ClaimSearchResult(0).CLAIM_NO & " "

            '請求先名
            AxReport1.Item("", "Claim_Name").Text = ClaimSearchResult(0).C_NAME & " "


            TaxTotal = Math.Round(Total * (1 + Tax / 100))
            '合計金額（税込）
            AxReport1.Item("", "TotalTax").Text = Format(TaxTotal, "#,#") & " "
            '合計金額（税抜）
            AxReport1.Item("", "Total").Text = Format(Total, "#,#") & " "
            '消費税額
            AxReport1.Item("", "Tax").Text = Format(TaxTotal - Total, "#,#") & " "

            '企業情報
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

            '企業振込先銀行名
            AxReport1.Item("", "Bank_Name").Text = Com_BANKNAME
            '企業振込先口座名
            AxReport1.Item("", "Account_Info").Text = Com_ACCOUNTINFO

            'メモ欄１
            AxReport1.Item("", "Memo1").Text = Memo1
            'メモ欄２
            AxReport1.Item("", "Memo2").Text = Memo2

            LineCount = 1
            'Max件数以下なら、データ件数分ループするように設定
            If ResultDataCount <= Max Then
                LoopCount = ResultDataCount
            Else
                'Max件数以上ならMax値までループを設定するが
                '最終ページはデータ件数までの値を設定する。
                If Page = MaxPage Then
                    LoopCount = ResultDataCount - ((Page - 1) * Max)
                Else
                    LoopCount = Max
                End If
            End If

            '一度明細行をクリア
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
            Next

            'ページごとの小計変数もクリア
            SubTotal = 0

            For i = 1 To LoopCount
                '発送日
                AxReport1.Item("", "Label" & i & "-1").Text = " " & ClaimSearchResult(DataCount).FIX_DATE
                '納品番号
                AxReport1.Item("", "Label" & i & "-2").Text = " " & ClaimSearchResult(DataCount).SHEET_NO
                '納品先コード
                AxReport1.Item("", "Label" & i & "-3").Text = " " & ClaimSearchResult(DataCount).C_CODE
                '納品先名
                AxReport1.Item("", "Label" & i & "-4").Text = " " & ClaimSearchResult(DataCount).D_NAME
                '納品数量
                AxReport1.Item("", "Label" & i & "-5").Text = ClaimSearchResult(DataCount).FIX_NUM & " "
                '納入金額
                AxReport1.Item("", "Label" & i & "-6").Text = Format(ClaimSearchResult(DataCount).FIX_NUM * ClaimSearchResult(DataCount).UNIT_COST, "#,#") & " "

                SubTotal += ClaimSearchResult(DataCount).FIX_NUM * ClaimSearchResult(DataCount).UNIT_COST

                LineCount += 1
                DataCount += 1
            Next

            'Footerの設定
            '小計
            AxReport1.Item("", "SubTotal").Text = Format(SubTotal, "#,#") & " "
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

        MsgBox("請求書の印刷が完了しました。")

        ClaimSearchFLg = True


    End Sub
End Class