Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms
Imports System.Globalization

Public Class IO_balance
    'プロダクトラインID
    Dim PL_Id As Integer

    Dim PLACE_ID As String

    '帳票ファイルの格納フォルダ指定
    Public PrtForm As String = System.Configuration.ConfigurationManager.AppSettings("PrtPath")

    'CSVの出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    'Graphの出力先
    Public GraphPath As String = System.Configuration.ConfigurationManager.AppSettings("GraphPath")

    Public FormLord As Boolean = False



    '表示件数を越えた時のアラートメッセージ
    Public MAXDataMessage As String = "表示されていないデータがあります。"

    Private Sub IO_balance_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

        Dim Search_List() As IO_Balance_List = Nothing
        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim PL_ID As Integer
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim OUTDate_From As String = Nothing
        Dim OUTDate_To As String = Nothing
        Dim Date_Check_Result As DateTime

        '検索条件のプロダクトラインをセット
        If ComboBox1.Text = "" Then
            PL_ID = 0
        Else
            PL_ID = ComboBox1.SelectedValue.ToString
        End If

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()

        '入力チェック
        '出庫年月日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            OUTDate_From = ""
            'MsgBox("出荷確定日は必須項目です。")
            'MaskedTextBox1.BackColor = Color.Salmon
            'MaskedTextBox1.Focus()
            'Exit Sub

        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUTDate_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("出荷確定日が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '希望納期Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            OUTDate_To = ""
            'MsgBox("出荷確定日は必須項目です。")
            'MaskedTextBox2.BackColor = Color.Salmon
            'MaskedTextBox2.Focus()
            'Exit Sub
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUTDate_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("出荷確定日が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        'データ取得Function
        Result = GET_IO_Balance(OUTDate_From, OUTDate_To, PL_ID, PLACE_ID, Search_List, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '結果を元にDataGridViewに表示する。
        'DataGridへ入力したデータを挿入
        For Count = 0 To Search_List.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '商品コード
                .Cells(1).Value = Search_List(Count).I_CODE
                '商品名
                .Cells(2).Value = Search_List(Count).I_NAME
                'JANコード
                .Cells(3).Value = Search_List(Count).JAN
                'プロダクトライン名
                .Cells(4).Value = Search_List(Count).PL_NAME
                '発注数
                .Cells(5).Value = Search_List(Count).PO_NUM
                '入庫予定数
                .Cells(6).Value = Search_List(Count).IN_NUM
                '入庫確定数
                .Cells(7).Value = Search_List(Count).IN_FIX_NUM

                '出庫予定数
                .Cells(8).Value = Search_List(Count).OUT_NUM
                '在庫数
                .Cells(9).Value = Search_List(Count).STOCK_NUM
                'I_ID
                .Cells(10).Value = Search_List(Count).I_ID
            End With
            DataGridView1.Rows.Add(SR_List)
        Next

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub IO_balance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PLResult As Boolean = True
        Dim PLErrorMessage As String = Nothing
        Dim PLList() As PL_List = Nothing

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "入荷出荷バランス"

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Search_List() As IO_Balance_List = Nothing
        Dim Date_From As String = Nothing
        Dim Date_To As String = Nothing
        Dim PL_ID As Integer
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim OUTDate_From As String = Nothing
        Dim OUTDate_To As String = Nothing
        Dim Date_Check_Result As DateTime
        Dim dtNow As DateTime = DateTime.Now
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter
        Dim LineData As String = Nothing
        Dim Csv_Complete_Message As String = Nothing
        Dim Data_Total As Integer = 0

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        '入力チェック
        '出庫年月日Fromのチェック
        If MaskedTextBox1.Text = "    /  /" Then
            '未入力なら""を格納
            OUTDate_From = ""
            MaskedTextBox1.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox1.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUTDate_From = MaskedTextBox1.Text
                MaskedTextBox1.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox1.BackColor = Color.Salmon
                MaskedTextBox1.Focus()
                Exit Sub
            End If
        End If

        '希望納期Toのチェック
        If MaskedTextBox2.Text = "    /  /" Then
            '未入力なら""を格納
            OUTDate_To = ""
            MaskedTextBox2.BackColor = Color.White
        Else
            '何か入力されていたら日付の妥当性チェックを行う。
            If DateTime.TryParseExact(MaskedTextBox2.Text, "yyyy/MM/dd", DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None, Date_Check_Result) Then
                'チェックを行い、問題がなければ格納
                OUTDate_To = MaskedTextBox2.Text
                MaskedTextBox2.BackColor = Color.White
            Else
                MsgBox("希望納期の日付が正しくありません。")
                MaskedTextBox2.BackColor = Color.Salmon
                MaskedTextBox2.Focus()
                Exit Sub
            End If
        End If

        'データ取得Function
        Result = GET_IO_Balance(OUTDate_From, OUTDate_To, PL_ID, PLACE_ID, Search_List, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'CSVファイル名と、項目行の設定
        Dim Sheet1_Name As String = "入荷出荷バランスデータ" & dtNow.ToString("yyyyMMddHHmm") & ".csv"
        'Header設定
        Dim Sheet1_Header As String = "商品コード,商品名,JAN,プロダクトライン名,発注数,入庫予定,入庫確定数,出庫予定,在庫数"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If Search_List.Length = 0 Then
            MsgBox("データがありません。")
            Exit Sub
        End If

        'ﾌｧｲﾙ名の設定
        strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)
        'headerを設定
        strStreamWriter.WriteLine(Sheet1_Header)

        'データ件数分ループ
        For i = 0 To Search_List.Length - 1
            '取得するデータをLineataに格納する。
            'A列に商品コード
            LineData = """" & Search_List(i).I_CODE & ""","
            'B列に商品名
            LineData &= """" & Search_List(i).I_NAME & ""","
            'C列にJAN
            LineData &= """" & Search_List(i).JAN & ""","
            'D列にプロダクトライン名
            LineData &= """" & Search_List(i).PL_NAME & ""","
            'E列に発注数
            LineData &= """" & Search_List(i).PO_NUM & ""","
            'F列に入庫予定
            LineData &= "" & Search_List(i).IN_NUM & ","
            'G列に入庫予定
            LineData &= "" & Search_List(i).IN_FIX_NUM & ","
            'H列に出庫予定
            LineData &= "" & Search_List(i).OUT_NUM & ","
            'I列に在庫数
            LineData &= """" & Search_List(i).STOCK_NUM & ""","

            '1行にしたデータを登録
            strStreamWriter.WriteLine(LineData)
        Next
        'ファイルを閉じる
        strStreamWriter.Close()

        Csv_Complete_Message = CSVPath & "に" & vbCr & "CSVの作成が完了しました。"
        MsgBox(Csv_Complete_Message)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Dim Item_Data() As Item_List = Nothing
        Dim I_Code As String = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Stock_NUM As Integer = Nothing

        'グラフ開始時の日付（今日 - 180日）の時点の在庫数を格納
        Dim START_STOCK_NUM As Integer = Nothing

        Dim GRAPH_Data() As GraphData = Nothing
        Dim GRAPH_Future_Data() As GraphData = Nothing

        Dim GRAPH_SummaryData() As GraphSummaryData = Nothing

        Dim Check_Flg As Boolean = False

        Dim SingleItem_Check_Flg As Boolean = False
        Dim CheckCount As Integer = 0

        Dim I_ID As Integer = 0

        Dim I_ID_List() As Item_List = Nothing

        Dim LoopCount As Integer = 0

        Dim GRAPH_Data_Count As Integer = 0
        Dim GRAPH_Future_Data_Count As Integer = 0

        'グラフを作る必要のない商品（入庫、出庫、発注などがない商品）の商品コードを格納
        Dim MessageList As String = Nothing
        Dim MessageFlg = True

        Dim DateCheck As Boolean = False

        ' 本日の日付を設定
        Dim dtNow As DateTime = DateTime.Now
        '本日を設定
        Dim Today As String = dtNow.ToString("yyyy/MM/dd")

        Dim LabelDatetime As String = dtNow.ToString("yyyy/MM/dd HH:mm")

        'Dim LooPEnd As Integer = (PastMonth * -1) + FutureMonth - 1

        'データ取得範囲Fromを指定
        Dim DateFrom As String = dtNow.AddDays(-180).ToString("yyyy/MM/dd")

        'データ取得範囲Toを指定
        Dim DateTo As String = dtNow.AddDays(90).ToString("yyyy/MM/dd")

        Dim MonthFrom As String = dtNow.AddDays(-180).ToString("yyyy/MM")
        Dim MonthToday As String = dtNow.ToString("yyyy/MM")

        '発注数合計
        Dim PO_Total As Integer = 0
        '入出庫数合計
        Dim INOUT_Total As Integer = 0

        '履歴の合計
        Dim Month_Total As Integer = 0

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
                Check_Flg = True
            End If
        Next

        'チェックが１件もなければエラーメッセージ
        If Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        Dim DataCount As Integer = 0
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                ReDim Preserve I_ID_List(0 To DataCount)
                'I_IDを格納
                I_ID_List(DataCount).ID = DataGridView1.Rows(Count).Cells(10).Value()
                DataCount += 1
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

        '処理中ウインドウを表示
        Processing.Show()

        'メッセージのセット
        'MessageList = "以下の商品コードは検索期間(" & DateFrom & "～" & DateTo & ")に入庫、出庫、発注がない為、" & vbCr & "グラフの作成が行われませんでした。" & vbCr
        MessageList = "以下の商品コードは検索期間に入庫、出庫、発注がない為、" & vbCr & "グラフの作成が行われませんでした。" & vbCr

        'Processingウインドウの「処理中です。」のメッセージをリフレッシュさせることにより表示
        Processing.Label1.Refresh()

        'レポートを開く
        AxReport1.ReportPath = PrtForm & "graph.crp"
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
        If AxReport1.OpenPrintJob("graph.crp", 512, -1, "グラフシート　プレビュー", 0) = False Then
            'エラー処理を記述します 
            MsgBox("印刷ジョブ開始時にエラーが発生しました。")
            '処理中ウインドウを表示
            Processing.Show()
            Exit Sub
        End If

        'チェックされた商品数ループ
        For I_ID_Loop = 0 To I_ID_List.Length - 1
            '商品IDを元に商品名を取得。
            'データ取得Function
            Result = GetItemInfo(I_ID_List(I_ID_Loop).ID, Item_Data, Result, ErrorMessage)
            If Result = False Then
                MsgBox(ErrorMessage)
                '処理中ウインドウを表示
                Processing.Show()
                Exit Sub
            End If

            GRAPH_Data_Count = 0
            GRAPH_Future_Data_Count = 0

            ReDim GRAPH_SummaryData(0 To 35)
            Dim j As Integer = 0

            For i = 0 To 35 Step 6
                GRAPH_SummaryData(i).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i).TYPE = "入庫確定"
                GRAPH_SummaryData(i + 1).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i + 1).TYPE = "出庫確定"
                GRAPH_SummaryData(i + 2).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i + 2).TYPE = "棚卸"
                GRAPH_SummaryData(i + 3).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i + 3).TYPE = "在庫調整"
                GRAPH_SummaryData(i + 4).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i + 4).TYPE = "セット組"
                GRAPH_SummaryData(i + 5).DATE_DATA = dtNow.AddMonths(j * -1).ToString("yyyy/MM")
                GRAPH_SummaryData(i + 5).TYPE = "セットばらし"
                j += 1

            Next

            'データ取得
            Result = GetGraphData(I_ID_List(I_ID_Loop).ID, DateFrom, Today, DateTo, MonthFrom, MonthToday, Stock_NUM, START_STOCK_NUM, PLACE_ID, GRAPH_Data, GRAPH_SummaryData, GRAPH_Data_Count, GRAPH_Future_Data, GRAPH_Future_Data_Count, Result, ErrorMessage)

            If Result = False Then
                MsgBox(ErrorMessage)
                '処理中ウインドウを表示
                Processing.Show()
                Exit Sub
            End If

            If GRAPH_Data_Count <> 0 Or GRAPH_Future_Data_Count <> 0 Then

                Dim ds As New DataSet
                Dim dt As New DataTable
                Dim dtRow As DataRow

                '列の作成
                dt.Columns.Add(" ", Type.GetType("System.String"))
                dt.Columns.Add("在庫数", Type.GetType("System.String"))
                dt.Columns.Add("在庫数+発注数", Type.GetType("System.String"))
                dt.Columns.Add("本日の日付(" & Today & ")", Type.GetType("System.String"))
                ds.Tables.Add(dt)

                If GRAPH_Data_Count <> 0 Then

                    With Chart1.ChartAreas(0).AxisX
                        .MajorGrid.Interval = 14
                        .LabelStyle.Interval = 14
                        ' .MajorGrid.IntervalOffset = 0.5

                        .LabelStyle.Format = "#0"  '桁区切りで表示の場合

                    End With

                    Dim INS_DATE As String = Nothing
                    Dim INS_STOCK_NUM1 As Integer = 0
                    Dim INS_STOCK_NUM2 As Integer = 0
                    'Fromの日付の時点での在庫数を設定する。
                    Dim Tmp_StockNUM1 As Integer = START_STOCK_NUM
                    Dim Tmp_StockNUM2 As Integer = START_STOCK_NUM
                    Dim FirstData As Boolean = False
                    Dim CheckDataCount As Integer = 0
                    Dim loopdate As Date = DateFrom
                    Do Until loopdate > Today
                        '過去のデータの追加
                        For i = GRAPH_Data.Length - 1 To 0 Step -1
                            If loopdate = GRAPH_Data(i).DATE_DATA Then
                                If DateCheck = False Then
                                    FirstData = True
                                End If
                                Tmp_StockNUM1 = GRAPH_Data(i).STOCK_NUM
                                Tmp_StockNUM2 = GRAPH_Data(i).STOCK_NUM
                                INS_DATE = CStr(loopdate)
                                INS_STOCK_NUM1 = GRAPH_Data(i).STOCK_NUM
                                INS_STOCK_NUM2 = GRAPH_Data(i).STOCK_NUM

                                If i <> 0 Then
                                    If loopdate = GRAPH_Data(i - 1).DATE_DATA Then
                                        DateCheck = True
                                    Else
                                        DateCheck = False
                                        Exit For
                                    End If
                                End If
                            Else
                                INS_DATE = CStr(loopdate)
                                INS_STOCK_NUM1 = Tmp_StockNUM1
                                INS_STOCK_NUM2 = Tmp_StockNUM1
                            End If

                        Next

                        dtRow = ds.Tables(0).NewRow
                        dtRow(0) = INS_DATE
                        If FirstData = True Then
                            dtRow(1) = INS_STOCK_NUM1
                            dtRow(2) = INS_STOCK_NUM2
                        Else
                            dtRow(1) = INS_STOCK_NUM1
                            dtRow(2) = INS_STOCK_NUM1
                        End If

                        ds.Tables(0).Rows.Add(dtRow)
                        loopdate = loopdate.AddDays(1).ToString("yyyy/MM/dd")
                        CheckDataCount += 1
                    Loop
                End If

                '今日の在庫数を追加
                dtRow = ds.Tables(0).NewRow
                dtRow(0) = Today
                dtRow(1) = Stock_NUM
                dtRow(2) = Stock_NUM
                dtRow(3) = Stock_NUM
                ds.Tables(0).Rows.Add(dtRow)

                Dim loopdate2 As Date = Today
                Dim INS_FUTURE_DATE As String = Nothing
                Dim INS_FUTURE_STOCK_NUM1 As Integer = 0
                Dim INS_FUTURE_STOCK_NUM2 As Integer = 0
                Dim FirstData2 As Boolean = False
                Dim Tmp_FutureStockNUM1 As Integer = Stock_NUM
                Dim Tmp_FutureStockNUM2 As Integer = Stock_NUM
                Dim FutureFirstData As Boolean = False
                Dim CheckDataCount2 As Integer = 0

                DateCheck = False

                Do Until loopdate2 > DateTo
                    If GRAPH_Future_Data_Count <> 0 Then

                        'データの追加
                        For i = 0 To GRAPH_Future_Data.Length - 1
                            If loopdate2 = GRAPH_Future_Data(i).DATE_DATA Then
                                If DateCheck = False Then
                                    FirstData2 = True
                                End If

                                Tmp_FutureStockNUM1 = GRAPH_Future_Data(i).STOCK_NUM
                                Tmp_FutureStockNUM2 = GRAPH_Future_Data(i).STOCK_ORDER_NUM
                                INS_FUTURE_DATE = CStr(loopdate2)
                                INS_FUTURE_STOCK_NUM1 = GRAPH_Future_Data(i).STOCK_NUM
                                INS_FUTURE_STOCK_NUM2 = GRAPH_Future_Data(i).STOCK_ORDER_NUM

                                If i <> GRAPH_Future_Data.Length - 1 Then
                                    If loopdate2 = GRAPH_Future_Data(i + 1).DATE_DATA Then
                                        DateCheck = True
                                    Else
                                        DateCheck = False
                                        Exit For
                                    End If
                                End If
                            Else
                                INS_FUTURE_DATE = CStr(loopdate2)
                                INS_FUTURE_STOCK_NUM1 = INS_FUTURE_STOCK_NUM1
                                INS_FUTURE_STOCK_NUM2 = INS_FUTURE_STOCK_NUM2
                            End If
                        Next
                    Else
                        '表示データがない場合、ここでは日付を設定する。
                        INS_FUTURE_DATE = CStr(loopdate2)
                    End If

                    dtRow = ds.Tables(0).NewRow
                    dtRow(0) = INS_FUTURE_DATE
                    If FirstData2 = True Then
                        dtRow(1) = INS_FUTURE_STOCK_NUM1
                        dtRow(2) = INS_FUTURE_STOCK_NUM2
                    Else
                        dtRow(1) = Stock_NUM
                        dtRow(2) = Stock_NUM
                    End If

                    ds.Tables(0).Rows.Add(dtRow)
                    CheckDataCount2 += 1
                    loopdate2 = loopdate2.AddDays(1).ToString("yyyy/MM/dd")
                Loop

                'グラフの表示設定
                With Chart1
                    .Series.Clear()         '系列を初期化
                    .DataSource = ds        'Chart に表示するデータソースを設定
                    Dim colum As Integer = ds.Tables(0).Columns.Count - 1   'データの系列数を取得
                    For i As Integer = 1 To colum
                        Dim columnName As String = ds.Tables(0).Columns(i).ColumnName.ToString()
                        '系列の設定(発注・在庫数・在庫数＋発注数)
                        .Series.Add(columnName)
                        'グラフの種類を折れ線グラフに設定
                        .Series(columnName).ChartType = DataVisualization.Charting.SeriesChartType.Line
                        'X 軸のラベルテキストの読込・設定(日付）
                        .Series(columnName).XValueMember = ds.Tables(0).Columns(0).ColumnName.ToString()
                        '軸の太さ
                        .Series(columnName).BorderWidth = 3
                        'グラフ用のデータの読込・設定
                        .Series(columnName).YValueMembers = columnName
                        '凡例の表示設定
                        '.Legends(0).Enabled = False

                        '凡例をグラフ上部に表示
                        .Legends(0).Alignment = StringAlignment.Near
                        .Legends(0).Docking = Docking.Top
                    Next

                    '今日の日付にポイントを打つ為の点グラフの設定
                    .Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    'ポイントのサイズ
                    .Series(2).MarkerSize = 10

                End With
                '作成した画像イメージを保存する。
                Using st As New System.IO.FileStream(GraphPath + Item_Data(0).I_CODE + ImageType, System.IO.FileMode.Create)
                    Chart1.SaveImage(st, ChartImageFormat.Png)
                End Using

                '商品コード設定
                AxReport1.Item("", "I_CODE").Text = " " & Item_Data(0).I_CODE
                '商品名の設定
                AxReport1.Item("", "I_NAME").Text = " " & Item_Data(0).I_NAME

                '参照期間の設定
                'AxReport1.Item("", "PERIOD").Text = DateFrom & "～" & DateTo
                'プロダクトラインの設定
                AxReport1.Item("", "PL_NAME").Text = " " & Item_Data(0).PL_NAME

                '画像を設定
                AxReport1.Item("", "Image1").SetImagePath(GraphPath + Item_Data(0).I_CODE + ImageType)
                '現在庫数の設定
                AxReport1.Item("", "Stock").Text = Stock_NUM & " "
                'フッターの出力日時の設定
                AxReport1.Item("", "Date").Text = LabelDatetime

                '発注データのアラート表示欄クリア
                AxReport1.Item("", "alert2").Text = ""
                '入出庫データのアラート表示欄クリア
                AxReport1.Item("", "alert3").Text = ""

                '発注データ、合計欄クリア
                AxReport1.Item("", "total2").Text = ""
                '入出庫データ、合計欄クリア
                AxReport1.Item("", "total3").Text = ""


                '過去在庫履歴表示枠のクリア
                For i = 0 To 35
                    '日付の設定
                    If i < 5 Then
                        AxReport1.Item("", "Data1_" & i + 1 & "-1").Text = ""
                    End If
                    '種別の設定
                    'AxReport1.Item("", "Data1_" & i + 1 & "-2").Text = ""
                    '数量の設定
                    AxReport1.Item("", "Data1_" & i + 1 & "-3").Text = ""
                    '在庫数の設定
                    'AxReport1.Item("", "Data1_" & i + 1 & "-4").Text = ""
                Next

                '発注予定表示枠、入出荷表示枠のクリア
                For i = 0 To 9
                    '発注予定表示枠のクリア
                    '日付の設定
                    AxReport1.Item("", "Data2_" & i + 1 & "-1").Text = ""
                    '種別の設定
                    AxReport1.Item("", "Data2_" & i + 1 & "-2").Text = ""
                    '数量の設定
                    AxReport1.Item("", "Data2_" & i + 1 & "-3").Text = ""
                    '在庫数の設定
                    AxReport1.Item("", "Data2_" & i + 1 & "-4").Text = ""
                Next

                For i = 0 To 9
                    '入出荷表示枠のクリア
                    '日付の設定
                    AxReport1.Item("", "Data3_" & i + 1 & "-1").Text = ""
                    '種別の設定
                    AxReport1.Item("", "Data3_" & i + 1 & "-2").Text = ""
                    '数量の設定
                    AxReport1.Item("", "Data3_" & i + 1 & "-3").Text = ""
                    '在庫数の設定
                    AxReport1.Item("", "Data3_" & i + 1 & "-4").Text = ""
                Next

                '過去六ヶ月のサマリーを表示

                '月を設定
                AxReport1.Item("", "Data1_1-1").Text = dtNow.AddMonths(0).ToString("yyyy/MM")
                AxReport1.Item("", "Data1_2-1").Text = dtNow.AddMonths(-1).ToString("yyyy/MM")
                AxReport1.Item("", "Data1_3-1").Text = dtNow.AddMonths(-2).ToString("yyyy/MM")
                AxReport1.Item("", "Data1_4-1").Text = dtNow.AddMonths(-3).ToString("yyyy/MM")
                AxReport1.Item("", "Data1_5-1").Text = dtNow.AddMonths(-4).ToString("yyyy/MM")
                AxReport1.Item("", "Data1_6-1").Text = dtNow.AddMonths(-5).ToString("yyyy/MM")

                j = 1
                '数量を設定
                For i = 0 To 35 Step 6
                    '入庫確定数を表示
                    AxReport1.Item("", "Data1_" & i + 1 & "-3").Text = GRAPH_SummaryData(i).NUM & " "
                    '出庫確定数を表示(０じゃなければマイナス表示)
                    If GRAPH_SummaryData(i + 1).NUM = 0 Then
                        AxReport1.Item("", "Data1_" & i + 2 & "-3").Text = GRAPH_SummaryData(i + 1).NUM & " "
                    Else
                        AxReport1.Item("", "Data1_" & i + 2 & "-3").Text = "-" & GRAPH_SummaryData(i + 1).NUM & " "
                    End If

                    '棚卸数を表示
                    AxReport1.Item("", "Data1_" & i + 3 & "-3").Text = GRAPH_SummaryData(i + 2).NUM & " "
                    '在庫調整数を表示
                    AxReport1.Item("", "Data1_" & i + 4 & "-3").Text = GRAPH_SummaryData(i + 3).NUM & " "
                    'セット組数を表示
                    AxReport1.Item("", "Data1_" & i + 5 & "-3").Text = GRAPH_SummaryData(i + 4).NUM & " "
                    'セットばらし数を表示
                    AxReport1.Item("", "Data1_" & i + 6 & "-3").Text = GRAPH_SummaryData(i + 5).NUM & " "

                    '合計を表示
                    AxReport1.Item("", "RIREKI_TOTAL" & j).Text = GRAPH_SummaryData(i).NUM - GRAPH_SummaryData(i + 1).NUM + _
                    GRAPH_SummaryData(i + 2).NUM + GRAPH_SummaryData(i + 3).NUM + GRAPH_SummaryData(i + 4).NUM + GRAPH_SummaryData(i + 5).NUM & " "

                    j += 1
                Next

                LoopCount = 0
                If GRAPH_Future_Data_Count <> 0 Then
                    '発注予定を設定
                    For i = 0 To GRAPH_Future_Data.Length - 1
                        If i < 9 Then
                            If GRAPH_Future_Data(i).TYPE = "発注" Then

                                'グラフに表示している期間は通常表示
                                '発注日の設定
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-1").FontColor = RGB(0, 0, 0)
                                If DateTo > GRAPH_Future_Data(i).DATE_DATA Then

                                    AxReport1.Item("", "Data2_" & LoopCount + 1 & "-1").Text = GRAPH_Future_Data(i).DATE_DATA
                                Else
                                    'グラフに表示できない期間は前にアスタリスクを表示。
                                    AxReport1.Item("", "Data2_" & LoopCount + 1 & "-1").Text = "*" & GRAPH_Future_Data(i).DATE_DATA
                                End If
                                '希望納期の設定
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-2").FontColor = RGB(0, 0, 0)
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-2").Text = GRAPH_Future_Data(i).ORDER_DATE
                                '数量の設定
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-3").FontColor = RGB(0, 0, 0)
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-3").Text = GRAPH_Future_Data(i).NUM & " "
                                '在庫数の設定
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-4").FontColor = RGB(0, 0, 0)
                                AxReport1.Item("", "Data2_" & LoopCount + 1 & "-4").Text = GRAPH_Future_Data(i).STOCK_ORDER_NUM & " "

                                '発注数の合計を計算
                                PO_Total = PO_Total + GRAPH_Future_Data(i).NUM

                                LoopCount += 1
                            End If
                        End If
                    Next

                    LoopCount = 0
                    '入出荷予定を設定
                    For i = 0 To GRAPH_Future_Data.Length - 1
                        If i < 9 Then
                            If GRAPH_Future_Data(i).TYPE = "入庫予定" Or GRAPH_Future_Data(i).TYPE = "出庫予定" Then

                                'グラフに表示している期間は通常表示
                                '日付の設定
                                AxReport1.Item("", "Data3_" & LoopCount + 1 & "-1").FontColor = RGB(0, 0, 0)
                                If DateTo > GRAPH_Future_Data(i).DATE_DATA Then
                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-1").Text = GRAPH_Future_Data(i).DATE_DATA
                                Else
                                    'グラフに表示できない期間は前にアスタリスクを表示。
                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-1").Text = "*" & GRAPH_Future_Data(i).DATE_DATA
                                End If
                                '種別の設定
                                AxReport1.Item("", "Data3_" & LoopCount + 1 & "-2").FontColor = RGB(0, 0, 0)
                                AxReport1.Item("", "Data3_" & LoopCount + 1 & "-2").Text = GRAPH_Future_Data(i).TYPE
                                '数量の設定
                                AxReport1.Item("", "Data3_" & LoopCount + 1 & "-3").FontColor = RGB(0, 0, 0)

                                If GRAPH_Future_Data(i).TYPE = "入庫予定" Then
                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-3").Text = GRAPH_Future_Data(i).NUM & " "
                                    '入出庫の合計を計算
                                    INOUT_Total = INOUT_Total + GRAPH_Future_Data(i).NUM & " "
                                ElseIf GRAPH_Future_Data(i).TYPE = "出庫予定" Then

                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-3").Text = "-" & GRAPH_Future_Data(i).NUM.ToString & " "
                                    '入出庫の合計を計算
                                    INOUT_Total = INOUT_Total - GRAPH_Future_Data(i).NUM & " "
                                End If

                                '在庫数の設定
                                If Today > GRAPH_Future_Data(i).DATE_DATA Then
                                    '過去のデータは在庫数の表示は行わない。
                                Else
                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-4").FontColor = RGB(0, 0, 0)
                                    AxReport1.Item("", "Data3_" & LoopCount + 1 & "-4").Text = GRAPH_Future_Data(i).STOCK_ORDER_NUM & " "

                                End If
                                LoopCount += 1

                            End If
                        End If
                    Next

                    '発注データの合計、入出庫データの合計を設定
                    AxReport1.Item("", "total2").Text = PO_Total & " "
                    AxReport1.Item("", "total3").Text = INOUT_Total & " "
                End If
                'レポートの印刷 
                If AxReport1.PrintReport() = False Then
                    'エラー処理
                    MsgBox("印刷時にエラーが発生しました。")
                    '処理中ウインドウを表示
                    Processing.Show()
                    Exit Sub
                End If
                System.IO.File.Delete(GraphPath + Item_Data(0).I_CODE + ImageType)
            Else
                'グラフの作成を行わないデータの商品コードを格納
                MessageList = MessageList & Item_Data(0).I_CODE & vbCr
                MessageFlg = False
            End If
        Next
        '処理中ウインドウをDispose。
        Processing.Dispose()

        If MessageFlg = False Then
            'グラフの作成を行わないデータの商品コードリストを表示
            MsgBox(MessageList)
        End If

        '印刷ＪＯＢの終了（ファイルを閉じる） 
        AxReport1.ClosePrintJob(True)

        'レポーを閉じる 
        AxReport1.ReportPath = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim loopcnt As Integer = DataGridView1.Rows.Count
        Dim Count As Integer = 0
        Dim IndexCount As Integer = 10
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
        Dim IndexCount As Integer = 10
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

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '選択されていればSelectedValueに入っているのでPLACE_IDに格納
        If ComboBox4.SelectedIndex <> -1 Then
            '格納
            PLACE_ID = ComboBox4.SelectedValue.ToString()
        End If
    End Sub
End Class