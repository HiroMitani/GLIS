Imports System.Windows.Forms
Imports System.Globalization
Imports System.IO

Public Class syukodenpyou

    Dim PLACE_ID As String

    '出荷伝票の出力先
    Public CSVPath As String = System.Configuration.ConfigurationManager.AppSettings("CSVPath")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim FileNameCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing
        Dim SlipList() As Slip_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        Dim Data_Total As Integer = 0
        Dim Data_Num_Total As Integer = 0

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()

        If Trim(TextBox1.Text) = "" Then
            MsgBox("出荷指示ファイル名を入力してください。")
            Exit Sub
        End If

        If StringChkVal(Trim(TextBox1.Text), True, False, FileNameCheckAfter, StringChkResult, StringErrorMessage) = False Then
            TextBox1.BackColor = Color.Salmon
            MsgBox("出荷指示ファイル名に不正な文字が入力されています。")
        Else
            'チェックに問題がなければ背景色を白に戻す。
            TextBox1.BackColor = Color.White
        End If

        '出荷伝票検索Function
        Result = GetSlipSearch(FileNameCheckAfter, PLACE_ID, SlipList, Data_Total, Data_Num_Total, Result, ErrorMessage)

        If Result = False Then
            Label14.Text = "商品数："
            Label15.Text = "総数："
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '件数を表示
        Label14.Text = "商品数：" & Data_Total
        '総商品数を表示
        Label15.Text = "総数：" & Data_Num_Total

        For i = 0 To SlipList.Length - 1
            Dim Slip_List As New DataGridViewRow
            Slip_List.CreateCells(DataGridView1)
            With Slip_List
                '出荷指示ファイル名
                .Cells(0).Value = SlipList(i).FILE_NAME
                '伝票番号
                .Cells(1).Value = SlipList(i).SHEET_NO
                'オーダー番号
                .Cells(2).Value = SlipList(i).ORDER_NO
                '商品ID
                .Cells(12).Value = SlipList(i).I_ID
                '商品コード
                .Cells(3).Value = SlipList(i).I_CODE
                '商品名
                .Cells(4).Value = SlipList(i).I_NAME
                '納品先ID
                .Cells(13).Value = SlipList(i).C_ID
                '納品先コード
                .Cells(5).Value = SlipList(i).C_CODE
                '納品先名
                .Cells(6).Value = SlipList(i).C_NAME
                '納入単価
                .Cells(7).Value = SlipList(i).UNIT_COST
                '数量
                .Cells(8).Value = SlipList(i).NUM
                'コメント１
                .Cells(9).Value = SlipList(i).COMMENT1
                'コメント２
                .Cells(10).Value = SlipList(i).COMMENT2
                '出荷予定日
                .Cells(11).Value = SlipList(i).O_DATE
            End With
            DataGridView1.Rows.Add(Slip_List)
        Next

    End Sub

    Private Sub ｓyukodenpyou_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "出荷伝票出力"

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

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim Slip_Check_Flg As Boolean = False
        Dim Slip_Data_Count As Integer = 0
        Dim Slip_Check() As SlipID_List = Nothing
        Dim Slip_List() As SlipID_List = Nothing
        Dim Slip_Data_List() As Slip_List = Nothing
        Dim Count As Integer = 0
        Dim Slip_Result As Boolean = True
        Dim Slip_ErrorMessage As String = Nothing
        Dim DataCount As Integer = 0
        Dim LineData As String = Nothing

        Dim Slip_Num As Integer = 0
        Dim Slip_Max_Num As Integer = 0
        Dim PageCount As Integer = 1
        Dim OrderYear As String
        Dim OrderMonth As String
        Dim OrderDay As String
        Dim D_YMD As Date
        Dim Delivery() As String
        Dim Y2YMD As String
        Dim Kanseki_Check As Boolean = False
        Dim Pfj_Check As Boolean = False
        Dim Takamiya_Check As Boolean = False
        Dim Csv_Complete_Message As String = Nothing
        Dim Total As Integer = 0
        Dim PageCounter As Integer = 1
        Dim SamePageCount As Integer = 1
        Dim MaxPagecount As Integer = 1
        Dim TempStr As String = Nothing
        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter

        Dim FileNameCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing

        Dim OutFileName As String = Nothing
        Dim CSVFileName As String = Nothing

        If Trim(TextBox1.Text) = "" Then
            MsgBox("出荷指示ファイル名を入力してください。")
            Exit Sub
        End If

        If StringChkVal(Trim(TextBox1.Text), True, False, FileNameCheckAfter, StringChkResult, StringErrorMessage) = False Then
            TextBox1.BackColor = Color.Salmon
            MsgBox("出荷指示ファイル名に不正な文字が入力されています。")
        Else
            'チェックに問題がなければ背景色を白に戻す。
            TextBox1.BackColor = Color.White
        End If

        '発注日をそれぞれ年月日にわける。
        OrderYear = DateTimePicker1.Value.ToString("yy")
        OrderMonth = DateTimePicker1.Value.ToString("MM")
        OrderDay = DateTimePicker1.Value.ToString("dd")

        D_YMD = DateTimePicker1.Value.ToString("yyyy/MM/dd")
        '納品日は発注日＋１にしたものを格納
        D_YMD = D_YMD.AddDays(1)
        Y2YMD = D_YMD.ToString("yy/MM/dd")
        Delivery = CStr(Y2YMD).Split("/")

        '検索していなかったらエラーメッセージ表示
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("検索を行ってからCSV出力ボタンを押してください。")
            Exit Sub
        End If

        'DataGridViewの１行目の出荷指示ファイル名を取得して
        '出力するCSVファイル名にする。
        OutFileName = DataGridView1.Rows(0).Cells(0).Value()

        'ファイル名が29文字を超えていたら
        '頭から29文字のみ使用する。
        If OutFileName.Length > 29 Then
            CSVFileName = OutFileName.Substring(0, 28)
        Else
            CSVFileName = OutFileName
        End If

        'カンセキ伝票のファイル名
        'Dim Sheet1_Name As String = "納品書" & dtNow.ToString("yyyyMMddHHmm") & "（株）カンセキ.csv"
        '2012/01/23 菊池様依頼によりファイル名を出荷指示ファイル名＋伝票タイプに変更
        Dim Sheet1_Name As String = CSVFileName & "-1.csv"
        'カンセキ伝票のHeader設定
        Dim Sheet1_Header As String = "伝票番号,CUST_NAME,納品先名,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,製品コード,製品名,JAN,U,V,W,X,Y,出荷数,仕切価格,AB,小売価格,AD,AE,AF,AG,備考,伝票枚数ID,伝票総枚数,CUST_CD,コメント2"
        'PFJ伝票のファイル名
        'Dim Sheet2_Name As String = "納品書" & dtNow.ToString("yyyyMMddHHmm") & "（ＰＦＪ）.csv"
        '2012/01/23 菊池様依頼によりファイル名を出荷指示ファイル名＋伝票タイプに変更
        Dim Sheet2_Name As String = CSVFileName & "-3.csv"
        'PFJ伝票のHeader設定
        Dim Sheet2_Header As String = "伝票番号,CUST_NAME,納品先名,納品先ID,E,F,G,H,I,J,K,L,M,N,O,P,Q,製品コード,製品名,JAN,U,V,W,X,Y,出荷数,仕切価格,AB,小売価格,AD,AE,AF,AG,備考,伝票枚数ID,伝票総枚数,CUST_CD,コメント2"
        'タカミヤ伝票ファイル名 
        'Dim Sheet3_Name As String = "納品書" & dtNow.ToString("yyyyMMddHHmm") & "タカミヤ.csv"
        '2012/01/23 菊池様依頼によりファイル名を出荷指示ファイル名＋伝票タイプに変更
        Dim Sheet3_Name As String = CSVFileName & "-T.csv"
        'タカミヤ伝票のHeader設定
        Dim Sheet3_Header As String = "伝票番号,発注年,発注月,発注日,出荷年,出荷月,出荷日,B,C,D,納入先住所,納品先名,発注者,納品先ID,店舗ｺｰﾄﾞ,伝票区分,取引先ｺｰﾄﾞ,製品コード,製品名,JAN,出荷数,訂正後数量,仕切価格,納入金額,小売価格,備考,数量合計,E,F,納入金額合計,H,I,J,35,コメント2"

        '文字コード設定
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        '出荷指示ファイルで検索されていることが条件なので
        '検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        '出荷指示ファイル名と同じであることをチェックする。
        If Trim(TextBox1.Text) = "" Then
            MsgBox("出荷伝票CSVを出力するには出荷指示ファイルでの検索が必須となります。")
            Exit Sub
        End If

        DataCount = 0
        Slip_Data_List = Nothing
        Slip_Result = True
        Slip_ErrorMessage = Nothing
        'カンセキ伝票のデータを取得
        Slip_Result = GetSlipCancelData(OutFileName, 1, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)

        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet1_Name, False, strEncoding)

            PageCount = 1
            'headerを設定
            strStreamWriter.WriteLine(Sheet1_Header)
            'データ件数分ループ
            For i = 0 To Slip_Data_List.Length - 1

                '取得するデータをLineataに格納する。
                'A列に伝票番号
                LineData = """" & Slip_Data_List(i).SHEET_NO & ""","
                'B列に企業名
                TempStr = Slip_Data_List(i).C_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'C列に納品先名
                TempStr = Slip_Data_List(i).D_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'D列は企業コードの６桁目以降を表示
                LineData &= """" & Slip_Data_List(i).C_CODE.Substring(5) & ""","
                'E列にオーダー番号：
                LineData &= """" & Slip_Data_List(i).ORDER_NO & ""","
                'F列は1固定（分類コード？）
                LineData &= "1,"
                'G列は31固定（伝票区分？）
                LineData &= "31,"
                'H列にPFJの社名
                LineData &= """" & Com_NAME & ""","
                'I列にPFJのTEL番号
                LineData &= """" & Com_TEL & ""","
                'J列にPFJのFAX番号
                LineData &= """" & Com_FAX & ""","
                'K列に発注年
                LineData &= OrderYear & ","
                'L列に発注月
                LineData &= OrderMonth & ","
                'M列に発注日
                LineData &= OrderDay & ","
                'N列に納品年
                LineData &= Delivery(0) & ","
                'O列は納品月
                LineData &= Delivery(1) & ","
                'P列は納品日
                LineData &= Delivery(2) & ","
                'Q列は627603固定
                LineData &= "627603,"
                'R列に製品コード
                LineData &= """" & Slip_Data_List(i).I_CODE & ""","
                'S列に製品名
                TempStr = Slip_Data_List(i).I_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'T列にJANコード
                LineData &= """" & Slip_Data_List(i).JAN & ""","
                'U列は設定なし（項目名：U)
                LineData &= ","
                'V列は設定なし（項目名：V）
                LineData &= ","
                'W列は設定なし（項目名：W）
                LineData &= ","
                'X列は設定なし（項目名：X）
                LineData &= ","
                'Y列は設定なし（項目名：Y)
                LineData &= ","
                'Z列は出荷数
                LineData &= Slip_Data_List(i).NUM & ","
                'AA列は仕切価格
                LineData &= Slip_Data_List(i).UNIT_COST & ","
                'AB列は設定なし（項目名：AB）
                LineData &= ","
                'AC列は設小売価格
                LineData &= Slip_Data_List(i).PRICE & ","
                'AD列は設定なし（項目名：AD）
                LineData &= ","
                'AE列は設定なし（項目名：AE）
                LineData &= ","
                'AF列は設定なし（項目名：AF）
                LineData &= ","
                'AG列は設定なし（項目名：AG）
                LineData &= ","
                'AH列は備考
                TempStr = Slip_Data_List(i).COMMENT1
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'AI列は伝票枚数ID
                LineData &= Slip_Data_List(i).PAGE & ","
                'AJ列は伝票総枚数
                LineData &= Slip_Data_List(i).MAXPAGE & ","
                'AK列はCUST_CD
                LineData &= """" & Slip_Data_List(i).C_CODE & ""","
                'AL列は設定なし
                LineData &= ","
                '1行にしたデータを登録
                strStreamWriter.WriteLine(LineData)
                'PageCount += 1

            Next
            'カンセキ伝票作成フラグ
            Kanseki_Check = True
            'ファイルを閉じる
            strStreamWriter.Close()
        End If


        'PFJ
        DataCount = 0
        Slip_Data_List = Nothing
        Slip_Result = True
        Slip_ErrorMessage = Nothing

        MaxPagecount = 1
        SamePageCount = 1
        PageCounter = 1

        'PFJ伝票のデータを取得
        Slip_Result = GetSlipCancelData(OutFileName, 2, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)
        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet2_Name, False, strEncoding)

            PageCount = 1
            'headerを設定
            strStreamWriter.WriteLine(Sheet2_Header)
            'データ件数分ループ
            For i = 0 To Slip_Data_List.Length - 1

                If i <> 0 Then
                    '2件目以降、１つ上のデータとチェックして違う伝票番号なら、ページ数をカウントアップする。
                    If Slip_Data_List(i).SHEET_NO <> Slip_Data_List(i - 1).SHEET_NO Then
                        PageCounter += 1
                        SamePageCount = 1
                    Else
                        If SamePageCount = 11 Then
                            'また、同じ伝票番号が10件以上続いた場合もページ数をカウントアップ。
                            PageCounter += 1
                            SamePageCount = 1
                        Else
                            SamePageCount += 1
                        End If
                    End If

                End If

                'End If
                '取得するデータをLineataに格納する。
                'A列に伝票番号
                LineData = """" & Slip_Data_List(i).SHEET_NO & ""","
                'B列に企業名
                TempStr = Slip_Data_List(i).C_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'C列に納品先名
                TempStr = Slip_Data_List(i).D_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'D列は納品先コード
                LineData &= """" & Slip_Data_List(i).C_CODE & ""","
                'E列にオーダー番号：
                LineData &= """" & Slip_Data_List(i).ORDER_NO & ""","
                'F列は1固定（分類コード？）
                LineData &= "1,"
                'G列は31固定（伝票区分？）
                LineData &= "31,"
                'H列にPFJの社名
                LineData &= """" & Com_NAME & ""","
                'I列にPFJのTEL番号
                LineData &= """" & Com_TEL & ""","
                'J列にPFJのFAX番号
                LineData &= """" & Com_FAX & ""","
                'K列に発注年
                LineData &= OrderYear & ","
                'L列に発注月
                LineData &= OrderMonth & ","
                'M列に発注日
                LineData &= OrderDay & ","
                'N列に納品年
                LineData &= Delivery(0) & ","
                'O列は納品月
                LineData &= Delivery(1) & ","
                'P列は納品日
                LineData &= Delivery(2) & ","
                'Q列は627603固定
                LineData &= "627603,"
                'R列に製品コード
                LineData &= """" & Slip_Data_List(i).I_CODE & ""","
                'S列に製品名
                TempStr = Slip_Data_List(i).I_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'T列にJANコード
                LineData &= """" & Slip_Data_List(i).JAN & ""","
                'U列は設定なし（項目名：U)
                LineData &= ","
                'V列は設定なし（項目名：V）
                LineData &= ","
                'W列は設定なし（項目名：W）
                LineData &= ","
                'X列は設定なし（項目名：X）
                LineData &= ","
                'Y列は設定なし（項目名：Y)
                LineData &= ","
                'Z列は出荷数
                LineData &= Slip_Data_List(i).NUM & ","
                'AA列は仕切価格
                LineData &= Slip_Data_List(i).UNIT_COST & ","
                'AB列は設定なし（項目名：AB）
                LineData &= ","
                'AC列は設小売価格
                LineData &= Slip_Data_List(i).PRICE & ","
                'AD列は設定なし（項目名：AD）
                LineData &= ","
                'AE列は設定なし（項目名：AE）
                LineData &= ","
                'AF列は設定なし（項目名：AF）
                LineData &= ","
                'AG列は設定なし（項目名：AG）
                LineData &= ","
                'AH列は備考
                TempStr = Slip_Data_List(i).COMMENT1
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'AI列は伝票枚数ID
                LineData &= Slip_Data_List(i).PAGE & ","
                'AJ列は伝票総枚数
                LineData &= Slip_Data_List(i).MAXPAGE & ","
                'AK列はCUST_CD
                LineData &= """" & Slip_Data_List(i).C_CODE & ""","
                'AL列は設定なし
                LineData &= ","
                '1行にしたデータを登録
                strStreamWriter.WriteLine(LineData)

                'PageCount += 1

            Next
            'Pfj伝票作成フラグ
            Pfj_Check = True
            'ファイルを閉じる
            strStreamWriter.Close()
        End If


        'タカミヤ伝票のデータを取得
        Slip_Result = GetSlipCancelData(OutFileName, 3, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)
        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(CSVPath & Sheet3_Name, False, strEncoding)

            'headerを設定
            strStreamWriter.WriteLine(Sheet3_Header)
            'データ件数分ループ
            For i = 0 To Slip_Data_List.Length - 1
                '取得するデータをLineataに格納する。
                'A列に伝票番号(下6桁を表示)
                LineData = """" & Strings.Right(Slip_Data_List(i).SHEET_NO, 6) & ""","
                'B列に発注年（出荷予定日の年）
                LineData &= OrderYear & ","
                'C列に発注月（出荷予定日の月）
                LineData &= OrderMonth & ","
                'D列に発注日（出荷予定日の日）
                LineData &= OrderDay & ","
                'E列に出荷年
                LineData &= Delivery(0) & ","
                'F列に出荷月
                LineData &= Delivery(1) & ","
                'G列に出荷日
                LineData &= Delivery(2) & ","
                'H列は設定なし
                LineData &= ","
                'I列は設定なし
                LineData &= ","
                'J列は30固定
                LineData &= "30,"
                'K列に納品先住所
                LineData &= """" & Slip_Data_List(i).D_ADDRESS & ""","
                'L列に納品先名
                TempStr = Slip_Data_List(i).D_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'M列に発注者
                TempStr = Slip_Data_List(i).C_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'N列に納品先ID(企業コード）
                LineData &= """" & Slip_Data_List(i).C_CODE & ""","
                'O列は設定なし（項目名：店舗コード）
                LineData &= ","
                'P列は1固定（伝票区分）
                LineData &= "1,"
                'Q列は271608固定（取引先コード）
                LineData &= "271608,"
                'R列に製品コード
                LineData &= """" & Slip_Data_List(i).I_CODE & ""","
                'S列に製品名
                TempStr = Slip_Data_List(i).I_NAME
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'T列にJANコード
                LineData &= """" & Slip_Data_List(i).JAN & ""","
                'U列に出荷数
                LineData &= Slip_Data_List(i).NUM & ","
                'V列は設定なし（項目名：訂正後数量）
                LineData &= ","
                'W列に仕切価格
                LineData &= Slip_Data_List(i).UNIT_COST & ","
                'X列は設定なし（項目名：納入金額）
                LineData &= ","
                'Y列に小売価格
                LineData &= Slip_Data_List(i).PRICE & ","
                'Z列は設定なし（項目名：備考）
                LineData &= ","
                'AA列は設定なし（項目名：数量合計）
                LineData &= ","
                'AB列にコメント１
                TempStr = Slip_Data_List(i).COMMENT1
                LineData &= """" & TempStr.Replace("""", """""") & ""","
                'AC列は設定なし（項目名：F)
                LineData &= ","
                'AD列は設定なし（納入金額合計）
                LineData &= ","
                'AE列にPFJの社名
                LineData &= Com_NAME & ","
                'AF列にPFJのTEL番号
                LineData &= Com_TEL & ","
                'AG列にPFJのFAX番号
                LineData &= Com_FAX & ","
                'AH列はT固定（項目名：35）
                LineData &= "T,"
                'AI列は設定なし（項目名：コメント２）
                LineData &= ","

                '1行にしたデータを登録
                strStreamWriter.WriteLine(LineData)
            Next
            'Pfj伝票作成フラグ
            Takamiya_Check = True
            'ファイルを閉じる
            strStreamWriter.Close()
        End If

        Csv_Complete_Message = CSVPath & "に" & vbCr & "以下の出荷伝票CSVの作成が完了しました。"
        '2013/2/27　原さんの依頼により、メッセージ変更
        'If Kanseki_Check = False Then
        '    Csv_Complete_Message &= vbCr & "データがない為、カンセキ伝票は作成されませんでした。"
        'End If
        'If Pfj_Check = False Then
        '    Csv_Complete_Message &= vbCr & "データがない為、TPFJ伝票は作成されませんでした。"
        'End If
        'If Takamiya_Check = False Then
        '    Csv_Complete_Message &= vbCr & "データがない為、タカミヤ伝票は作成されませんでした。"
        'End If
        If Kanseki_Check = True Then
            Csv_Complete_Message &= vbCr & "カンセキ伝票が作成されました。"
        End If
        If Pfj_Check = True Then
            Csv_Complete_Message &= vbCr & "T3伝票が作成されました。"
        End If
        If Takamiya_Check = True Then
            Csv_Complete_Message &= vbCr & "タカミヤ伝票が作成されました。"
        End If

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