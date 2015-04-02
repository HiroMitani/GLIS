Imports System.IO
Imports System
Imports System.Text

Public Class syukoshiji

    Dim PLACE_ID As String

    'テンポラリーテーブルに登録する際のユニークキーを作成（日時）
    Dim dtNow As DateTime

    '出荷指示ファイル名格納
    Dim ExcelFileName As String

    '取り込んだエクセルの納品先コード、商品コードがマスタにない場合
    '登録を不可能にするためのBoolean（True：登録不可、False:登録可能）
    Dim Regist_Check_Flg As Boolean = False

    '取り込んだファイルの伝票番号がすでにDBに存在する場合、
    'アラートを表示するためのBoolean。（True：重複なし、False:重複あり）
    Dim OrderNo_Duplication_Check As Boolean = True

    '出荷伝票の出力先
    '開発
    Dim FilePath As String = "C:\Documents and Settings\Administrator\My Documents\"
    '本番
    'Dim FilePath As String = "C:\Users\PFJ\Documents\"

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        '出荷指示ファイル名が既に取り込まれているかチェックする為の変数
        Dim FNameErrorMessage As String = Nothing
        Dim FNameResult As Boolean = True

        Dim Out_Data() As Out_Regist_List = Nothing

        Dim FileName As String = Nothing

        Dim tmp_Comment1 As String = Nothing
        Dim tmp_Comment2 As String = Nothing

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("出荷指示ファイルを取り込んでから登録ボタンを押してください。")
            Exit Sub
        End If

        If Regist_Check_Flg = True Then
            MsgBox("取り込んだデータの中にマスタ登録されていない納品先、商品があります。")
            Exit Sub
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("出荷指示ファイルを取り込んでください。")
            Exit Sub
        End If

        '選択されたファイルがxlsかチェック
        Dim FileExtension As String = System.IO.Path.GetExtension(TextBox1.Text)
        'もし拡張子が.xlsじゃなかったらエラー終了
        If FileExtension <> ".xls" AndAlso FileExtension <> ".xlsx" Then
            MsgBox("エクセルファイルを指定してください。")
            Exit Sub
        End If

        FileName = Trim(TextBox1.Text)

        If FileExtension = ".xls" Then
            FileName = FileName.Replace(".xls", "")
        ElseIf FileExtension <> ".xlsx" Then
            FileName = FileName.Replace(".xlsx", "")
        End If

        '取り込もうとしている出荷指示ファイル名が、すでにDBに登録されているかチェックを行う。
        FNameResult = FileName_Check(FileName, FNameResult, FNameErrorMessage)

        If FNameResult = False Then
            MsgBox("ファイル名""" & TextBox1.Text & """は" & FNameErrorMessage)
            TextBox1.Text = ""
            Exit Sub
        End If

        '取り込もうとしている出荷指示ファイル名が、すでにDBに登録されているかチェックを行う。
        FNameResult = FileName_Check(FileName, FNameResult, FNameErrorMessage)

        If FNameResult = False Then
            MsgBox("ファイル名""" & TextBox1.Text & """は" & FNameErrorMessage)
            TextBox1.Text = ""
            Exit Sub
        End If

        If OrderNo_Duplication_Check = False Then

            '取り込んだデータに重複がある場合、ダイアログを表示。
            Dim Duplication_Message_Result As DialogResult = MessageBox.Show("同じ伝票番号で出荷指示を行おうとしています。" & vbCr & "出荷指示を行ってもよろしいですか？", _
                                         "確認", _
                                         MessageBoxButtons.YesNo, _
                                         MessageBoxIcon.Exclamation, _
                                         MessageBoxDefaultButton.Button2)

            '何が選択されたか調べる 
            If Duplication_Message_Result = DialogResult.No Then
                'Noを選択した場合、処理終了
                Exit Sub
            End If
        End If

        'DataGridViewの情報を配列に入れる

        ReDim Out_Data(DataGridView1.Rows.Count - 1)
        For Count = 0 To DataGridView1.Rows.Count - 1
            '処理ステータス
            If DataGridView1(1, Count).Value = "通常出荷指示データ" Then
                Out_Data(Count).STATUS = 1
            ElseIf DataGridView1(1, Count).Value = "重複通常出荷指示データ" Then
                Out_Data(Count).STATUS = 1
            ElseIf DataGridView1(1, Count).Value = "出荷指示キャンセル" Then
                Out_Data(Count).STATUS = 3
            ElseIf DataGridView1(1, Count).Value = "ピッキング戻し" Then
                Out_Data(Count).STATUS = 4
            End If

            'もし、伝票出力のみにチェックが入れられていたら、ステータスを５に修正。
            If CheckBox1.Checked = True Then
                Out_Data(Count).STATUS = 5
            End If

            '納品先ID
            Out_Data(Count).C_ID = DataGridView1(17, Count).Value
            'オーダー番号
            Out_Data(Count).ORDER_NO = DataGridView1(4, Count).Value
            '伝票番号
            Out_Data(Count).SHEET_NO = DataGridView1(5, Count).Value
            '商品ID
            Out_Data(Count).I_ID = DataGridView1(16, Count).Value
            '価格（納入単価）
            Out_Data(Count).UNIT_COST = DataGridView1(8, Count).Value
            '数量
            Out_Data(Count).NUM = DataGridView1(9, Count).Value
            '合計金額
            Out_Data(Count).TOTAL_AMOUNT = DataGridView1(10, Count).Value
            'コメント１（伝票用）
            tmp_Comment1 = Trim(DataGridView1(11, Count).Value)
            tmp_Comment1 = tmp_Comment1.Replace("'", "''")
            tmp_Comment1 = tmp_Comment1.Replace("\", "\\")
            Out_Data(Count).COMMENT1 = tmp_Comment1
            'コメント２
            tmp_Comment2 = Trim(DataGridView1(12, Count).Value)
            tmp_Comment2 = tmp_Comment2.Replace("'", "''")
            tmp_Comment2 = tmp_Comment2.Replace("\", "\\")
            Out_Data(Count).COMMENT2 = tmp_Comment2
            '出荷予定日
            Out_Data(Count).O_DATE = DataGridView1(13, Count).Value
            '売単価
            Out_Data(Count).PRICE = DataGridView1(14, Count).Value
            'OUT_TBL.ID
            Out_Data(Count).OUT_ID = DataGridView1(18, Count).Value
            '不良区分（出荷指示ファイルのデータは全て良品）
            Out_Data(Count).I_STATUS = 1
            'OUT_PRT.ID
            Out_Data(Count).ID = DataGridView1(15, Count).Value
        Next

        Result = Ins_Out_Instructions(Out_Data, FileName, PLACE_ID, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("出荷指示ファイルの取り込みが完了しました。")

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        '出荷指示ファイル名をクリアする。
        TextBox1.Text = ""

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'syukoshijiDialog1.Show()

        dtNow = DateTime.Now

        'Excelのデータを格納するための配列
        Dim ExcelData() As ExcelData_List = Nothing

        '出荷指示ファイル名が既に取り込まれているかチェックする為の変数
        Dim FNameErrorMessage As String = Nothing
        Dim FNameResult As Boolean = True

        'ExcelファイルのデータをDBに登録できたかチェックする為の変数
        Dim ExcelErrorMessage As String = Nothing
        Dim ExcelResult As Boolean = True

        Dim CheckErrorMessage As String = Nothing
        Dim CheckResult As Boolean = True
        Dim CheckData() As Check_Rsult_List = Nothing

        Dim Sheet_No_List() As Duplication_Num_List = Nothing

        Dim Duplication_PL_ErrorMessageList As String = Nothing
        Dim Duplication_Zero_ErrorMessageList As String = Nothing

        Dim Out_Shiped_MessageList As String = Nothing
        Dim Out_NodataCancel_MessageList As String = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Cancel_Message As String = Nothing

        OrderNo_Duplication_Check = True

        Dim FileName As String = Nothing

        Dim ChkFileName As String = Nothing

        Dim DataErrorMessage As String = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        Dim ChkDateAfter As String = Nothing
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True

        Dim ChkComment1 As String = Nothing
        Dim ChkComment2 As String = Nothing

        Dim Tmp_Comment1 As String = Nothing
        Dim Tmp_Comment2 As String = Nothing

        Dim PrtOnlyData As Boolean = False

        ',オブジェクト変数の宣言
        Dim OBJ As Object

        '開くエクセルファイルのある、ディレクトリ、ファイル名を格納
        Dim fnm As String

        ' OpenFileDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
        Dim OpenFileDialog1 As New OpenFileDialog()

        ' ダイアログのタイトルを設定する
        OpenFileDialog1.Title = "ファイルを選択してください。"

        ' 初期表示するディレクトリを設定する
        OpenFileDialog1.InitialDirectory = "C:\Documents and Settings\Administrator\My Documents"

        ' 初期表示するファイル名を設定する
        OpenFileDialog1.FileName = ""

        ' ファイルのフィルタを設定する
        'OpenFileDialog1.Filter = "xlsファイル|*.xls;"
        OpenFileDialog1.Filter = "xlsxファイル|*.xlsx;|xlsファイル|*.xls;|すべてのファイル|*.*"

        ' ファイルの種類 の初期設定を 2 番目に設定する (初期値 1)
        OpenFileDialog1.FilterIndex = 1

        ' ダイアログボックスを閉じる前に現在のディレクトリを復元する (初期値 False)
        OpenFileDialog1.RestoreDirectory = True

        ' 複数のファイルを選択可能にする (初期値 False)
        OpenFileDialog1.Multiselect = False

        ' [ヘルプ] ボタンを表示する (初期値 False)
        OpenFileDialog1.ShowHelp = False

        ' [読み取り専用] チェックボックスを表示する (初期値 False)
        OpenFileDialog1.ShowReadOnly = False

        ' [読み取り専用] チェックボックスをオンにする (初期値 False)
        OpenFileDialog1.ReadOnlyChecked = False

        ' 存在しないファイルを指定した場合は警告を表示する (初期値 True)
        'OpenFileDialog1.CheckFileExists = True

        ' 存在しないパスを指定した場合は警告を表示する (初期値 True)
        'OpenFileDialog1.CheckPathExists = True

        ' 拡張子を指定しない場合は自動的に拡張子を付加する (初期値 True)
        'OpenFileDialog1.AddExtension = True

        ' 有効な Win32 ファイル名だけを受け入れるようにする (初期値 True)
        'OpenFileDialog1.ValidateNames = True

        ' ダイアログを表示し、キャンセルボタンが押された場合、処理終了
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Exit Sub
        End If

        '選択されたファイルがxlsかチェック
        Dim FileExtension As String = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
        'もし拡張子が.xlsじゃなかったらエラー終了
        If FileExtension <> ".xls" AndAlso FileExtension <> ".xlsx" Then
            MsgBox("エクセルファイルを指定してください。")
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        OBJ = GetObject(OpenFileDialog1.FileName) 'オブジェクト取得

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()

        'ファイル名を取得
        TextBox1.Text = Path.GetFileName(OpenFileDialog1.FileName)

        'ファイル名が513文字以上ならエラーメッセージを表示
        ChkFileName = Trim(TextBox1.Text)
        If ChkFileName.Length > 512 Then
            MsgBox("出荷指示ファイル名が長すぎます。")
            Exit Sub
        End If

        FileName = TextBox1.Text

        If FileExtension = ".xls" Then
            FileName = FileName.Replace(".xls", "")
        ElseIf FileExtension <> ".xlsx" Then
            FileName = FileName.Replace(".xlsx", "")
        End If

        '取り込もうとしている出荷指示ファイル名が、すでにDBに登録されているかチェックを行う。
        FNameResult = FileName_Check(FileName, FNameResult, FNameErrorMessage)

        If FNameResult = False Then
            MsgBox("ファイル名""" & TextBox1.Text & """は" & FNameErrorMessage)
            TextBox1.Text = ""
            Exit Sub
        End If

        'Excelファイルを読み込みDBに登録する為の変数設定
        '1行目は項目行の為、2に設定する。（1の場合、項目名も取得）
        Dim Count As Integer = 2
        'シート名取得
        Dim SheetName As String = OBJ.Worksheets(1).name

        'ExcelのデータをテンポラリーTableに取り込む。（データの不備等は登録後チェックを行う為、無条件で全件取り込む）
        Do While Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 1).value) <> ""
            '配列の再設定
            ReDim Preserve ExcelData(0 To Count - 2)

            '納品先コード
            ExcelData(Count - 2).C_CODE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value)
            '納品先コードの長さをチェック。
            If ExcelData(Count - 2).C_CODE.Length > 24 Then
                DataErrorMessage &= Count - 1 & "行目の納品先コードが長すぎます。" & vbCr
            End If

            'オーダー番号
            ExcelData(Count - 2).ORDER_NO = Trim(OBJ.Worksheets(SheetName).Cells(Count, 2).value)
            If ExcelData(Count - 2).ORDER_NO.Length > 20 Then
                DataErrorMessage &= Count - 1 & "行目のオーダー番号が長すぎます。" & vbCr
            End If

            '伝票番号
            ExcelData(Count - 2).SHEET_NO = Trim(OBJ.Worksheets(SheetName).Cells(Count, 3).value)
            If ExcelData(Count - 2).SHEET_NO.Length > 20 Then
                DataErrorMessage &= Count - 1 & "行目の伝票番号が長すぎます。" & vbCr
            End If
            '商品コード
            ExcelData(Count - 2).I_CODE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 4).value)
            If ExcelData(Count - 2).I_CODE.Length > 24 Then
                DataErrorMessage &= Count - 1 & "行目の商品コードが長すぎます。" & vbCr
            End If
            '納入単価
            If NumChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 5).value), "INTEGER", False, True, True, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataErrorMessage &= Count - 1 & "行目の納入単価が正しくありません。" & vbCr
            Else
                ExcelData(Count - 2).UNIT_COST = OBJ.Worksheets(SheetName).Cells(Count, 5).value
            End If

            '数量
            If NumChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 6).value), "INTEGER", False, False, True, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataErrorMessage &= Count - 1 & "行目の数量が正しくありません。" & vbCr
            Else
                ExcelData(Count - 2).NUM = OBJ.Worksheets(SheetName).Cells(Count, 6).value
            End If

            '合計金額
            If NumChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 7).value), "INTEGER", False, True, True, NumChkResult, NumChkErrorMessage) = False Then
                'MsgBox(NumChkErrorMessage)
                DataErrorMessage &= Count - 1 & "行目の合計金額が正しくありません。" & vbCr
            Else
                ExcelData(Count - 2).TOTAL_AMOUNT = OBJ.Worksheets(SheetName).Cells(Count, 7).value
            End If

            'コメント１
            ChkComment1 = Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value)

            If ChkComment1.Length > 401 Then
                DataErrorMessage &= Count - 1 & "行目のコメント1が長すぎます。" & vbCr
            Else
                '　アポストロフィ対応とバックスラッシュ対応
                Tmp_Comment1 = Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value)
                Tmp_Comment1 = Tmp_Comment1.Replace("'", "''")
                Tmp_Comment1 = Tmp_Comment1.Replace("\", "\\")
                ExcelData(Count - 2).COMMENT1 = Tmp_Comment1

            End If

            'コメント２
            ChkComment2 = Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value)
            If ChkComment2.Length > 401 Then
                DataErrorMessage &= Count - 1 & "行目のコメント2が長すぎます。" & vbCr
            Else
                Tmp_Comment2 = Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value)
                Tmp_Comment2 = Tmp_Comment2.Replace("'", "''")
                Tmp_Comment2 = Tmp_Comment2.Replace("\", "\\")
                ExcelData(Count - 2).COMMENT2 = Tmp_Comment2
            End If

            '出荷予定日
            If DateChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 10).value), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                DataErrorMessage &= Count + 1 & "行目の入荷予定日が正しくありません。" & vbCr
            Else
                '入荷予定日
                ExcelData(Count - 2).O_DATE = ChkDateAfter
            End If

            'ユニークキー
            ExcelData(Count - 2).KEY = dtNow.ToString("yyyyMMddHHmmss")

            Count += 1
        Loop

        'DataErrorMessageに値が入っていたら、エラー内容を表示
        If DataErrorMessage <> "" Then
            MsgBox(DataErrorMessage)
            Exit Sub
        End If

        If FileExtension = ".xls" Then
            FileName = FileName.Replace(".xls", "")
        ElseIf FileExtension <> ".xlsx" Then
            FileName = FileName.Replace(".xlsx", "")
        End If

        'OUT_PRTへ登録
        ExcelResult = Ins_ExcelData(ExcelData, FileName, PLACE_ID, ExcelResult, ExcelErrorMessage)

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If

        '上記で登録したデータで同一伝票内に数量がプラスとマイナスのものが存在するか、数量０のものが存在するかチェックを行う。
        Result = Sheet_No_Duplication_Check(FileName, dtNow.ToString("yyyyMMddHHmmss"), Sheet_No_List, Result, ErrorMessage)

        '数量がプラスとマイナスのものがあるかチェック
        For i = 0 To Sheet_No_List.Length - 1
            If Sheet_No_List(i).RESULT = 2 Then
                If Duplication_PL_ErrorMessageList = "" Then
                    Duplication_PL_ErrorMessageList = "伝票番号：" & vbCr & Sheet_No_List(i).SHEET_NO & vbCr
                Else
                    Duplication_PL_ErrorMessageList &= Sheet_No_List(i).SHEET_NO & vbCr
                End If
            End If
            If Sheet_No_List(i).RESULT = 3 Then
                If Duplication_Zero_ErrorMessageList = "" Then
                    Duplication_Zero_ErrorMessageList = "伝票番号：" & vbCr & Sheet_No_List(i).SHEET_NO & vbCr
                Else
                    Duplication_Zero_ErrorMessageList &= Sheet_No_List(i).SHEET_NO & vbCr
                End If
            End If
        Next

        If Duplication_PL_ErrorMessageList <> "" Then
            Duplication_PL_ErrorMessageList &= "は数量にプラスとマイナスのものが混在している為、出荷指示ファイルの取り込みを中止します。"
        End If
        If Duplication_Zero_ErrorMessageList <> "" Then
            Duplication_Zero_ErrorMessageList &= "は数量が0のものがある為、出荷指示ファイルの取り込みを中止します。"

        End If
        If Duplication_PL_ErrorMessageList <> "" Or Duplication_Zero_ErrorMessageList <> "" Then
            MsgBox(Duplication_PL_ErrorMessageList & vbCr & vbCr & Duplication_Zero_ErrorMessageList)
            'ファイル選択ボタンをクリックした時にユニークキーを作成している為、
            '再度指定させる為にファイル名表示欄をクリアする。
            TextBox1.Text = ""

            Exit Sub
        End If

        '伝票出力のみかチェック
        If CheckBox1.Checked = True Then
            PrtOnlyData = True
        End If

        'OUT_PRTへ登録したデータとOUT_TBLに存在するデータを取得し、データの確認を行う。
        CheckResult = GetCheckData(FileName, dtNow.ToString("yyyyMMddHHmmss"), PrtOnlyData, CheckData, CheckResult, CheckErrorMessage)

        If CheckResult = False Then
            MsgBox(CheckErrorMessage)
            Exit Sub
        End If

        '数量がプラスでステータスが出荷済みのもの。または、数量がマイナスで出荷テーブルにデータがなければメッセージを出して終了。
        'また、ピッキング済みデータの出荷指示キャンセルデータで、OUT_TBLの予定数量 - キャンセル数量 = 0でなければメッセージを出して終了。
        For Count = 0 To CheckData.Length - 1
            If CheckData(Count).OUTPRT_NUM < 0 And CheckData(Count).STATUS = "出荷済み" Or CheckData(Count).STATUS = "ピッキング戻し" Then
                'すでに出荷済みの場合

                Out_Shiped_MessageList &= "伝票番号：" & CheckData(Count).SHEET_NO & "、商品コード：" & CheckData(Count).I_CODE & vbCr
            ElseIf CheckData(Count).OUTPRT_NUM < 0 And CheckData(Count).STATUS = "0" Then
                '数量がマイナスでステータスが存在しない（データがない）場合

                Out_NodataCancel_MessageList &= "伝票番号：" & CheckData(Count).SHEET_NO & "、商品コード：" & CheckData(Count).I_CODE & vbCr

            End If

            'ピッキング済みのデータに対するキャンセル指示の数量チェック
            If CheckData(Count).STATUS = "ピッキング済み" And (CheckData(Count).OUTTBL_NUM + CheckData(Count).OUTPRT_NUM) <> 0 Then
                Cancel_Message = "伝票番号：" & CheckData(Count).SHEET_NO & "、商品コード：" & CheckData(Count).I_CODE & vbCr
            End If
        Next

        If Cancel_Message <> "" Then
            Cancel_Message &= "ピッキング済みデータに対する出荷指示キャンセルは" & vbCr & "予定数量とキャンセル数量が同一でなければなりません。" & vbCr & "出荷指示ファイルの取り込みを中止します。"
            MsgBox(Cancel_Message)
            Exit Sub
        End If

        If Out_Shiped_MessageList <> "" Then
            Out_Shiped_MessageList &= "は既に出荷済み、もしくは出荷指示のキャンセルが行えない為、出荷指示ファイルの取り込みを中止します。"
        End If
        If Out_NodataCancel_MessageList <> "" Then
            Out_NodataCancel_MessageList &= "はキャンセルする出荷指示データがない為、出荷指示ファイルの取り込みを中止します。"
        End If
        If Out_Shiped_MessageList <> "" Then
            MsgBox(Out_Shiped_MessageList)
            'ファイル選択ボタンをクリックした時にユニークキーを作成している為、
            '再度指定させる為にファイル名表示欄をクリアする。
            TextBox1.Text = ""
            Exit Sub
        End If

        If CheckBox1.Checked = "False" Then

            If Out_NodataCancel_MessageList <> "" Then
                MsgBox(Out_NodataCancel_MessageList)
                'ファイル選択ボタンをクリックした時にユニークキーを作成している為、
                '再度指定させる為にファイル名表示欄をクリアする。
                TextBox1.Text = ""
                Exit Sub
            End If
        End If

        '上記で登録したデータの一覧を表示する。
        For Count = 0 To CheckData.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '№
                .Cells(0).Value = Count + 1
                'ステータス
                '.Cells(1).Value = CheckData(Count).STATUS
                '数量がプラスで、OUTTBL.IDが0(データ無し）なら、通常出荷指示データ
                If CheckData(Count).OUTPRT_NUM > 0 And CheckData(Count).OUT_ID = 0 Then
                    .Cells(1).Value = "通常出荷指示データ"
                ElseIf CheckData(Count).OUTPRT_NUM > 0 And CheckData(Count).OUT_ID <> 0 Then
                    '数量がプラスで、OUTTBL.IDが0以外（同じ伝票番号がすでにDBに存在しているデータ）なら、
                    '重複登録の通常出荷指示データ（登録ボタンを押した時にアラート表示対象）
                    .Cells(1).Value = "重複通常出荷指示データ"
                    OrderNo_Duplication_Check = False
                ElseIf CheckData(Count).OUTPRT_NUM < 0 And CheckData(Count).STATUS = "出荷予定" Then
                    '数量がマイナスでステータスが出荷予定なら、出荷指示取り消し。
                    .Cells(1).Value = "出荷指示キャンセル"
                ElseIf CheckData(Count).OUTPRT_NUM < 0 And CheckData(Count).STATUS = "ピッキング済み" Then
                    '数量がマイナスでステータスがピッキング済み、既存データをピッキング戻しに設定。
                    .Cells(1).Value = "ピッキング戻し"
                End If

                '納品先コード
                .Cells(2).Value = CheckData(Count).C_CODE
                '納品先名
                .Cells(3).Value = CheckData(Count).C_NAME
                If CheckData(Count).C_ID = "0" Then
                    'マスタに該当する商品データがなければ、登録不可とし、背景色を変更。
                    Regist_Check_Flg = True
                    .Cells(3).Style.BackColor = Color.Salmon
                End If
                'オーダー番号
                .Cells(4).Value = CheckData(Count).ORDER_NO
                '伝票番号
                .Cells(5).Value = CheckData(Count).SHEET_NO
                '商品コード
                .Cells(6).Value = CheckData(Count).I_CODE
                '商品名
                .Cells(7).Value = CheckData(Count).I_NAME
                If CheckData(Count).I_ID = "0" Then
                    'マスタに該当する商品データがなければ、登録不可とし、背景色を変更。
                    Regist_Check_Flg = True
                    .Cells(7).Style.BackColor = Color.Salmon
                End If

                '価格
                .Cells(8).Value = CheckData(Count).UNIT_COST
                '数量
                .Cells(9).Value = CheckData(Count).OUTPRT_NUM
                '合計
                .Cells(10).Value = CheckData(Count).TOTAL_AMOUNT
                'コメント１（伝票用）
                .Cells(11).Value = CheckData(Count).COMMENT1
                'コメント２
                .Cells(12).Value = CheckData(Count).COMMENT2
                '出荷予定日
                .Cells(13).Value = CheckData(Count).O_DATE
                '売単価
                .Cells(14).Value = CheckData(Count).PRICE
                'OUT_PRT.ID
                .Cells(15).Value = CheckData(Count).ID
                'OUT_PRT.I_ID
                .Cells(16).Value = CheckData(Count).I_ID
                'OUT_PRT.C_ID
                .Cells(17).Value = CheckData(Count).C_ID
                'OUT.ID
                .Cells(18).Value = CheckData(Count).OUT_ID

                '出荷テーブル数量
                .Cells(19).Value = CheckData(Count).OUTTBL_NUM

            End With
            DataGridView1.Rows.Add(SR_List)
        Next

        If Regist_Check_Flg = True Then
            MsgBox("取り込んだデータの中にマスタ登録されていない納品先、商品があります。")
            Exit Sub
        End If

        'オブジェクトの破棄
        OBJ = Nothing

    End Sub

    Private Sub syukoshiji_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim PlaceData() As Place_List = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Disp_Title As String = "出荷指示ファイル取込"

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

        'DataGridViewを書き込み禁止にする。
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub syukoshiji_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

        'ファイル名用に日時を取得
        Dim dtNowFile As DateTime = DateTime.Now

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

        Dim FileName As String = Nothing
        FileName = Trim(TextBox1.Text)
        FileName = FileName.Replace(".xls", "")


        '発注日をそれぞれ年月日にわける。
        OrderYear = DateTimePicker1.Value.ToString("yy")
        OrderMonth = DateTimePicker1.Value.ToString("MM")
        OrderDay = DateTimePicker1.Value.ToString("dd")

        D_YMD = DateTimePicker1.Value.ToString("yyyy/MM/dd")
        '納品日は発注日＋１にしたものを格納
        D_YMD = D_YMD.AddDays(1)
        Y2YMD = D_YMD.ToString("yy/MM/dd")
        Delivery = CStr(Y2YMD).Split("/")

        'カンセキ伝票のファイル名
        Dim Sheet1_Name As String = "納品書" & dtNowFile.ToString("yyyyMMddHHmm") & "（株）カンセキ.csv"
        'カンセキ伝票のHeader設定
        Dim Sheet1_Header As String = "伝票番号,CUST_NAME,納品先名,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,製品コード,製品名,JAN,U,V,W,X,Y,出荷数,仕切価格,AB,小売価格,AD,AE,AF,AG,備考,伝票枚数ID,伝票総枚数,CUST_CD,コメント2"
        'T3伝票のファイル名
        Dim Sheet2_Name As String = "納品書" & dtNowFile.ToString("yyyyMMddHHmm") & "（T3）.csv"
        'T3伝票のHeader設定
        Dim Sheet2_Header As String = "伝票番号,CUST_NAME,納品先名,納品先ID,E,F,G,H,I,J,K,L,M,N,O,P,Q,製品コード,製品名,JAN,U,V,W,X,Y,出荷数,仕切価格,AB,小売価格,AD,AE,AF,AG,備考,伝票枚数ID,伝票総枚数,CUST_CD,コメント2"
        'タカミヤ伝票ファイル名 
        Dim Sheet3_Name As String = "納品書" & dtNowFile.ToString("yyyyMMddHHmm") & "タカミヤ.csv"
        'タカミヤ伝票のHeader設定
        Dim Sheet3_Header As String = "伝票番号,発注年,発注月,発注日,出荷年,出荷月,出荷日,B,C,D,納入先住所,納品先名,発注者,納品先ID,店舗ｺｰﾄﾞ,伝票区分,取引先ｺｰﾄﾞ,製品コード,製品名,JAN,出荷数,訂正後数量,仕切価格,納入金額,小売価格,備考,数量合計,E,F,納入金額合計,H,I,J,35,コメント2"


        Dim strEncoding As System.Text.Encoding
        Dim strStreamWriter As System.IO.StreamWriter
        strEncoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データが選択されていません。")
            Exit Sub
        End If

        'チェックされた商品があるか確認
        For Count = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Count).Cells(0).Value() = 1 Then
                Slip_Check_Flg = True
            End If
        Next

        If Slip_Check_Flg = False Then
            MsgBox("チェックされた商品がありません。")
            Exit Sub
        End If

        ''出荷指示ファイルで検索されていることが条件なので
        ''検索条件に出荷指示ファイル名が入っていて、DataGridViewに表示されている全てのデータが
        ''出荷指示ファイル名と同じであることをチェックする。
        'If Trim(TextBox3.Text) = "" Then
        '    MsgBox("出荷伝票CSVを出力するには出荷指示ファイルでの検索が必須となります。")
        '    Exit Sub
        'End If

        'For Count = 0 To DataGridView1.Rows.Count - 1
        '    If Trim(TextBox3.Text) <> DataGridView1.Rows(Count).Cells(11).Value() Then
        '        MsgBox("出荷伝票CSVを出力するには" & vbCr & "出荷指示ファイル単位での指定が必須です。")
        '        Exit Sub
        '    End If
        'Next

        DataCount = 0
        Slip_Data_List = Nothing
        Slip_Result = True
        Slip_ErrorMessage = Nothing
        'カンセキ伝票のデータを取得
        Slip_Result = GetSlipCancelData(FileName, 1, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)

        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(FilePath & Sheet1_Name, False, strEncoding)

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
                'H列に会社の社名
                LineData &= """" & Com_NAME & ""","
                'I列に会社のTEL番号
                LineData &= """" & Com_TEL & ""","
                'J列に会社のFAX番号
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


        'T3
        DataCount = 0
        Slip_Data_List = Nothing
        Slip_Result = True
        Slip_ErrorMessage = Nothing

        MaxPagecount = 1
        SamePageCount = 1
        PageCounter = 1

        'T3伝票のデータを取得
        Slip_Result = GetSlipCancelData(FileName, 2, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)
        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(FilePath & Sheet2_Name, False, strEncoding)

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
                'H列に会社の社名
                LineData &= """" & Com_NAME & ""","
                'I列に会社のTEL番号
                LineData &= """" & Com_TEL & ""","
                'J列に会社のFAX番号
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
            'T3伝票作成フラグ
            Pfj_Check = True
            'ファイルを閉じる
            strStreamWriter.Close()
        End If


        'タカミヤ伝票のデータを取得
        Slip_Result = GetSlipCancelData(FileName, 3, Slip_Data_List, DataCount, Slip_Result, Slip_ErrorMessage)
        If Slip_Result = False Then
            MsgBox(Slip_ErrorMessage)
            Exit Sub
        End If

        'データがなかったらファイルを作成しない
        If DataCount <> 0 Then

            'ﾌｧｲﾙ名の設定
            strStreamWriter = New System.IO.StreamWriter(FilePath & Sheet3_Name, False, strEncoding)

            'headerを設定
            strStreamWriter.WriteLine(Sheet3_Header)
            'データ件数分ループ
            For i = 0 To Slip_Data_List.Length - 1
                '取得するデータをLineataに格納する。
                'A列に伝票番号(下6桁を表示)
                LineData = """" & Slip_Data_List(i).SHEET_NO.Substring(4) & ""","

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
                'AE列に会社の社名
                LineData &= Com_NAME & ","
                'AF列に会社のTEL番号
                LineData &= Com_TEL & ","
                'AG列に会社のFAX番号
                LineData &= Com_FAX & ","
                'AH列はT固定（項目名：35）
                LineData &= "T,"
                'AI列は設定なし（項目名：コメント２）
                LineData &= ","

                '1行にしたデータを登録
                strStreamWriter.WriteLine(LineData)
            Next
            'takamiya伝票作成フラグ
            Takamiya_Check = True
            'ファイルを閉じる
            strStreamWriter.Close()
        End If

        Csv_Complete_Message = FilePath & "に" & vbCr & "出荷伝票CSVの作成が完了しました。"
        If Kanseki_Check = False Then
            Csv_Complete_Message &= vbCr & "データがない為、カンセキ伝票は作成されませんでした。"
        End If
        If Pfj_Check = False Then
            Csv_Complete_Message &= vbCr & "データがない為、T3伝票は作成されませんでした。"
        End If
        If Takamiya_Check = False Then
            Csv_Complete_Message &= vbCr & "データがない為、タカミヤ伝票は作成されませんでした。"
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