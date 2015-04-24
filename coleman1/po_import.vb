Imports System.IO
Imports System
Imports System.Text

Public Class po_import

    Private Sub po_import_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        'Excelのデータを格納するための配列
        Dim ExcelData() As PO_List = Nothing

        
        Dim FNameErrorMessage As String = Nothing
        Dim FNameResult As Boolean = True

        'ExcelファイルのデータをDBに登録できたかチェックする為の変数
        Dim ExcelErrorMessage As String = Nothing
        Dim ExcelResult As Boolean = True

        Dim ChkFileName As String = Nothing

        Dim DataErrorMessage As String = Nothing

        Dim Tmp_PO_DATE As Date = Nothing

        Dim ChkDateAfter As String = Nothing
        Dim DateChkErrorMessage As String = Nothing
        Dim DateChkResult As Boolean = True

        Dim ChkComment1 As String = Nothing
        Dim Tmp_Comment1 As String = Nothing

        Dim Tmp_Remarks As String = Nothing

        'TMP_P_ORDER
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True
        'TMP_P_ORDERからデータを取得し格納するための配列
        Dim Search_List() As PO_List = Nothing


        Dim In_Data_Count As Integer = 0
        Dim Dup_Count As Integer = 0
        Dim Defect_Count As Integer = 0
        Dim Registed_Count As Integer = 0

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

        'Excelファイルを読み込みDBに登録する為の変数設定
        '1行目は項目行の為、2に設定する。（1の場合、項目名も取得）
        Dim Count As Integer = 2
        'シート名取得
        Dim SheetName As String = OBJ.Worksheets(1).name

        'ExcelのデータをテンポラリーTableに取り込む。（データの不備等は登録後チェックを行う為、無条件で全件取り込む）
        Do While Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 1).value) <> ""
            '配列の再設定
            ReDim Preserve ExcelData(0 To Count - 2)

            '発注先コード
            ExcelData(Count - 2).PO_CODE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 2).value)
            If ExcelData(Count - 2).PO_CODE.Length > 20 Then
                DataErrorMessage &= Count - 1 & "行目の発注先が長すぎます。" & vbCr
            End If

            '発注日
            If DateChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 5).value), False, ChkDateAfter, DateChkResult, DateChkErrorMessage) = False Then
                DataErrorMessage &= Count + 1 & "行目の発注日が正しくありません。" & vbCr
            Else
                '発注日
                ExcelData(Count - 2).ORDER_DATE = ChkDateAfter
            End If

            '発注№
            ExcelData(Count - 2).PO_NO = Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value)
            '発注№の長さをチェック。
            If ExcelData(Count - 2).PO_NO.Length > 20 Then
                DataErrorMessage &= Count - 1 & "行目の発注№が長すぎます。" & vbCr
            End If

            '商品コード
            ExcelData(Count - 2).I_CODE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 7).value)
            If ExcelData(Count - 2).I_CODE.Length > 24 Then
                DataErrorMessage &= Count - 1 & "行目の商品コードが長すぎます。" & vbCr
            End If


            '発注数
            '半角数値のみかチェック
            If System.Text.RegularExpressions.Regex.IsMatch(Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value), "^[0-9]+$", _
                System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                ExcelData(Count - 2).PO_NUM = Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value)
            Else
                DataErrorMessage &= Count - 1 & "行目の発注数が数値ではありません。" & vbCr
            End If

            '希望納期
            '日付の妥当性チェック
            If IsDate(Trim(OBJ.Worksheets(SheetName).Cells(Count, 6).value)) Then
                ExcelData(Count - 2).PO_DATE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 6).value)
            Else
                DataErrorMessage &= Count - 1 & "行目の希望納期が正しくありません。" & vbCr
            End If

            '備考
            'ExcelData(Count - 2).REMARKS = Trim(OBJ.Worksheets(SheetName).Cells(Count, 30).value)

            '　アポストロフィ対応とバックスラッシュ対応
            'Tmp_Remarks = Trim(OBJ.Worksheets(SheetName).Cells(Count, 30).value)
            'Tmp_Remarks = Tmp_Remarks.Replace("'", "''")
            'Tmp_Remarks = Tmp_Remarks.Replace("\", "\\")
            'ExcelData(Count - 2).REMARKS = Tmp_Remarks
            'If ExcelData(Count - 2).REMARKS.Length > 500 Then
            '    DataErrorMessage &= Count - 1 & "行目の備考が長すぎます。" & vbCr
            'Else
            '    'シングルクォート対応

            'End If

            Count += 1
        Loop

        'DataErrorMessageに値が入っていたら、エラー内容を表示
        If DataErrorMessage <> "" Then
            MsgBox(DataErrorMessage)
            Exit Sub
        End If

        'TMPテーブル作成
        ExcelResult = CREATE_TMP_P_ORDER(Result, ErrorMessage)

        If ExcelResult = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If


        'TMP_P_ORDERへ登録
        ExcelResult = Ins_TMP_P_ORDER(ExcelData, ExcelResult, ExcelErrorMessage)

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If


        '登録したTMP_P_ORDERを読み込む。
        Result = GET_TMP_P_ORDER(Search_List, Result, ErrorMessage)

        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'DataGridViewに表示
        For Count = 0 To Search_List.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '取込状況
                If Search_List(Count).R_CHECK = 0 Then
                    If Search_List(Count).DUPLICATE_CHECK = 1 Then
                        If Search_List(Count).I_ID = 0 Then
                            .Cells(0).Value = "データ不備"

                            Defect_Count += 1
                        Else
                            .Cells(0).Value = "○"
                            In_Data_Count += 1
                        End If
                    Else
                        .Cells(0).Value = "重複データ有"
                        .Cells(0).Style.BackColor = Color.Yellow
                        Dup_Count += 1
                    End If
                Else
                    .Cells(0).Value = "取込済み"
                    .Cells(0).Style.BackColor = Color.Gray
                    Registed_Count += 1
                End If

                '№
                .Cells(1).Value = Count + 1

                '発注先コード
                .Cells(2).Value = Search_List(Count).PO_CODE
                '発注先名
                .Cells(3).Value = Search_List(Count).PO_NAME

                '発注日
                .Cells(4).Value = Search_List(Count).ORDER_DATE
                '発注No
                .Cells(5).Value = Search_List(Count).PO_NO
                '商品コード
                .Cells(6).Value = Search_List(Count).I_CODE
                If Search_List(Count).I_ID = 0 Then
                    .Cells(6).Style.BackColor = Color.Salmon

                End If
                '商品名
                .Cells(7).Value = Search_List(Count).I_NAME
                '発注数
                .Cells(8).Value = Search_List(Count).PO_NUM
                '希望納期
                .Cells(9).Value = Search_List(Count).PO_DATE
                '備考
                .Cells(10).Value = Search_List(Count).REMARKS
                'I_ID
                .Cells(11).Value = Search_List(Count).I_ID
                'C_ID
                .Cells(12).Value = Search_List(Count).C_ID
                'PO_M_ID
                .Cells(13).Value = Search_List(Count).PO_M_ID
            End With
            DataGridView1.Rows.Add(SR_List)
        Next

        'オブジェクトの破棄
        OBJ = Nothing

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim Order_Data() As PO_List = Nothing
        Dim Count As Integer = 0
        Dim Regist_Data_Count As Integer = 0

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim Tmp_Remarks As String = Nothing

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("発注ファイルの取込みを行ってから、登録ボタンを押してください。")
            Exit Sub
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("発注ファイルを取り込んでください。")
            Exit Sub
        End If

        For Count = 0 To DataGridView1.Rows.Count - 1
            '取込状況
            'データが○のものだけ登録する。
            If DataGridView1(0, Count).Value = "○" Then
                ReDim Preserve Order_Data(0 To Regist_Data_Count)
                '商品ID
                Order_Data(Regist_Data_Count).I_ID = DataGridView1(11, Count).Value
                '発注No
                Order_Data(Regist_Data_Count).PO_NO = DataGridView1(5, Count).Value
                '発注数
                Order_Data(Regist_Data_Count).PO_NUM = DataGridView1(8, Count).Value
                '希望納期
                Order_Data(Regist_Data_Count).PO_DATE = DataGridView1(9, Count).Value
                '備考
                '　アポストロフィ対応とバックスラッシュ対応
                Tmp_Remarks = DataGridView1(10, Count).Value
                Tmp_Remarks = Tmp_Remarks.Replace("'", "''")
                Tmp_Remarks = Tmp_Remarks.Replace("\", "\\")
                Order_Data(Regist_Data_Count).REMARKS = Tmp_Remarks

                '発注日
                Order_Data(Regist_Data_Count).ORDER_DATE = DataGridView1(4, Count).Value

                'PO_M_ID
                Order_Data(Regist_Data_Count).PO_M_ID = DataGridView1(13, Count).Value

                Regist_Data_Count += 1

            End If
        Next
        'データが0件なら処理終了
        If Regist_Data_Count = 0 Then
            MsgBox("取込み可能なデータがありません。")
            Exit Sub
        End If

        Result = Ins_P_ORDER(Order_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox(Regist_Data_Count & "件の発注データの取り込みが完了しました。")

        'TMP_P_ORDERをDROPする。
        Result = DROP_TMP_P_ORDER(Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        '発注ファイル名をクリアする。
        TextBox1.Text = ""

    End Sub

    Private Sub po_import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "発注（PO）ファイル取込"

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