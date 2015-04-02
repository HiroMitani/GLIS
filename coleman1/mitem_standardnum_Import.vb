Imports System.IO
Imports System
Imports System.Text

Public Class mitem_standardnum_Import

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim FileName As String = Nothing

        Dim ChkFileName As String = Nothing

        Dim DataErrorMessage As String = Nothing

        Dim NumChkErrorMessage As String = Nothing
        Dim NumChkResult As Boolean = True

        'ExcelファイルのデータをDBに登録できたかチェックする為の変数
        Dim ExcelErrorMessage As String = Nothing
        Dim ExcelResult As Boolean = True

        'Excelのデータを格納するための配列
        Dim ExcelData() As Standardnum_Import_List = Nothing
        'TMPテーブルからデータを取得する際に格納する配列
        Dim Search_List() As Standardnum_Import_List = Nothing

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

        'Excelファイルを読み込みDBに登録する為の変数設定
        '1行目は項目行の為、2に設定する。（1の場合、項目名も取得）
        Dim Count As Integer = 2
        'シート名取得
        Dim SheetName As String = OBJ.Worksheets(1).name

        'ExcelのデータをテンポラリーTableに取り込む。（データの不備等は登録後チェックを行う為、無条件で全件取り込む）
        Do While Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 1).value) <> ""
            '配列の再設定
            ReDim Preserve ExcelData(0 To Count - 2)

            '商品コード
            ExcelData(Count - 2).I_CODE = Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value)
            '商品コードの長さをチェック。
            If ExcelData(Count - 2).I_CODE.Length > 24 Then
                DataErrorMessage &= Count - 1 & "行目の商品コードが長すぎます。" & vbCr
            End If

            '基準値
            'NULL=NG、マイナス値=NG,0の値=OK
            If NumChkVal(Trim(OBJ.Worksheets(SheetName).Cells(Count, 2).value), "INTEGER", False, False, True, NumChkResult, NumChkErrorMessage) = False Then

                DataErrorMessage &= Count - 1 & "行目の基準値が正しくありません。" & vbCr
            Else
                ExcelData(Count - 2).STANDARD_NUM = OBJ.Worksheets(SheetName).Cells(Count, 2).value
            End If

            Count += 1
        Loop

        'DataErrorMessageに値が入っていたら、エラー内容を表示
        If DataErrorMessage <> "" Then
            MsgBox(DataErrorMessage)
            Exit Sub
        End If

        'TMPテーブル作成
        ExcelResult = CREATE_TMP_M_ITEM(Result, ErrorMessage)

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If

        'TMP_M_ITEMへ登録
        ExcelResult = Ins_TMP_M_ITEM(ExcelData, ExcelResult, ExcelErrorMessage)

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If

        '登録したTMP_M_ITEMを読み込む。
        Result = GET_TMP_M_ITEM(Search_List, Result, ErrorMessage)

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If

        'DataGridViewに表示
        For Count = 0 To Search_List.Length - 1
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            With SR_List
                '取込状況
                If Search_List(Count).I_CODE = "" Then
                    .Cells(1).Value = "データ不備"
                Else
                    .Cells(1).Value = "○"
                End If

                '№
                .Cells(0).Value = Count + 1

                '商品コード
                .Cells(2).Value = Search_List(Count).I_CODE
                '商品名
                .Cells(3).Value = Search_List(Count).I_NAME

                'JANコード
                .Cells(4).Value = Search_List(Count).JAN
                '基準値
                .Cells(5).Value = Search_List(Count).STANDARD_NUM
                'プロダクトライン名
                .Cells(6).Value = Search_List(Count).PL_NAME
                'I_ID
                .Cells(7).Value = Search_List(Count).I_ID

            End With
            DataGridView1.Rows.Add(SR_List)
        Next

        'オブジェクトの破棄
        OBJ = Nothing

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub mitem_standardnum_Import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim Mitem_Data() As Standardnum_Import_List = Nothing

        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("基準値変更ファイルの取込みを行ってから、登録ボタンを押してください。")
            Exit Sub
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("基準値変更ファイルを取り込んでください。")
            Exit Sub
        End If

        'DataGridViewの情報を配列に入れる

        ReDim Mitem_Data(DataGridView1.Rows.Count - 1)
        For Count = 0 To DataGridView1.Rows.Count - 1

            'I_ID
            Mitem_Data(Count).I_ID = DataGridView1(7, Count).Value
            '基準値
            Mitem_Data(Count).STANDARD_NUM = DataGridView1(5, Count).Value

        Next

        Result = Upd_StandardNum(Mitem_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("基準値を更新が完了しました。")

        'TMP_M_ITEMをDROPする。
        Result = DROP_TMP_M_ITEM(Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()
        '発注ファイル名をクリアする。
        TextBox1.Text = ""



    End Sub
End Class