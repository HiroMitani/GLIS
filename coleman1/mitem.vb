Imports System.IO
Imports System
Imports System.Text

Public Class mitem

    '取り込んだエクセルのロケーションに不備がある場合
    '登録を不可能にするためのBoolean（True：登録不可、False:登録可能）
    Dim Location_Check_Flg As Boolean = False
    Dim Location_Check_ErrorMessage As String = Nothing

    '取り込んだエクセルのデータに不備がある場合
    '登録を不可能にするためのBoolean（True：登録不可、False:登録可能）
    Dim Regist_Check_Flg As Boolean = False

    Private Sub mitem_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub mitem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "商品マスタ登録"

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

        'DataGridViewを書き込み禁止にする。
        DataGridView1.ReadOnly = True

        Button1.Focus()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim ExcelErrorMessage As String = Nothing
        Dim ExcelResult As Boolean = True

        Dim GetItemErrorMessage As String = Nothing
        Dim GetItemResult As Boolean = True
        Dim GetItemFlg As Boolean = True

        Dim GetPLErrorMessage As String = Nothing
        Dim GetPLResult As Boolean = True

        Dim DataCount As Integer = 0

        Dim PLName As String = Nothing

        'Excelのデータを格納するための配列
        Dim ExcelData() As MIns_Item_List = Nothing

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

        'ExcelのデータをDataGridViewに取り込む。
        Do While Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 1).value) <> ""
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            DataCount = 0
            PLName = Nothing

            With SR_List
                '商品コードがすでに登録されているかチェック
                GetItemResult = GetItemDuplicationCheck(Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value), DataCount, GetItemResult, GetItemErrorMessage)
                If GetItemResult = False Then
                    MsgBox(GetItemErrorMessage)
                    Exit Sub
                End If

                '商品コード
                .Cells(0).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value)
                '商品名
                .Cells(1).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 2).value)
                'すでに同じ商品コードが登録されていたら、フラグをFalseにして、背景色をつける。
                If DataCount <> 0 Then
                    .Cells(1).Style.BackColor = Color.Salmon
                    Regist_Check_Flg = True
                    ExcelResult = False
                    ExcelErrorMessage = "商品コードはすでに登録されています。"
                End If

                'プロダクトラインコードがすでに登録されているかチェック
                GetPLResult = GETPLName(Integer.Parse(Trim(OBJ.Worksheets(SheetName).Cells(Count, 3).value)), PLName, GetPLResult, GetPLErrorMessage)
                If GetPLResult = False Then
                    MsgBox(GetPLErrorMessage)
                    Exit Sub
                End If

                'プロダクトラインコード
                .Cells(7).Value = Integer.Parse(Trim(OBJ.Worksheets(SheetName).Cells(Count, 3).value))
                If IsDBNull(PLName) Then
                    .Cells(7).Style.BackColor = Color.Salmon
                    Regist_Check_Flg = True
                Else
                    .Cells(8).Value = PLName
                End If

                'JANコード
                .Cells(2).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 4).value)
                '価格
                .Cells(3).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 5).value)
                '仕入金額
                .Cells(4).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 6).value)
                '免責金額
                .Cells(5).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 7).value)
                '修理金額
                .Cells(6).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value)

                'ロケーション
                If Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value) <> "" Then
                    .Cells(9).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value)
                Else
                    '未入力なら、背景色を変更する。
                    .Cells(9).Style.BackColor = Color.Salmon
                    Location_Check_Flg = True
                End If

                '2はセット商品、それ以外は通常商品とする。
                If Trim(Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 10).value)) = "2" Then
                    .Cells(10).Value = "セット商品"
                Else
                    .Cells(10).Value = "通常商品"
                End If

                Count += 1
            End With
            DataGridView1.Rows.Add(SR_List)

        Loop

        If ExcelResult = False Then
            MsgBox(ExcelErrorMessage)
            Exit Sub
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim MItem_Data() As MIns_Item_List = Nothing

        Dim GetItemErrorMessage As String = Nothing
        Dim GetItemResult As Boolean = True

        Dim DataCount As Integer = 0

        Dim ItemCheckFLG As Boolean = False

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("商品マスタを取り込んでから登録ボタンを押してください。")
            Exit Sub
        End If

        If Regist_Check_Flg = True Then
            MsgBox("取り込んだデータの中に登録済みのデータがあります。")
            Exit Sub
        End If

        If Location_Check_Flg = True Then

            'ダイアログ設定
            Dim locationcheck As DialogResult = MessageBox.Show("ロケーションが未設定ですが登録してもよろしいですか？", _
                                                         "確認", _
                                                         MessageBoxButtons.YesNo, _
                                                         MessageBoxIcon.Question)

            If locationcheck = DialogResult.No Then
                'ダイアログで「いいえ」が選択された時 
                'Application.Exit()
                'DataGridViewのクリア
                DataGridView1.Rows.Clear()
                'ファイル名のクリア
                TextBox1.Text = ""
                Exit Sub
            Else

            End If
        End If

        ReDim MItem_Data(DataGridView1.Rows.Count - 1)
        For Count = 0 To DataGridView1.Rows.Count - 1

            '商品コードがすでに登録されているかチェック
            GetItemResult = GetItemDuplicationCheck(DataGridView1(0, Count).Value, DataCount, GetItemResult, GetItemErrorMessage)
            If GetItemResult = False Then
                MsgBox(GetItemErrorMessage)
                Exit Sub
            End If
            If DataCount = 1 Then
                MsgBox("すでに登録済みのデータがあります。")
                DataGridView1.Item(1, Count).Style.BackColor = Color.Salmon
                ItemCheckFLG = True

            End If

            '商品コード
            MItem_Data(Count).I_CODE = DataGridView1(0, Count).Value
            '商品名
            MItem_Data(Count).I_NAME = DataGridView1(1, Count).Value
            'JANコード
            MItem_Data(Count).JAN = DataGridView1(2, Count).Value
            '価格
            MItem_Data(Count).PRICE = DataGridView1(3, Count).Value
            '仕入金額
            MItem_Data(Count).PURCHASE_PRICE = DataGridView1(4, Count).Value
            '免責金額
            MItem_Data(Count).IMMUNITY_PRICE = DataGridView1(5, Count).Value
            '修理金額
            MItem_Data(Count).REPAIR_PRICE = DataGridView1(6, Count).Value

            'プロダクトラインコード
            MItem_Data(Count).PL_CODE = DataGridView1(7, Count).Value
            'ロケーション
            MItem_Data(Count).LOCATION = DataGridView1(9, Count).Value
            'セット商品区分(通常商品なら0、セット商品なら1)
            If DataGridView1(10, Count).Value = "通常商品" Then
                MItem_Data(Count).SET_FLG = 0
            ElseIf DataGridView1(10, Count).Value = "セット商品" Then
                MItem_Data(Count).SET_FLG = 1
            End If
        Next

        If ItemCheckFLG = True Then
            MsgBox("すでに登録済みのデータがあります。")
        End If

        Result = MIns_Item(MItem_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("商品マスタの登録が完了しました。")

        '情報をクリア
        '企業マスタファイル名のクリア
        TextBox1.Text = ""
        'DataGridViewのクリア
        DataGridView1.Rows.Clear()

    End Sub
End Class