Imports System.IO
Imports System
Imports System.Text

Public Class mcustomer

    '伝票タイプの入力チェック
    '取り込んだエクセルのデータに不備がある場合
    '登録を不可能にするためのBoolean（True：登録不可、False:登録可能）
    Dim Sheet_Check_Flg As Boolean = False

    '取り込んだエクセルのデータに不備がある場合
    '登録を不可能にするためのBoolean（True：登録不可、False:登録可能）
    Dim Regist_Check_Flg As Boolean = False

    Private Sub mcustomer_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub mcustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Disp_Title As String = "企業マスタ登録"

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

        Button1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ExcelErrorMessage As String = Nothing
        Dim ExcelResult As Boolean = True

        Dim GetCustomerErrorMessage As String = Nothing
        Dim GetCustomerResult As Boolean = True
        Dim GetCustomerFlg As Boolean = True

        Dim GetPLErrorMessage As String = Nothing
        Dim GetPLResult As Boolean = True

        '商品情報取得用

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
        '納品先コードがある間はループ
        Do While Convert.ToString(OBJ.Worksheets(SheetName).Cells(Count, 1).value) <> ""
            Dim SR_List As New DataGridViewRow
            SR_List.CreateCells(DataGridView1)
            DataCount = 0
            PLName = Nothing

            With SR_List
                '企業コードがすでに登録されているかチェック
                GetCustomerResult = GetCustomerDuplicationCheck(Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value), DataCount, GetCustomerResult, GetCustomerErrorMessage)
                If GetCustomerResult = False Then
                    MsgBox(GetCustomerErrorMessage)
                    Exit Sub
                End If

                '企業コード
                .Cells(0).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 1).value)
                '企業名
                '2012/2/10 企業マスタのフォーマットが変更になったのに伴い、修正
                '.Cells(1).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 7).value)
                .Cells(2).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 3).value)

                'すでに同じ企業コードが登録されていたら、フラグをFalseにして、背景色をつける。
                If DataCount <> 0 Then
                    .Cells(0).Style.BackColor = Color.Salmon
                    Regist_Check_Flg = True
                Else
                    .Cells(0).Style.BackColor = Color.White
                End If

                '請求先コード   追加 2015.02.02 h.mitani
                .Cells(1).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 2).value)

                '伝票タイプ
                '2012/2/10 企業マスタのフォーマットが変更になったのに伴い、修正
                'If Trim(OBJ.Worksheets(SheetName).Cells(Count, 26).value) = "1" Then
                If Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value) = "1" Then
                    '1はカンセキ伝票
                    .Cells(3).Value = "カンセキ伝票"
                    '2012/2/10 企業マスタのフォーマットが変更になったのに伴い、修正
                    'ElseIf Trim(OBJ.Worksheets(SheetName).Cells(Count, 26).value) = "2" Then
                ElseIf Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value) = "2" Then
                    '2はPFJ伝票
                    .Cells(3).Value = "T3伝票"
                    '2012/2/10 企業マスタのフォーマットが変更になったのに伴い、修正
                    'ElseIf Trim(OBJ.Worksheets(SheetName).Cells(Count, 26).value) = "3" Then
                ElseIf Trim(OBJ.Worksheets(SheetName).Cells(Count, 9).value) = "3" Then
                    '3はタカミヤ伝票
                    .Cells(3).Value = "タカミヤ伝票"
                Else
                    .Cells(3).Style.BackColor = Color.Salmon
                    Sheet_Check_Flg = True
                End If

                '掛け率（空白だったら0を入れる）
                If Trim(OBJ.Worksheets(SheetName).Cells(Count, 10).value) = "" Then
                    .Cells(4).Value = 0
                Else
                    .Cells(4).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 10).value)
                End If

                '納品先名
                .Cells(5).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 4).value)
                '納品先郵便番号
                .Cells(6).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 5).value)
                '納品先住所
                .Cells(7).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 6).value)
                '納品先TEL
                .Cells(8).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 7).value)
                '納品先FAX
                .Cells(9).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 8).value)
                'カスタマー種別
                .Cells(10).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 12).value)
                '納品先別出荷リストフラグ
                .Cells(11).Value = Trim(OBJ.Worksheets(SheetName).Cells(Count, 11).value)
                Count += 1
            End With
            DataGridView1.Rows.Add(SR_List)

        Loop

        If Sheet_Check_Flg = True Then
            MsgBox("伝票の種類が指定されていません。")
            Exit Sub
        End If

        If Regist_Check_Flg = True Then
            MsgBox("取り込んだデータはすでに登録されています。")
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        topmenu.Show()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim MCustomer_Data() As MIns_Customer_List = Nothing

        Dim GetCustomerErrorMessage As String = Nothing
        Dim GetCustomerResult As Boolean = True

        Dim DataCount As Integer = 0

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("企業マスタファイルを取り込んでから登録ボタンを押してください。")
            Exit Sub
        End If

        If Sheet_Check_Flg = True Then
            MsgBox("伝票の種類が指定されていません。")
            Exit Sub
        End If

        If Regist_Check_Flg = True Then
            MsgBox("取り込んだデータの中に登録済みのデータがあります。")
            Exit Sub
        End If

        ReDim MCustomer_Data(DataGridView1.Rows.Count - 1)
        For Count = 0 To DataGridView1.Rows.Count - 1

            '企業コードがすでに登録されているかチェック
            GetCustomerResult = GetItemDuplicationCheck(DataGridView1(0, Count).Value, DataCount, GetCustomerResult, GetCustomerErrorMessage)
            If GetCustomerResult = False Then
                MsgBox(GetCustomerErrorMessage)
                Exit Sub
            End If
            'すでに同じ企業コードが登録されていたら、フラグをFalseにして、背景色をつける。
            If DataCount <> 0 Then
                DataGridView1.Item(2, Count).Style.BackColor = Color.Salmon
                Regist_Check_Flg = True
            End If


            DataGridView1.Item(2, Count).Style.BackColor = Color.White
            '企業コード
            MCustomer_Data(Count).C_CODE = DataGridView1(0, Count).Value
            '請求先コード
            MCustomer_Data(Count).CLAIM_CODE = DataGridView1(1, Count).Value

            '企業名
            MCustomer_Data(Count).C_NAME = DataGridView1(2, Count).Value
            '伝票タイプ
            If DataGridView1(3, Count).Value = "T3伝票" Then
                MCustomer_Data(Count).SHEET_TYPE = 2
            ElseIf DataGridView1(3, Count).Value = "カンセキ伝票" Then
                MCustomer_Data(Count).SHEET_TYPE = 1
            ElseIf DataGridView1(3, Count).Value = "タカミヤ伝票" Then
                MCustomer_Data(Count).SHEET_TYPE = 3
            End If
            '掛け率
            If DataGridView1(4, Count).Value = "" Then
                MCustomer_Data(Count).DISCOUNT_RATE = 0
            Else
                MCustomer_Data(Count).DISCOUNT_RATE = DataGridView1(4, Count).Value
            End If

            '納品先名
            MCustomer_Data(Count).D_NAME = DataGridView1(5, Count).Value
            '納品先郵便番号
            MCustomer_Data(Count).D_ZIP = DataGridView1(6, Count).Value
            '納品先住所
            MCustomer_Data(Count).D_ADDRESS = DataGridView1(7, Count).Value
            '納品先TEL
            MCustomer_Data(Count).D_TEL = DataGridView1(8, Count).Value
            '納品先FAX
            MCustomer_Data(Count).D_FAX = DataGridView1(9, Count).Value

            'カスタマー種別
            MCustomer_Data(Count).CUSTOMER_TYPE = DataGridView1(10, Count).Value
            '納品先別出荷リスト
            MCustomer_Data(Count).DELIVERY_OUTPUT_FLG = DataGridView1(11, Count).Value
        Next

        If Regist_Check_Flg = True Then
            MsgBox("すでに登録済みのデータがあります。")
            Exit Sub
        End If

        Result = MIns_Customer(MCustomer_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        MsgBox("企業マスタの登録が完了しました。")

        '情報をクリア
        '企業マスタファイル名のクリア
        TextBox1.Text = ""
        'DataGridViewのクリア
        DataGridView1.Rows.Clear()

    End Sub
End Class