'チャート関係の名前空間です。
Imports System.Windows.Forms.DataVisualization.Charting

Public Class IO_Graph


    Private Sub IO_Graph_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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


    Private Sub IO_Graph_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle



        Dim Item_Data() As Item_List = Nothing
        Dim I_Code As String = Nothing
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        Dim PO_Data() As GraphData = Nothing
        Dim IN_Data() As GraphData = Nothing
        Dim OUT_Data() As GraphData = Nothing
        Dim Prediction_Data() As GraphData = Nothing

        Dim DISMAN_Data() As GraphData = Nothing



        Dim STOCK_Data() As GraphData = Nothing


        '商品IDを元に商品名を取得。
        'データ取得Function
        Result = GetItemInfo(Label1.Text, Item_Data, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        '当月を取得
        Dim NowYM As String

        ' 必要な変数を宣言する
        Dim dtNow As DateTime = DateTime.Now

        '     MsgBox(dtNow.ToString("yyyy/MM"))

        ' 年 (Year) を取得する
        Dim Year As Integer = dtNow.Year
        Dim Month As Integer = dtNow.Month
        NowYM = Year & "/" & Month

        Dim DateFrom As String = dtNow.AddMonths(-2).ToString("yyyy/MM")
        Dim DateTo As String = dtNow.AddMonths(2).ToString("yyyy/MM")
        ReDim Preserve PO_Data(0 To 4)
        'PO_Data(0).YM = dtNow.AddMonths(-2).ToString("yyyy/MM")
 

        'データを取得（前後二ヶ月含め、計５か月分取得）
        ' Result = GetGraphData(Label1.Text, DateFrom, DateTo, PO_Data, IN_Data, OUT_Data, PO_DetailData, IN_DetailData, OUT_DetailData, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Button1.Visible = False

        'PrintDocumentオブジェクトの作成
        Dim pd As New System.Drawing.Printing.PrintDocument
        ''PrintPageイベントハンドラの追加
        AddHandler pd.PrintPage, AddressOf pd_PrintPage
        ''印刷を開始する
        'pd.Print()



        pd.DefaultPageSettings.Landscape = Not _
                                pd.DefaultPageSettings.Landscape

        Dim ppd As New PrintPreviewDialog

        ppd.Document = pd
        ppd.PrintPreviewControl.Zoom = 1
        ppd.Size = New Size(700, 800)


        ppd.ShowDialog()

        Button1.Visible = True

    End Sub

    Private Sub pd_PrintPage(ByVal sender As Object, _
        ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        '画像を読み込む
        Dim img As Image = Image.FromFile("C:\Documents and Settings\Administrator\My Documents\test.png")
        '画像を描画する
        ' e.Graphics.DrawImage(img, e.MarginBounds)


        e.Graphics.DrawImage(img, 20, 20, 1.1F * img.Width, 1.1F * img.Height)
        '次のページがないことを通知する
        e.HasMorePages = False
        '後始末をする
        img.Dispose()
    End Sub

End Class