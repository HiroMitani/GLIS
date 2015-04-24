Imports MySql.Data.MySqlClient

Module DBMOB

    'システム名
    Public SYSTEM_NAME As String = "GLIS"

    '画面の表示位置設定
    Public DisplayX As Integer = System.Configuration.ConfigurationManager.AppSettings("DisplayX")
    Public DisplayY As Integer = System.Configuration.ConfigurationManager.AppSettings("DisplayY")
    '消費税額設定
    Public Tax As String = System.Configuration.ConfigurationManager.AppSettings("Tax")

    '作成するグラフの拡張子
    Public ImageType As String = ".png"

    'DB接続情報（開発用）
    Public Constring As String = System.Configuration.ConfigurationManager.AppSettings("DB")

    '帳票印字用企業情報
    Public Com_NAME As String = System.Configuration.ConfigurationManager.AppSettings("Com_NAME")
    Public Com_POST As String = System.Configuration.ConfigurationManager.AppSettings("Com_POST")
    Public Com_ADDRESS As String = System.Configuration.ConfigurationManager.AppSettings("Com_ADDRESS")
    Public Com_TEL As String = System.Configuration.ConfigurationManager.AppSettings("Com_TEL")
    Public Com_FAX As String = System.Configuration.ConfigurationManager.AppSettings("Com_FAX")
    Public Com_BANKNAME As String = System.Configuration.ConfigurationManager.AppSettings("Com_BANKNAME")
    Public Com_ACCOUNTINFO As String = System.Configuration.ConfigurationManager.AppSettings("Com_ACCOUNTINFO")

    'メモ欄情報
    Public Memo1 As String = System.Configuration.ConfigurationManager.AppSettings("Memo1")
    Public Memo2 As String = System.Configuration.ConfigurationManager.AppSettings("Memo2")

    'ログインしたユーザーのユーザーID、権限を格納
    Public R_User As Integer
    Public Authority As String
    'メニュー表示情報格納
    Public DISPLIST() As DispID_List

    Public Structure DispID_List
        Dim ID As Integer
        Dim FORM_X As Integer
        Dim FORM_Y As Integer
        Dim BUTTON_NO As Integer
    End Structure

    Public Structure Place_List
        Dim ID As Integer
        Dim NAME As String
    End Structure

    Public Structure C_List
        Dim ID As String
        Dim NAME As String
        Dim DISCOUNT_RATE As String
    End Structure

    Public Structure PL_List
        Dim ID As Integer
        Dim NAME As String
        Dim SHEET_TYPE As String
    End Structure

    Public Structure Item_List
        Dim ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim PL_CODE As String
        Dim LOCATION As String
        Dim PRICE As Integer
        Dim PL_NAME As String
        Dim PURCHASE_PRICE As Integer
        Dim IMMUNITY_PRICE As Integer
        Dim REPAIR_PRICE As Integer
        Dim DISCOUNT_RATE As Decimal

    End Structure

    Public Structure Ins_Item_List
        Dim ITEMCODE As String
        Dim NAME As String
        Dim NUM As Integer
        Dim ID As Integer
    End Structure

    Public Structure Search_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim DOC_NO As String
        Dim I_NAME As String
        Dim I_CODE As String
        Dim I_ID As Integer
        Dim JAN_CODE As String
        Dim NUM As Integer
        Dim N_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim STATUS As String
        Dim CATEGORY As String
        Dim DEFECT_TYPE As String
        Dim LOCATION As String
        Dim C_ID As Integer
        Dim C_CODE As String
        Dim C_NAME As String
        Dim REMARKS As String
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure ItemID_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
    End Structure

    Public Structure Output_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim SHEET_TYPE As String
    End Structure

    Public Structure In_Check_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim SHEET_TYPE As String
    End Structure

    Public Structure InDefinition_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_CODE As String
        Dim I_STATUS As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim LOCATION As String
        Dim I_ID As Integer
        Dim STOCK_COMMENT As String
        Dim IN_COMMENT As String
        Dim PLACE As Integer
    End Structure

    Public Structure ExcelData_List
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim FILE_NAME As String
        Dim I_CODE As String
        Dim C_CODE As String
        Dim NUM As Integer
        Dim UNIT_COST As Integer
        Dim TOTAL_AMOUNT As Integer
        Dim O_DATE As String
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim KEY As String
    End Structure

    Public Structure Upd_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim NUM As Integer
        Dim N_DATE As Date
        Dim CATEGORY As String
        Dim DEFECT_TYPE As String
    End Structure

    Public Structure Out_Upd_List
        Dim ID As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
    End Structure

    Public Structure OutDefinition_List
        Dim ID As Integer
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim I_ID As Integer
        Dim STOCK_NUM As Integer
        Dim I_STATUS As String
        Dim PLACE As Integer
    End Structure

    Public Structure Out_Search_List
        '出庫関連検索用。ロット番号CSV出力にも使用するため、Ｎｏ、ロット番号、保証書番号項目を追加している
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_ID As Integer
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim C_CODE As String
        Dim C_NAME As String
        Dim NUM As Integer
        Dim O_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim FILE_NAME As String
        Dim STATUS As String
        Dim CATEGORY As String
        Dim DEFECT_TYPE As String
        Dim COST As String
        Dim PRICE As String
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim REMARKS As String
        Dim LOCATION As String
        Dim PRT_DATE As String
        Dim STOCK_NUM As Integer
        Dim PLACE As String
        Dim P_ID As Integer
        Dim D_ZIP As String
        Dim D_ADDRESS As String
        Dim D_TEL As String
        Dim INQUIRY_NO As String
        Dim NO As Integer
        Dim LOT_NUMBER As String
        Dim WARRANTY_CARD_NUMBER As String
    End Structure

    Public Structure Out_Regist_List
        Dim ID As Integer
        Dim DETAIL_ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim UNIT_COST As Integer
        Dim NUM As Integer
        Dim TOTAL_AMOUNT As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim O_DATE As String
        Dim PRICE As Integer
        Dim OUT_ID As Integer
        Dim OUTDETAIL_ID As Integer
        Dim STATUS As String
        Dim I_STATUS As String
    End Structure

    Public Structure SheetNo_List
        Dim SHEET_NO As String
        Dim RESULT As Boolean
    End Structure

    Public Structure Duplication_Num_List
        Dim SHEET_NO As String
        Dim NUM As Integer
        Dim RESULT As Integer
    End Structure

    Public Structure Picking_List
        Dim ID As String
        Dim NUM As Integer
        Dim I_ID As Integer
        Dim I_STATUS As String
    End Structure

    Public Structure Picking_Prt_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim PL_ID As Integer
        Dim FILE_NAME As String
        Dim PL_NAME As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim NUM As Integer
        Dim STOCK_NUM As Integer
        Dim I_STATUS As String
        Dim PLACE_ID As Integer
        Dim C_CODE As String
        Dim LOCATION As String
    End Structure

    Public Structure Check_Rsult_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim C_CODE As String
        Dim C_NAME As String
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim OUTTBL_NUM As Integer
        Dim TOTAL_AMOUNT As Integer
        Dim OUTPRT_NUM As Integer
        Dim UNIT_COST As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim O_DATE As String
        Dim PRICE As Integer
        Dim STATUS As String
        Dim OUT_ID As Integer
        Dim I_STATUS As String
    End Structure

    Public Structure Stock_Search_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim LOCATION As String
        Dim I_STATUS As String
        Dim PL_NAME As String
        Dim PACKAGE_FLG As Boolean
        Dim P_ID As Integer
        Dim PLACE As String
        Dim SHIPPING_NUM As Integer
        Dim OUT_NUM As Integer
    End Structure

    Public Structure Stock_Tanaoroshi_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim LOCATION As String
        Dim I_STATUS As String
        Dim PL_NAME As String
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure Tanaoroshi_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim NUM As Integer
        Dim NEW_NUM As Integer
        Dim I_STATUS As String
        Dim REMARKS As String
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure Set_Item_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim BEFORE_NUM As Integer
        Dim SET_NUM As Integer
        Dim STOCK_NUM As Integer
        Dim I_STATUS As Integer
        Dim PLACE As Integer
    End Structure

    Public Structure IStatus_Change_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim NUM As Integer
        Dim CHANGE_NUM As Integer
        Dim I_STATUS As String
        Dim CHANGE_I_STATUS As String
        Dim NEW_LOCATION As String
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure Location_Change_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim CHANGE_NUM As Integer
        Dim I_STATUS As String
        Dim NEW_LOCATION As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure SlipID_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
    End Structure

    Public Structure Slip_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim SHEET_NO As String
        Dim ORDER_NO As String
        Dim O_DATE As String
        Dim D_ADDRESS As String
        Dim D_NAME As String
        Dim C_NAME As String
        Dim C_CODE As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As String
        Dim UNIT_COST As Integer
        Dim PRICE As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim FILE_NAME As String
        Dim PAGE As Integer
        Dim DATA_NUM As Integer
        Dim MAXPAGE As Integer
        Dim PRT_STATUS As String
    End Structure

    Public Structure DeliveryID_List
        Dim ID As Integer
    End Structure

    Public Structure DeliveryGroupID_List
        Dim ID As Integer
        Dim C_ID As Integer
    End Structure

    Public Structure Delivery_List
        Dim ID As Integer
        Dim C_ID As Integer
        Dim FILE_NAME As String
        Dim D_NAME As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim NUM As Integer
        Dim NUMTOTAL As Integer
        Dim C_NAME As String
    End Structure

    Public Structure CheckID_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
    End Structure

    Public Structure Check_List
        Dim ID As Integer
        Dim C_ID As Integer
        Dim C_CODE As String
        Dim D_NAME As String
        Dim SHEET_NO As String
        Dim DATA_NUM As Integer
        Dim FILE_NAME As String
    End Structure

    Public Structure CommentID_List
        Dim ID As Integer
    End Structure

    Public Structure StockID_List
        Dim STOCK_ID As Integer
    End Structure

    Public Structure Comment_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim FILE_NAME As String
        Dim C_CODE As String
        Dim SHEET_NO As String
        Dim I_CODE As String
        Dim NUM As Integer
        Dim COMMENT1 As String
        Dim COMMENT2 As String
    End Structure

    Public Structure StockLog_List
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim I_STATUS As String
        Dim NUM As String
        Dim I_FLG As String
        Dim PACKAGE_NO As String
        Dim U_DATE As String
        Dim COMMENT As String
        Dim PLACE As String
    End Structure

    Public Structure MIns_Item_List
        Dim I_CODE As String
        Dim I_NAME As String
        Dim PRICE As Integer
        Dim PL_CODE As String
        Dim JAN As String
        Dim PACKAGE_FLG As Boolean
        Dim LOCATION As String
        Dim SET_FLG As Integer
        Dim PURCHASE_PRICE As Integer
        Dim IMMUNITY_PRICE As Integer
        Dim REPAIR_PRICE As Integer
        Dim DISCOUNT_RATE As Decimal
    End Structure

    Public Structure MIns_Customer_List
        Dim C_CODE As String
        Dim C_NAME As String
        Dim CLAIM_CODE As String
        Dim SHEET_TYPE As Integer
        Dim D_NAME As String
        Dim D_ZIP As String
        Dim D_ADDRESS As String
        Dim D_TEL As String
        Dim D_FAX As String
        Dim DISCOUNT_RATE As Decimal
        Dim CUSTOMER_TYPE As Integer
        Dim DELIVERY_OUTPUT_FLG As Integer
    End Structure

    Public Structure Dismantling_List
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim NUM As Integer
        Dim I_STATUS As String
        Dim LOCATION As String
        Dim I_NAME As String
        Dim PLACE As String
        Dim PLACE_ID As Integer
    End Structure

    Public Structure RackCard_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim STOCK_NUM As Integer
    End Structure

    Public Structure Tanaoroshi_PrtList
        Dim LOCATION As String
        Dim PL_NAME As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim NUM As Integer
    End Structure

    Public Structure Stock_Item_List
        Dim STOCK_ID As Integer
        Dim I_ID As Integer
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim PL_NAME As String
        Dim I_STATUS As String
        Dim LOCATION As String
    End Structure

    Public Structure Summary_List
        Dim WorkDate As String
        Dim Line As Integer
        Dim Bait As Integer
        Dim Rods As Integer
        Dim B_Reels As Integer
        Dim S_Reels As Integer
        Dim Combos As Integer
        Dim Accessory As Integer
        Dim Accessory2 As Integer
        Dim Hard_Lure As Integer
        Dim Bag As Integer
        Dim Rod_OEM As Integer
        Dim Rod_Parts As Integer
        Dim ReelParts As Integer
    End Structure

    Public Structure PO_List
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim PO_NO As String
        Dim PO_NAME As String
        Dim PO_CODE As String
        Dim PO_DATE As String
        Dim PO_NUM As Integer
        Dim ORDER_DATE As String
        Dim REMARKS As String
        Dim R_CHECK As Integer
        Dim C_ID As Integer
        Dim C_CODE As String
        Dim DUPLICATE_CHECK As Integer
        Dim PO_M_ID As Integer
    End Structure

    Public Structure PO_Search_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim C_CODE As String
        Dim PO_NO As String
        Dim PO_DATE As String
        Dim PO_NUM As Integer
        Dim ORDER_DATE As String
        Dim REMARKS As String
        Dim STATUS As String
        Dim IN_NUM As Integer
        Dim IN_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim CANCEL_NUM As Integer
        Dim CANCEL_DATE As String
        Dim I_NAME As String
        Dim PL_NAME As String
        Dim STOCK_NUM As Integer
        Dim C_NAME As String
        Dim PO_M_NAME As String
        Dim PO_M_CODE As String
    End Structure

    Public Structure PO_In_List
        Dim VENDER_CODE As String
        Dim VENDER As String
        Dim PL_NAME As String
        Dim PO_NO As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim STATUS As String
        Dim PO_NUM As Integer
        Dim PO_DATE As String
        Dim REMARKS As String
        Dim IN_NUM As Integer
        Dim IN_DATE As String
        Dim FIX_NUM As Integer
        Dim FIX_DATE As String
        Dim CANCEL_NUM As Integer
        Dim CANCELED_NUM As Integer
        Dim CANCELED_DATE As String
        Dim ID As Integer
        Dim I_ID As Integer
        Dim INVOICE_NO As String
        Dim STOCK_NUM As Integer
        Dim C_NAME As String
        Dim PLACE As Integer
    End Structure

    Public Structure Monthly_Report_Search_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim PL_NAME As String
        Dim JAN As String
        Dim STOCK_NUM As Integer
        Dim ORDER_NUM As Integer
        Dim IN_NUM As Integer
        Dim Month1 As Integer
        Dim Month2 As Integer
        Dim Month3 As Integer
        Dim Month4 As Integer
        Dim Month5 As Integer
        Dim Month6 As Integer
        Dim Month7 As Integer
        Dim Month8 As Integer
        Dim Month9 As Integer
        Dim Month10 As Integer
        Dim Month11 As Integer
        Dim Month12 As Integer
    End Structure

    Public Structure M_Item_List
        Dim ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim PL_ID As Integer
        Dim PL_NAME As String
        Dim PRICE As Integer
        Dim PACKAGE_FLG As Integer
        Dim IN_BOX_NUM As Integer
        Dim MASTER_CARTON_SIZE As String
        Dim LOCATION As String
        Dim C_ID As Integer
        Dim C_CODE As String
        Dim C_NAME As String
        Dim PURCHASE_PRICE As Integer
        Dim IMMUNITY_PRICE As Integer
        Dim REPAIR_PRICE As Integer
    End Structure

    Public Structure M_Customer_List
        Dim ID As Integer
        Dim C_CODE As String
        Dim C_NAME As String
        Dim SHEET_TYPE As String
        Dim D_NAME As String
        Dim D_ZIP As String
        Dim D_ADDRESS As String
        Dim D_TEL As String
        Dim D_FAX As String
        Dim CUSTOMER_TYPE As String
        Dim DELIVERY_PRT_FLG As Integer
        Dim CLAIM_CODE As String
        Dim DISCOUNT_RATE As Decimal
    End Structure

    Public Structure IO_Balance_List
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim PL_NAME As String
        Dim PO_NUM As Integer
        Dim IN_NUM As Integer
        Dim IN_FIX_NUM As Integer
        Dim OUT_NUM As Integer
        Dim STOCK_NUM As Integer
    End Structure

    Public Structure IO_Graph_List
        Dim ID As Integer
        Dim PO_NUM As String
        Dim PO_DATE As String
    End Structure

    Public Structure Label_Prt_List
        Dim SHEET_NO As String
        Dim I_CODE As String
        Dim NUM As Integer
    End Structure

    Public Structure GraphData
        Dim DATE_DATA As String
        Dim ORDER_DATE As String
        Dim NUM As Integer
        Dim TYPE As String
        Dim STOCK_NUM As Integer
        Dim STOCK_ORDER_NUM As Integer
    End Structure

    Public Structure GraphSummaryData
        Dim DATE_DATA As String
        Dim TYPE As String
        Dim NUM As Integer
        Dim STOCK_NUM As Integer
    End Structure

    Public Structure PO_Prediction_List
        Dim ID As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim PL_NAME As String
        Dim STOCK_NUM As Integer
        Dim IN_NUM As Integer
        Dim OUT_NUM As Integer
        Dim PO_NUM As Integer
        Dim STANDARD_NUM As Integer
        Dim DIFFERENCE_NUM As Integer
        Dim IN_PLAN_NUM As Integer
        Dim OUT_PLAN_NUM As Integer
        Dim IN_PREDICTION_NUM As Integer
        Dim OUT_PREDICTION_NUM As Integer
    End Structure

    Public Structure Standardnum_Import_List
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim PL_NAME As String
        Dim JAN As String
        Dim STANDARD_NUM As Integer
    End Structure

    Public Structure Delivery_HeaderList
        Dim ID As Integer
        Dim C_ID As Integer
        Dim SHEET_NO As String
        Dim FIX_DATE As String
        Dim C_NAME As String
        Dim D_NAME As String
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim C_CODE As String
        Dim OUT_DATE As String
        Dim ORDER_NO As String
    End Structure

    Public Structure Delivery_DetailList
        Dim I_ID As Integer
        Dim I_CODE As String
        Dim I_NAME As String
        Dim JAN As String
        Dim NUM As Integer
        Dim UNIT_PRICE As Integer
        Dim REFERENCE_PRICE As Integer
    End Structure

    Public Structure CheckDeliveryID_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim SHEET_NO As String
    End Structure

    Public Structure Lot_List
        Dim OUT_ID As Integer
        Dim NO As Integer
        Dim LOT_NUMBER As String
        Dim WARRANTY_CARD_NUMBER As String
    End Structure

    Public Structure S_Yotei_List
        Dim I_CODE As String
        Dim UNIT_PRICE As Integer
        Dim NUM As Integer
        Dim PURCHASE_PRICE As Integer
        Dim I_ID As Integer
    End Structure

    Public Structure OutShipping_Search_List
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim P_ID As Integer
        Dim C_CODE As String
        Dim C_NAME As String
        Dim D_NAME As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim P_NAME As String
        Dim CUSTOMER_ORDER_NO As String
        Dim OUT_DATE As String
        Dim I_STATUS As String
        Dim NUM As Integer
        Dim PLAN_NUM As Integer
        Dim FIX_NUM As Integer
        Dim STOCK_NUM As Integer
        Dim OUT_NUM As Integer
        Dim STATUS As String
        Dim D_UNIT_PRICE As Integer
        Dim S_STATUS As String
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim UNIT_PRICE As Integer
        Dim U_DATE As String
    End Structure

    Public Structure Claim_List
        Dim STATUS As String
        Dim SHEET_NO As String
        Dim C_CODE As String
        Dim D_NAME As String
        Dim I_CODE As String
        Dim I_NAME As String
        Dim FIX_NUM As Integer
        Dim UNIT_COST As Integer
        Dim FILE_NAME As String
        Dim ORDER_NO As String
        Dim FIX_DATE As String
        Dim P_NAME As String
        Dim COMMENT1 As String
        Dim COMMENT2 As String
        Dim ID As Integer
        Dim I_ID As Integer
        Dim C_ID As Integer
        Dim CLAIM_PRT_DATE As String
        Dim C_NAME As String
        Dim CLAIM_CODE As String
        Dim CLAIM_NO As String
    End Structure

    '***********************************************
    ' ログインの認証チェックを行い、表示させるメニューを取得。
    ' <引数>
    ' loginid : ログインID
    ' password : パスワード
    ' <戻り値>
    ' Disp_List : 表示させるメニューの一覧
    ' result : True（成功） , False(失敗）
    ' errorMessage : エラーメッセージ
    '***********************************************
    Public Function LoginCheck(ByVal loginid As String, _
                               ByVal Password As String, _
                               ByRef Disp_List() As DispID_List, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean

        Dim GetUserData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Dim GetDispData As MySqlDataReader
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID,AUTHORITY FROM M_USER WHERE LOGIN_ID='"
            Command.CommandText &= loginid
            Command.CommandText &= "' AND  PW='"
            Command.CommandText &= Password
            Command.CommandText &= "' AND DEL_FLG=0"

            'データ取得
            GetUserData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If GetUserData.Read Then
                ' レコードが取得できた時の処理
                R_User = CStr(GetUserData("ID"))
                Authority = CStr(GetUserData("AUTHORITY"))

                'COUNT値を取得できたのでClose
                GetUserData.Close()

                '続いて表示させるメニューを取得
                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT FUNCTION_ID,FORM_X,FORM_Y,BUTTON_NO FROM M_DISP_MENU "
                Command.CommandText &= "INNER JOIN M_FUNCTION ON M_FUNCTION.ID=M_DISP_MENU.FUNCTION_ID"
                Command.CommandText &= " WHERE M_FUNCTION.DEL_FLG = 0 AND USER_ID='"
                Command.CommandText &= R_User
                Command.CommandText &= "' ORDER BY FUNCTION_ID"

                'Select実行
                GetDispData = Command.ExecuteReader()

                Do While (GetDispData.Read)
                    'ReDim Preserve Disp_List(0 To Count)
                    ReDim Preserve DISPLIST(0 To Count)
                    ' Disp_List(Count).ID = GetDispData("FUNCTION_ID")
                    DISPLIST(Count).ID = GetDispData("FUNCTION_ID")
                    DISPLIST(Count).FORM_X = GetDispData("FORM_X")
                    DISPLIST(Count).FORM_Y = GetDispData("FORM_Y")
                    DISPLIST(Count).BUTTON_NO = GetDispData("BUTTON_NO")
                    Count += 1
                Loop

                GetDispData.Close()

            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "ログインIDまたはパスワードが誤っています。ご確認のうえ、再度ご入力ください。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try

        Return Result
    End Function

    '***********************************************
    ' 機能名をを取得する。
    ' <引数>
    ' 
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetFunctionList(ByRef ListData() As DispID_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean

        Dim CustmerListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID FROM M_FUNCTION WHERE DEL_FLG=0 ORDER BY ID"

            'データ取得
            CustmerListData = Command.ExecuteReader

            Do While (CustmerListData.Read)
                ReDim Preserve ListData(0 To Count)
                ListData(Count).ID = CustmerListData("ID")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' M_PLACEから有効な倉庫情報を取得する。
    ' <引数>
    ' 
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPLACEList(ByRef ListData() As PLACE_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean

        Dim CustmerListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID,NAME FROM M_PLACE WHERE DEL_FLG=0 ORDER BY ID"

            'データ取得
            CustmerListData = Command.ExecuteReader

            Do While (CustmerListData.Read)
                ReDim Preserve ListData(0 To Count)
                ListData(Count).ID = CustmerListData("ID")
                ListData(Count).NAME = CustmerListData("NAME")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "倉庫データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業コード、納品先名(全件)を取得する。
    ' <引数>
    ' 
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetCustomerNameList(ByRef ListData() As C_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean

        Dim CustmerListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID ,D_NAME AS NAME,DISCOUNT_RATE FROM M_CUSTOMER WHERE DEL_FLG=0 ORDER BY D_NAME"

            'データ取得
            CustmerListData = Command.ExecuteReader

            Do While (CustmerListData.Read)
                ReDim Preserve ListData(0 To Count)
                ListData(Count).ID = CustmerListData("ID")
                ListData(Count).NAME = CustmerListData("NAME")
                ListData(Count).DISCOUNT_RATE = CustmerListData("DISCOUNT_RATE")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "入庫元データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業コードから企業名を取得する（１件）
    ' <引数>
    ' CustomerCode : 企業コード
    ' Type : 1:企業コードからデータを取得、 2:企業IDからデータを取得
    ' <戻り値>
    ' CustomerName : 企業名称
    ' C_Id : 企業ID
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetCustomerName(ByRef SearchCode As String, _
                                    ByRef Type As Integer, _
                                    ByRef CustomerName As String, _
                                    ByRef C_Id As Integer, _
                                    ByRef C_Code As String, _
                                    ByRef C_DISCOUNT_RATE As Decimal, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim CustmerData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            If Type = 1 Then


                Command = Connection.CreateCommand
                Command.CommandText = " SELECT ID AS C_ID,C_CODE,D_NAME AS NAME,DISCOUNT_RATE FROM M_CUSTOMER WHERE C_CODE='"
                Command.CommandText &= SearchCode
                Command.CommandText &= "' AND DEL_FLG=0"
            ElseIf Type = 2 Then
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT ID AS C_ID,C_CODE,D_NAME AS NAME,DISCOUNT_RATE FROM M_CUSTOMER WHERE ID='"
                Command.CommandText &= SearchCode
                Command.CommandText &= "' AND DEL_FLG=0"
            End If


            'データ取得
            CustmerData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If CustmerData.Read Then
                ' レコードが取得できた時の処理
                CustomerName = CustmerData("NAME")
                C_Id = CustmerData("C_ID")
                C_Code = CustmerData("C_CODE")
                C_DISCOUNT_RATE = CustmerData("DISCOUNT_RATE")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した企業コードに該当する企業名がみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 最新のドキュメント№を取得する
    ' <引数>
    ' DocNo : 企業コード
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function DocNo_Check(ByRef DocNo As String, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim CustmerData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim ResultCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM IN_HEADER WHERE SHEET_NO='"
            Command.CommandText &= DocNo
            Command.CommandText &= "';"

            'データ取得
            CustmerData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If CustmerData.Read Then
                ' レコードが取得できた時の処理
                ResultCount = CustmerData("COUNT")

                If ResultCount >= 1 Then
                    ErrorMessage = "入力されたドキュメント№は既に登録済みです。"
                    Result = False
                    Exit Function
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "エラーが発生しました。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function
    '***********************************************
    ' 商品コード(全件)を取得する。
    ' <引数>
    '
    ' <戻り値>
    ' ItemCodeList : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemCodeList(ByRef ItemCodeList() As String, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim ItemCodeListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT I_CODE FROM M_ITEM WHERE DEL_FLG=0 ORDER BY ID"

            'データ取得
            ItemCodeListData = Command.ExecuteReader

            Do While (ItemCodeListData.Read)
                ReDim Preserve ItemCodeList(0 To Count)
                ItemCodeList(Count) = ItemCodeListData("I_CODE")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "商品データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品情報を取得する。
    ' <引数>
    ' Type:パッケージ商品を含めるか含めないか(通常商品のみ：０、パッケージのみ：１、全て：２）
    ' <戻り値>
    ' ItemList : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemList(ByVal Type As Integer, _
                                ByRef ItemList() As Item_List, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim ItemCodeListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID,I_CODE,I_NAME,JAN,PL_CODE FROM M_ITEM WHERE DEL_FLG=0 AND PACKAGE_FLG="
            Command.CommandText &= Type
            Command.CommandText &= " ORDER BY ID"


            'データ取得
            ItemCodeListData = Command.ExecuteReader

            Do While (ItemCodeListData.Read)
                ReDim Preserve ItemList(0 To Count)
                ItemList(Count).ID = ItemCodeListData("ID")
                ItemList(Count).I_CODE = ItemCodeListData("I_CODE")
                ItemList(Count).I_NAME = ItemCodeListData("I_NAME")
                ItemList(Count).JAN = ItemCodeListData("JAN")
                ItemList(Count).PL_CODE = ItemCodeListData("PL_CODE")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                If Type = 2 Then
                    ErrorMessage = "商品データがみつかりません。"
                Else
                    ErrorMessage = "セット商品データがみつかりません。" & vbCr & "商品マスタにセット商品を登録してからセット組を行ってください。"
                End If

                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品コードを元に商品情報を取得する。（１件）
    ' <引数>
    ' ItemCode : 商品コード
    ' Search_Flg : 商品コード（1） or JANコード（2）のどちらで検索するか。
    ' <戻り値>
    ' Item_Data : 取得した商品情報
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemName(ByVal Item_Code As String, _
                                ByVal Search_Flg As Integer, _
                                ByRef Item_Data() As Item_List, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim ItemData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            If Search_Flg = 1 Then
                Sql = "I_CODE='"
            ElseIf Search_Flg = 2 Then
                Sql = "JAN='"
            End If

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT ID,I_CODE,I_NAME,JAN,PRICE,LOCATION,PL_CODE,PURCHASE_PRICE,IMMUNITY_PRICE,REPAIR_PRICE FROM M_ITEM WHERE "
            Command.CommandText &= Sql
            Command.CommandText &= Item_Code
            Command.CommandText &= "' AND DEL_FLG=0"

            'データ取得
            ItemData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If ItemData.Read Then
                ' レコードが取得できた時の処理
                ReDim Preserve Item_Data(0 To 0)
                Item_Data(0).ID = ItemData("ID")
                Item_Data(0).I_NAME = ItemData("I_NAME")
                Item_Data(0).PRICE = ItemData("PRICE")
                Item_Data(0).JAN = ItemData("JAN")
                If IsDBNull(ItemData("LOCATION")) Then
                    Item_Data(0).LOCATION = ""
                Else
                    Item_Data(0).LOCATION = ItemData("LOCATION")
                End If

                Item_Data(0).I_CODE = ItemData("I_CODE")
                Item_Data(0).PL_CODE = ItemData("PL_CODE")

                Item_Data(0).PURCHASE_PRICE = ItemData("PURCHASE_PRICE")
                Item_Data(0).IMMUNITY_PRICE = ItemData("IMMUNITY_PRICE")
                Item_Data(0).REPAIR_PRICE = ItemData("REPAIR_PRICE")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                Exit Function
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品コードを元にセット商品情報を取得する。（１件）
    ' <引数>
    ' ItemCode : 商品コード
    '
    ' Type:パッケージ商品を含めるか含めないか(通常商品のみ：０、パッケージのみ：１、全て：２）
    ' <戻り値>
    ' Item_Data : 取得した商品情報
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetSetItemName(ByVal Item_Code As String, _
                                ByVal Search_Flg As Integer, _
                                ByRef Item_Data() As Item_List, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim ItemData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            If Search_Flg = 1 Then
                Sql = "I_CODE='"
            ElseIf Search_Flg = 2 Then
                Sql = "JAN='"
            End If


            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT ID,I_CODE,I_NAME,JAN,PRICE,LOCATION,PL_CODE FROM M_ITEM WHERE "
            Command.CommandText &= Sql
            Command.CommandText &= Item_Code
            Command.CommandText &= "' AND DEL_FLG=0 AND PACKAGE_FLG=1"

            'データ取得
            ItemData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If ItemData.Read Then
                ' レコードが取得できた時の処理
                ReDim Preserve Item_Data(0 To 0)
                Item_Data(0).ID = ItemData("ID")
                Item_Data(0).I_NAME = ItemData("I_NAME")
                Item_Data(0).PRICE = ItemData("PRICE")
                Item_Data(0).JAN = ItemData("JAN")
                If IsDBNull(ItemData("LOCATION")) Then
                    Item_Data(0).LOCATION = ""
                Else
                    Item_Data(0).LOCATION = ItemData("LOCATION")
                End If

                Item_Data(0).I_CODE = ItemData("I_CODE")
                Item_Data(0).PL_CODE = ItemData("PL_CODE")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                Exit Function
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function


    '***********************************************
    ' 商品IDを元に商品情報を取得する。（１件）
    ' <引数>
    ' I_ID : 商品ID
    ' <戻り値>
    ' Item_Data : 取得した商品情報
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemInfo(ByVal I_ID As Integer, _
                                ByRef Item_Data() As Item_List, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim ItemData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT M_ITEM.ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_PLINE.NAME AS PL_NAME FROM M_ITEM INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID WHERE M_ITEM.ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND M_ITEM.DEL_FLG=0"

            'データ取得
            ItemData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If ItemData.Read Then
                ' レコードが取得できた時の処理
                ReDim Preserve Item_Data(0 To 0)
                Item_Data(0).ID = ItemData("ID")
                Item_Data(0).I_CODE = ItemData("I_CODE")
                Item_Data(0).I_NAME = ItemData("I_NAME")
                Item_Data(0).PL_NAME = ItemData("PL_NAME")

            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "選択された商品はマスタに登録されていません。"
                Exit Function
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品の基礎ロケーションを取得する。
    ' <引数>
    ' I_ID : 商品ID
    ' <戻り値>
    ' Location : ロケーション
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemLocation(ByVal I_ID As Integer, _
                                ByRef Location As String, _
                                ByRef ItemCode As String, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim ItemLocationData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT I_CODE,LOCATION FROM M_ITEM WHERE ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DEL_FLG=0"

            'データ取得
            ItemLocationData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If ItemLocationData.Read Then
                If IsDBNull(ItemLocationData("LOCATION")) Then
                    Location = ""
                    ItemCode = ItemLocationData("I_CODE")
                Else
                    Location = ItemLocationData("LOCATION")
                    ItemCode = ItemLocationData("I_CODE")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "選択した商品に該当する商品情報がみつかりません。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品コードがマスタに存在するか確認する（１件）
    ' <引数>
    ' ItemCode : 商品コード
    ' <戻り値>
    ' Count : 件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetItemDuplicationCheck(ByVal Item_Code As String, _
                                            ByRef Count As Integer, _
                                            ByRef Result As String, _
                                            ByRef ErrorMessage As String) As Boolean
        Dim ItemData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM M_ITEM WHERE I_CODE='"
            Command.CommandText &= Item_Code
            Command.CommandText &= "' AND DEL_FLG=0"

            'データ取得
            ItemData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If ItemData.Read Then
                ' レコードが取得できた時の処理
                Count = ItemData("COUNT")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "商品情報取得中にエラーが発生しました。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' プロダクトラインコード（ID)がマスタに存在するか確認する（１件）
    ' <引数>
    ' PL_ID : プロダクトラインコード（ID)
    ' <戻り値>
    ' PL_NAME : プロダクトライン名
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GETPLName(ByVal PL_ID As Integer, _
                              ByRef PL_NAME As String, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean
        Dim PLData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT NAME FROM M_PLINE WHERE ID="
            Command.CommandText &= PL_ID
            Command.CommandText &= " AND DEL_FLG=0"

            'データ取得
            PLData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If PLData.Read Then
                ' レコードが取得できた時の処理
                PL_NAME = PLData("NAME")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力したプロダクトラインコードに該当するプロダクトライン名がみつかりません。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業コードがマスタに存在するか確認する（１件）
    ' <引数>
    ' C_Code : 商品コード
    ' <戻り値>
    ' Count : 件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetCustomerDuplicationCheck(ByVal C_Code As String, _
                                                ByRef Count As Integer, _
                                                ByRef Result As String, _
                                                ByRef ErrorMessage As String) As Boolean
        Dim CustomerData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM M_CUSTOMER WHERE C_CODE='"
            Command.CommandText &= C_Code
            Command.CommandText &= "' AND DEL_FLG=0"

            'データ取得
            CustomerData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If CustomerData.Read Then
                ' レコードが取得できた時の処理
                Count = CustomerData("COUNT")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "企業情報取得中にエラーが発生しました。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品を登録する（入荷予定入力）
    ' <引数>
    ' Sheet_No : シート番号
    ' Defect_Type : 不良区分
    ' Category : 種別
    ' In_Date : 予定入荷日
    ' In_Code : 入庫元コード
    ' R_User : ログインしているユーザーID
    ' Place : 倉庫情報
    ' Dt : DataGridViewに登録された商品データ
    ' ItemCode : 商品コード
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function InsItem(ByRef Sheet_No As String, _
                            ByRef Defect_Type As String, _
                            ByRef Category As String, _
                            ByRef In_Date As String, _
                            ByRef In_C_Id As String, _
                            ByRef R_User As String, _
                            ByRef Place As Integer, _
                            ByRef Dt() As Ins_Item_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Id_Data As MySqlDataReader
        Dim Max_Id As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'INテーブル用　コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "INSERT INTO IN_HEADER(SHEET_NO,PLACE_ID,C_ID,U_USER)VALUES('"
            Command.CommandText &= Sheet_No
            Command.CommandText &= "',"
            Command.CommandText &= Place
            Command.CommandText &= ","
            Command.CommandText &= In_C_Id
            Command.CommandText &= ","
            Command.CommandText &= R_User
            Command.CommandText &= ");"

            'INテーブルへデータ登録
            Command.ExecuteNonQuery()

            '上記で登録したデータのIDを取得する。
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT MAX(ID) as ID FROM IN_HEADER"
            Id_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
            If Id_Data.Read Then
                ' レコードが取得できた時の処理
                Max_Id = CStr(Id_Data("ID"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                Exit Function
            End If
            'MAX値を取得できたのでClose
            Id_Data.Close()

            'IN_DETAILテーブル用　コマンド作成
            For Count = 0 To Dt.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO IN_DETAIL(ID,STATUS,CATEGORY,I_ID,I_STATUS,NUM,FIX_NUM,DATE,U_USER,PO_ID)VALUES("
                Command.CommandText &= Max_Id
                Command.CommandText &= ",1,'"
                Command.CommandText &= Category
                Command.CommandText &= "','"
                Command.CommandText &= Dt(Count).ID
                Command.CommandText &= "','"
                Command.CommandText &= Defect_Type
                Command.CommandText &= "',"
                Command.CommandText &= Dt(Count).NUM
                Command.CommandText &= ",0,'"
                Command.CommandText &= In_Date
                Command.CommandText &= "',"
                Command.CommandText &= R_User
                Command.CommandText &= ",0)"
                'Insert実行
                Command.ExecuteNonQuery()
            Next

            'IN_HEADERテーブル、IN_DETAILテーブルに全てINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 入荷予定検索 - 入荷予定、入庫済み情報を取得する。
    ' <引数>
    ' Doc_No : ドキュメント№
    ' Date_From : 入荷予定日From
    ' Date_To : 入荷予定日To
    ' Fix_Date_From : 入荷日From
    ' Fix_Date_To : 入荷日To
    ' Item_Code : 商品コード
    ' Status1 : ステータス（入荷予定）
    ' Status2 : ステータス（入荷済み）
    ' Category1 : 種別（通常入荷）
    ' Category2 : 種別（返品入荷）
    ' Defect_Type1 : 良品
    ' Defect_Type2 : 不良品
    ' Place : 倉庫(1:八潮、2:東久留米）
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetInSeach(ByRef Doc_No As String, _
                               ByRef Date_From As String, _
                               ByRef Date_To As String, _
                               ByRef Fix_Date_From As String, _
                               ByRef Fix_Date_To As String, _
                               ByRef Item_Code As String, _
                               ByRef Status1 As String, _
                               ByRef Status2 As String, _
                               ByRef Category1 As String, _
                               ByRef Category2 As String, _
                               ByRef Defect_Type1 As String, _
                               ByRef Defect_Type2 As String, _
                               ByRef Place As Integer, _
                               ByRef SearchResult() As Search_List, _
                               ByRef Data_Total As Integer, _
                               ByRef Data_Num_Total As Integer, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            '検索条件よりWhereの作成
            '倉庫
            If WhereSql <> "" Then
                WhereSql = " AND IN_HEADER.PLACE_ID ="
                WhereSql &= Place
                WhereSql &= " "
            Else
                WhereSql = " IN_HEADER.PLACE_ID ="
                WhereSql &= Place
                WhereSql &= " "
            End If


            'ドキュメント№
            If Doc_No <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_HEADER.SHEET_NO ='"
                    WhereSql &= Doc_No
                    WhereSql &= "'"
                Else
                    WhereSql = " IN_HEADER.SHEET_NO ='"
                    WhereSql &= Doc_No
                    WhereSql &= "'"
                End If
            End If
            '入荷予定日From
            If Date_From <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.DATE >= '"
                    WhereSql &= Date_From
                    WhereSql &= "'"
                Else
                    WhereSql &= " IN_DETAIL.DATE >= '"
                    WhereSql &= Date_From
                    WhereSql &= "'"
                End If

            End If

            '入荷予定日To
            If Date_To <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.DATE <= '"
                    WhereSql &= Date_To
                    WhereSql &= "'"
                Else
                    WhereSql &= " IN_DETAIL.DATE <= '"
                    WhereSql &= Date_To
                    WhereSql &= "'"
                End If

            End If

            '入荷日From
            If Fix_Date_From <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.FIX_DATE >= '"
                    WhereSql &= Fix_Date_From
                    WhereSql &= "'"
                Else
                    WhereSql &= " IN_DETAIL.FIX_DATE >= '"
                    WhereSql &= Fix_Date_From
                    WhereSql &= "'"
                End If

            End If

            '入荷日To
            If Fix_Date_To <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.FIX_DATE <= '"
                    WhereSql &= Fix_Date_To
                    WhereSql &= "'"
                Else
                    WhereSql &= " IN_DETAIL.FIX_DATE <= '"
                    WhereSql &= Fix_Date_To
                    WhereSql &= "'"
                End If
            End If

            '商品コード
            If Item_Code <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND M_ITEM.I_CODE ='"
                    WhereSql &= Item_Code
                    WhereSql &= "'"
                Else
                    WhereSql &= " M_ITEM.I_CODE ='"
                    WhereSql &= Item_Code
                    WhereSql &= "'"
                End If
            End If

            'ステータスの入荷予定、入荷済みのどちらにもチェックが入っている
            If Status1 = "True" And Status2 = "True" Then
                'If WhereSql <> "" Then
                '    WhereSql &= " AND "
                'End If
                'WhereSql &= " AND IN_DETAIL.STATUS in('1','2')"
            ElseIf Status1 = "True" And Status2 = "False" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.STATUS = 1"
                Else
                    WhereSql &= " IN_DETAIL.STATUS = 1"
                End If

            ElseIf Status1 = "False" And Status2 = "True" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.STATUS = 2"
                Else
                    WhereSql &= " IN_DETAIL.STATUS = 2"
                End If

            End If

            '種別の通常入荷、返品入荷のどちらにもチェックが入っている
            If Category1 = "True" And Category2 = "True" Then
                'If WhereSql <> "" Then
                '    WhereSql &= " AND "
                'End If
                'WhereSql &= " AND IN_DETAIL.CATEGORY in('1','2')"
            ElseIf Category1 = "True" And Category2 = "False" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.CATEGORY = 1"
                Else
                    WhereSql &= " IN_DETAIL.CATEGORY = 1"
                End If

            ElseIf Category1 = "False" And Category2 = "True" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.CATEGORY = 2"
                Else
                    WhereSql &= " IN_DETAIL.CATEGORY = 2"
                End If

            End If

            '不良区分の良品、不良品のどちらにもチェックが入っている
            If Defect_Type1 = "True" And Defect_Type2 = "True" Then
                'If WhereSql <> "" Then
                '    WhereSql &= " AND "
                'End If
                'WhereSql &= " AND IN_DETAIL.I_STATUS in('1','2')"
            ElseIf Defect_Type1 = "True" And Defect_Type2 = "False" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.I_STATUS = 1"
                Else
                    WhereSql &= " IN_DETAIL.I_STATUS = 1"
                End If

            ElseIf Defect_Type1 = "False" And Defect_Type2 = "True" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND IN_DETAIL.I_STATUS = 2"
                Else
                    WhereSql &= " IN_DETAIL.I_STATUS = 2"
                End If
            End If

            'オープン
            Connection.Open()

            'データ件数取得用コマンド作成
            Command = Connection.CreateCommand
            '2012/2/27 原様依頼により、入庫元コード、入庫元名追加に伴い改修
            'Command.CommandText = "SELECT COUNT(*) AS COUNT FROM IN_HEADER INNER JOIN IN_DETAIL INNER JOIN M_ITEM WHERE "
            'Command.CommandText &= "IN_HEADER.ID=IN_DETAIL.ID AND IN_DETAIL.I_ID=M_ITEM.ID "
            'Command.CommandText &= WhereSql
            'Command.CommandText &= " order by IN_HEADER.ID"

            Command.CommandText = "SELECT COUNT(*) AS COUNT FROM IN_HEADER INNER JOIN IN_DETAIL ON IN_HEADER.ID=IN_DETAIL.ID "
            Command.CommandText &= "INNER JOIN M_ITEM ON IN_DETAIL.I_ID=M_ITEM.ID "
            Command.CommandText &= "LEFT JOIN M_CUSTOMER ON IN_HEADER.C_ID = M_CUSTOMER.ID WHERE "
            Command.CommandText &= WhereSql
            Command.CommandText &= " order by IN_HEADER.ID"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(SearchData("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            '2012/2/27 原様依頼により、入庫元コード、入庫元名追加に伴い改修
            'Command.CommandText = "SELECT SUM(NUM) AS TOTAL FROM IN_HEADER INNER JOIN IN_DETAIL INNER JOIN M_ITEM WHERE "
            'Command.CommandText &= "IN_HEADER.ID=IN_DETAIL.ID AND IN_DETAIL.I_ID=M_ITEM.ID "
            'Command.CommandText &= WhereSql
            'Command.CommandText &= " order by IN_HEADER.ID"

            Command.CommandText = "SELECT SUM(NUM) AS TOTAL FROM IN_HEADER INNER JOIN IN_DETAIL ON IN_HEADER.ID=IN_DETAIL.ID "
            Command.CommandText &= "INNER JOIN M_ITEM ON IN_DETAIL.I_ID=M_ITEM.ID "
            Command.CommandText &= "LEFT JOIN M_CUSTOMER ON IN_HEADER.C_ID = M_CUSTOMER.ID WHERE "
            Command.CommandText &= WhereSql
            Command.CommandText &= " order by IN_HEADER.ID"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理

                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0

                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'コマンド作成
            '2012/2/27 原様依頼により、入庫元コード、入庫元名追加に伴い改修
            'Command = Connection.CreateCommand
            'Command.CommandText = "SELECT IN_HEADER.ID, IN_DETAIL.DETAIL_ID,IN_HEADER.SHEET_NO,M_ITEM.I_CODE, M_ITEM.I_NAME,"
            'Command.CommandText &= "IN_DETAIL.I_ID,M_ITEM.JAN, IN_DETAIL.NUM, IN_DETAIL.DATE, IN_DETAIL.FIX_NUM,"
            'Command.CommandText &= "IN_DETAIL.FIX_DATE, IN_DETAIL.STATUS, IN_DETAIL.CATEGORY, IN_DETAIL.I_STATUS, M_ITEM.LOCATION "
            'Command.CommandText &= "FROM(IN_DETAIL) INNER JOIN IN_HEADER ON IN_DETAIL.ID = IN_HEADER.ID INNER JOIN M_ITEM ON IN_DETAIL.I_ID = M_ITEM.ID"
            'Command.CommandText &= WhereSql
            'Command.CommandText &= " order by IN_HEADER.ID"
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT IN_HEADER.ID, IN_DETAIL.DETAIL_ID,IN_HEADER.SHEET_NO,M_ITEM.I_CODE, M_ITEM.I_NAME,"
            Command.CommandText &= " IN_DETAIL.I_ID,M_ITEM.JAN, IN_DETAIL.NUM, IN_DETAIL.DATE, IN_DETAIL.FIX_NUM,IN_DETAIL.FIX_DATE,IN_DETAIL.REMARKS,"
            Command.CommandText &= " IN_DETAIL.STATUS, IN_DETAIL.CATEGORY, IN_DETAIL.I_STATUS, M_ITEM.LOCATION,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME AS C_NAME,IN_HEADER.PLACE_ID,M_PLACE.NAME AS PLACE "
            Command.CommandText &= " FROM(IN_DETAIL) INNER JOIN IN_HEADER ON IN_DETAIL.ID = IN_HEADER.ID"
            Command.CommandText &= "  INNER JOIN M_ITEM ON IN_DETAIL.I_ID = M_ITEM.ID"
            Command.CommandText &= "  INNER JOIN M_PLACE ON IN_HEADER.PLACE_ID = M_PLACE.ID"
            Command.CommandText &= " LEFT JOIN M_CUSTOMER ON IN_HEADER.C_ID = M_CUSTOMER.ID WHERE "
            Command.CommandText &= WhereSql
            Command.CommandText &= " order by IN_HEADER.ID,IN_DETAIL.DETAIL_ID"


            'データ取得
            SearchData = Command.ExecuteReader()

            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)

                SearchResult(Count).ID = SearchData("ID")
                SearchResult(Count).DETAIL_ID = SearchData("DETAIL_ID")
                SearchResult(Count).DOC_NO = SearchData("SHEET_NO")
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                SearchResult(Count).I_ID = SearchData("I_ID")
                SearchResult(Count).JAN_CODE = SearchData("JAN")
                SearchResult(Count).NUM = SearchData("NUM")
                SearchResult(Count).N_DATE = SearchData("DATE")
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                If IsDBNull(SearchData("FIX_DATE")) Then
                    SearchResult(Count).FIX_DATE = ""
                Else
                    SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                End If
                SearchResult(Count).STATUS = SearchData("STATUS")
                SearchResult(Count).CATEGORY = SearchData("CATEGORY")
                SearchResult(Count).DEFECT_TYPE = SearchData("I_STATUS")
                'SearchResult(Count).LOCATION = SearchData("LOCATION")
                If IsDBNull(SearchData("LOCATION")) Then
                    SearchResult(Count).LOCATION = ""
                Else
                    SearchResult(Count).LOCATION = SearchData("LOCATION")
                End If
                '2012/2/27 原様依頼により、入庫元コード、入庫元名追加に伴い以下２行追加

                If IsDBNull(SearchData("C_CODE")) Then
                    SearchResult(Count).C_CODE = ""
                    SearchResult(Count).C_NAME = ""
                Else
                    SearchResult(Count).C_CODE = SearchData("C_CODE")
                    SearchResult(Count).C_NAME = SearchData("C_NAME")
                End If

                '2012/2/27 コメント欄を入庫コメント、在庫コメントにするため、以下追加（入庫コメント）
                'SearchResult(Count).REMARKS = SearchData("REMARKS")
                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(Count).REMARKS = ""
                Else
                    SearchResult(Count).REMARKS = SearchData("REMARKS")
                End If

                SearchResult(Count).PLACE = SearchData("PLACE")
                SearchResult(Count).PLACE_ID = SearchData("PLACE_ID")
                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function
    '***********************************************
    ' 予定変更画面に表示されたデータを修正する。
    ' <引数>
    ' Upd_List : 更新するデータが格納された配列
    ' DocNo : ドキュメント№
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_In_Item(ByRef Upd_List() As Upd_List, _
                                ByRef DocNo As String, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction
            Command = Connection.CreateCommand
            Command.CommandText = "UPDATE IN_HEADER SET SHEET_NO='"
            Command.CommandText &= DocNo
            Command.CommandText &= "' WHERE IN_HEADER.ID="
            Command.CommandText &= Upd_List(0).ID
            Command.CommandText &= ";"
            'UPDATE実行
            Command.ExecuteNonQuery()

            'DETAILテーブル UPDATE
            For Count = 0 To Upd_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE IN_DETAIL SET NUM="
                Command.CommandText &= Upd_List(Count).NUM
                Command.CommandText &= ",DATE='"
                Command.CommandText &= Upd_List(Count).N_DATE
                Command.CommandText &= "',CATEGORY="
                Command.CommandText &= Upd_List(Count).CATEGORY
                Command.CommandText &= ",I_STATUS="
                Command.CommandText &= Upd_List(Count).DEFECT_TYPE
                Command.CommandText &= ",U_USER="
                Command.CommandText &= R_User
                Command.CommandText &= ",U_DATE=Current_Timestamp WHERE DETAIL_ID="
                Command.CommandText &= Upd_List(Count).DETAIL_ID
                Command.CommandText &= ";"
                'UPDATE実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 選択された商品を削除する。
    ' <引数>
    ' Del_List : 削除する商品IDが格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Del_Item(ByRef Del_List() As ItemID_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Item_Check As MySqlDataReader
        Dim Item_Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'DETAILテーブル 削除
            For Count = 0 To Del_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "DELETE FROM IN_DETAIL WHERE DETAIL_ID="
                Command.CommandText &= Del_List(Count).DETAIL_ID
                Command.CommandText &= ";"
                'Delete実行
                Command.ExecuteNonQuery()
            Next

            '上記で削除したDETAIL_IDのINテーブルの情報が0件だったらINテーブルの情報も削除する。
            For Count = 0 To Del_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT COUNT(*) AS COUNT FROM IN_DETAIL WHERE ID="
                Command.CommandText &= Del_List(Count).ID
                Command.CommandText &= ";"
                'Select実行
                Item_Check = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Item_Check.Read Then
                    Item_Count = Item_Check("COUNT")
                    'Close
                    Item_Check.Close()
                    'Item_Countの結果が0件ならINを削除する。
                    If Item_Count = 0 Then
                        'データを削除する。
                        Command = Connection.CreateCommand
                        Command.CommandText = "DELETE FROM IN_HEADER WHERE ID="
                        Command.CommandText &= Del_List(Count).ID
                        Command.CommandText &= ";"
                        'Delete実行
                        Command.ExecuteNonQuery()
                    End If
                End If
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 入庫確定、在庫の登録（更新）を行う。
    ' <引数>
    ' IN_List : 入庫確定をするデータを格納
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function InDefinition(ByRef IN_List() As InDefinition_List, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Stock_Num As MySqlDataReader
        Dim PO_ID_Data As MySqlDataReader
        Dim PO_Data As MySqlDataReader
        Dim NUM As Integer
        Dim TOTAL_NUM As Integer
        Dim PO_ID As Integer

        Dim TMP_FIX_NUM As Integer
        Dim TMP_FIX_DATE As String = Nothing


        '日時を取得
        Dim dtNow As DateTime = DateTime.Now


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To IN_List.Length - 1
                'IN_DETAILをUPDATEするためのコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE IN_DETAIL SET STATUS ='2',FIX_NUM ="
                Command.CommandText &= IN_List(Count).FIX_NUM
                Command.CommandText &= " ,FIX_DATE ='"
                Command.CommandText &= IN_List(Count).FIX_DATE
                Command.CommandText &= "',U_USER ="
                Command.CommandText &= R_User
                Command.CommandText &= " ,REMARKS='"
                Command.CommandText &= IN_List(Count).IN_COMMENT
                Command.CommandText &= "',U_DATE=Current_Timestamp WHERE DETAIL_ID = "
                Command.CommandText &= IN_List(Count).DETAIL_ID
                Command.CommandText &= ";"

                'UPDATEを実行しIN_DETAILテーブルをUPDATEする。
                Command.ExecuteNonQuery()

                'STOCKにデータがあるかチェックし、あれば数量をプラスする。
                'なければ、新たにSTOCKテーブルにデータを追加。
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT ID,NUM FROM STOCK WHERE I_ID="
                Command.CommandText &= IN_List(Count).I_ID
                Command.CommandText &= " AND I_STATUS="
                Command.CommandText &= IN_List(Count).I_STATUS
                Command.CommandText &= " AND PLACE_ID="
                Command.CommandText &= IN_List(Count).PLACE
                Command.CommandText &= ";"
                'データ取得
                Stock_Num = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Stock_Num.Read Then
                    ' レコードが取得できた時の処理
                    NUM = Stock_Num("NUM")
                    '検索終了（使用終了）したので破棄する。
                    Stock_Num.Dispose()
                    '在庫数と今回入庫した数量をプラスする。
                    TOTAL_NUM = NUM + IN_List(Count).FIX_NUM

                    'データが存在するので、STOCKの該当商品レコードの在庫数を更新する。
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE STOCK SET NUM ="
                    Command.CommandText &= TOTAL_NUM
                    Command.CommandText &= ", U_DATE=Current_Timestamp WHERE I_ID="
                    Command.CommandText &= IN_List(Count).I_ID
                    Command.CommandText &= " AND I_STATUS="
                    Command.CommandText &= IN_List(Count).I_STATUS
                    Command.CommandText &= " AND PLACE_ID="
                    Command.CommandText &= IN_List(Count).PLACE
                    Command.CommandText &= ";"
                    'SQLを実行しSTOCKテーブルをUPDATEする。
                    Command.ExecuteNonQuery()
                Else

                    '検索終了（使用終了）したので破棄する。
                    Stock_Num.Dispose()

                    'Elseの場合はSTOCKに該当商品在庫レコードがないので、新規でInsertする。
                    '在庫Tにデータを登録する
                    'STOCKにINSERTするためのコマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,PLACE_ID,U_DATE)VALUES('"
                    Command.CommandText &= IN_List(Count).I_ID
                    Command.CommandText &= "',"
                    Command.CommandText &= IN_List(Count).FIX_NUM
                    Command.CommandText &= ","
                    Command.CommandText &= IN_List(Count).I_STATUS
                    Command.CommandText &= ",'"
                    Command.CommandText &= IN_List(Count).LOCATION
                    Command.CommandText &= "',"
                    Command.CommandText &= IN_List(Count).PLACE
                    Command.CommandText &= ",Current_Timestamp);"

                    'INSERTを実行しSTOCKテーブルへINSERTする。
                    Command.ExecuteNonQuery()
                End If

                'STOCK_LOGに入荷情報をInsertする。
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,REMARKS,I_FLG,U_USER,PLACE_ID)VALUES("
                Command.CommandText &= IN_List(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= IN_List(Count).FIX_NUM
                Command.CommandText &= ","
                Command.CommandText &= IN_List(Count).I_STATUS
                Command.CommandText &= ",'"
                Command.CommandText &= IN_List(Count).STOCK_COMMENT
                Command.CommandText &= "',1,"
                Command.CommandText &= R_User
                Command.CommandText &= ","
                Command.CommandText &= IN_List(Count).PLACE
                Command.CommandText &= ");"
                'INSERTを実行しSTOCK_LOGテーブルへINSERTする。
                Command.ExecuteNonQuery()


                '2012/4/9 改修
                'POからの入荷予定に対して、IN_HEADER.PO_IDに値が入っていればPOからの入荷予定データとして扱い
                'P_ORDERテーブルのFIX_NUM、FIX_DATEをUPDATEする。
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT PO_ID FROM IN_DETAIL WHERE DETAIL_ID="
                Command.CommandText &= IN_List(Count).DETAIL_ID
                'データ取得
                PO_ID_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If PO_ID_Data.Read Then
                    ' レコードが取得できた時の処理
                    PO_ID = PO_ID_Data("PO_ID")
                    '検索終了（使用終了）したので破棄する。
                    PO_ID_Data.Dispose()
                    '0じゃなければ、PO_IDを元にP_ORDERのFIX_NUM,FIX_DATEを取得。
                    If PO_ID <> 0 Then
                        Command = Connection.CreateCommand
                        Command.CommandText = "SELECT FIX_NUM,FIX_DATE FROM P_ORDER WHERE ID="
                        Command.CommandText &= PO_ID
                        'データ取得
                        PO_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                        If PO_Data.Read Then

                            TMP_FIX_NUM = PO_Data("FIX_NUM") + IN_List(Count).FIX_NUM
                            '空じゃなかったらカンマ区切りで日付を追加。
                            If IsDBNull(PO_Data("FIX_DATE")) Then
                                TMP_FIX_DATE = dtNow.ToString("yyyy/MM/dd")
                            Else
                                TMP_FIX_DATE &= PO_Data("FIX_DATE") & "、" & dtNow.ToString("yyyy/MM/dd")
                            End If

                            '検索終了（使用終了）したので破棄する。
                            PO_Data.Dispose()

                            'P_ORDERのFIX_NUM、FIX_DATEをUPDATEする。
                            Command = Connection.CreateCommand
                            Command.CommandText = "UPDATE P_ORDER SET FIX_NUM="
                            Command.CommandText &= TMP_FIX_NUM
                            Command.CommandText &= ",FIX_DATE='"
                            Command.CommandText &= TMP_FIX_DATE
                            Command.CommandText &= "' WHERE ID="
                            Command.CommandText &= PO_ID

                            'INSERTを実行しP_ORDERテーブルへUPDATEする。
                            Command.ExecuteNonQuery()
                        Else
                            ' レコードが取得できなかった時の処理
                            ErrorMessage = "P_ORDERテーブルのデータの取得ができませんでした。"
                            Exit Function
                        End If
                    End If
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "POデータの取得ができませんでした。"
                    Exit Function
                End If
            Next

            'IN_DETAILテーブル、SOTCK、STOCK_LOGに全てINSERTやUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 検品チェックシートに出力する情報の取得。
    ' <引数>
    ' WherePL : プロダクトラインコードSQL用
    ' WhereSQL : DataGridViewでチェックされたOUT_TBL.DETAIL_IDを格納
    ' <戻り値>
    ' In_Check_List_Data : 出力する情報を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOutputList(ByVal WherePL As String, _
                                  ByVal WhereSQL As String, _
                                  ByRef In_Check_List_Data() As In_Check_List, _
                                  ByRef LineCount As Integer, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean
        Dim OutputData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim BaitCount As Integer = 0
        Dim RodsCount As Integer = 0
        Dim ReelsCount As Integer = 0
        Dim AccessoryCount As Integer = 0
        Dim BagCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,IN_DETAIL.NUM,M_PLINE.SHEET_TYPE FROM "
            Command.CommandText &= "IN_HEADER INNER JOIN IN_DETAIL ON IN_HEADER.ID = IN_DETAIL.ID "
            Command.CommandText &= "INNER JOIN M_ITEM ON IN_DETAIL.I_ID= M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
            Command.CommandText &= "WHERE IN_DETAIL.DETAIL_ID IN("
            Command.CommandText &= WhereSQL
            Command.CommandText &= ") AND M_PLINE.ID in("
            Command.CommandText &= WherePL
            Command.CommandText &= ") ORDER BY M_ITEM.I_CODE;"

            'データ取得
            OutputData = Command.ExecuteReader()

            Do While (OutputData.Read)
                ReDim Preserve In_Check_List_Data(0 To LineCount)
                In_Check_List_Data(LineCount).I_CODE = OutputData("I_CODE")
                In_Check_List_Data(LineCount).I_NAME = OutputData("I_NAME")
                In_Check_List_Data(LineCount).JAN = OutputData("JAN")
                In_Check_List_Data(LineCount).NUM = OutputData("NUM")
                In_Check_List_Data(LineCount).SHEET_TYPE = OutputData("SHEET_TYPE")
                LineCount += 1
            Loop
            OutputData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 最新のドキュメント№を取得する。
    ' <引数>
    ' 
    ' <戻り値>
    ' DocNo : 入力済みの最新ドキュメント№
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDocNo(ByRef DocNo As String, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim DocNoData As MySqlDataReader
        Dim CountData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select SHEET_NO from IN_HEADER where ID=(select max(ID) from IN_HEADER)"

            'データ取得
            DocNoData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If DocNoData.Read Then
                ' レコードが取得できた時の処理
                DocNo = CStr(DocNoData("SHEET_NO"))
                DocNoData.Close()
            Else
                DocNoData.Close()
                ' レコードが取得できなかった時はIN_HEADERの件数を確認し、0件ならOK。
                Command = Connection.CreateCommand
                Command.CommandText = "select COUNT(*) AS COUNT from IN_HEADER"
                CountData = Command.ExecuteReader(CommandBehavior.SingleRow)
                If CountData.Read Then
                    If CountData("COUNT") = 0 Then
                        Return True
                        Exit Function
                    Else
                        ErrorMessage = "最新のドキュメント№を取得できませんでした。"
                        Exit Function
                    End If
                Else
                    ErrorMessage = "。データの取得に失敗しました。"
                    Exit Function
                End If
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示ファイルを登録。
    ' <引数>
    ' ExcelData : DataGridViewに取り込まれた出荷指示データ
    ' File_Name : 出荷指示ファイル名
    ' Place : 倉庫ID
    ' <戻り値>
    ' ItemName : 商品名
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_ExcelData(ByRef ExcelData() As ExcelData_List, _
                                  ByRef File_Name As String, _
                                  ByRef Place As Integer, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean


        Dim ItemID As MySqlDataReader
        Dim CustomerID As MySqlDataReader

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim FileName_List() As String = Nothing
        Dim Count As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim PlusFlg As Boolean = False
        Dim MinusFlg As Boolean = False
        Dim ZeroFlg As Boolean = False


        Dim I_ID As Integer = 0
        Dim C_ID As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To ExcelData.Length - 1

                I_ID = 0
                C_ID = 0

                '商品IDを取得する。
                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT ID FROM M_ITEM WHERE I_CODE='"
                Command.CommandText &= ExcelData(Count).I_CODE
                Command.CommandText &= "';"

                'データ取得
                ItemID = Command.ExecuteReader(CommandBehavior.SingleRow)
                If ItemID.Read Then
                    'レコードが取得できた時の処理()
                    If IsDBNull(ItemID("ID")) Then
                        I_ID = 0
                    Else
                        I_ID = ItemID("ID")
                    End If

                End If
                ItemID.Close()


                '納品先IDを取得する。
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT ID FROM M_CUSTOMER WHERE C_CODE='"
                Command.CommandText &= ExcelData(Count).C_CODE
                Command.CommandText &= "';"

                'データ取得
                CustomerID = Command.ExecuteReader(CommandBehavior.SingleRow)
                If CustomerID.Read Then
                    'レコードが取得できた時の処理()
                    If IsDBNull(CustomerID("ID")) Then
                        C_ID = 0
                    Else
                        C_ID = CustomerID("ID")
                    End If

                End If
                CustomerID.Close()

                'OUT_PRTテーブル用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO OUT_PRT(UNIQUE_KEY,FILE_NAME,C_CODE,SHEET_NO,ORDER_NO,I_CODE,"
                Command.CommandText &= "UNIT_COST,NUM,TOTAL_AMOUNT,COMMENT1,COMMENT2,O_DATE,PRT_STATUS,U_DATE,U_USER,R_STATUS,I_ID,C_ID,PLACE_ID)VALUES('"
                Command.CommandText &= ExcelData(Count).KEY
                Command.CommandText &= "','"
                Command.CommandText &= File_Name
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).C_CODE
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).SHEET_NO
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).ORDER_NO
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).I_CODE
                Command.CommandText &= "',"
                Command.CommandText &= ExcelData(Count).UNIT_COST
                Command.CommandText &= ","
                Command.CommandText &= ExcelData(Count).NUM
                Command.CommandText &= ","
                Command.CommandText &= ExcelData(Count).TOTAL_AMOUNT
                Command.CommandText &= ",'"
                Command.CommandText &= ExcelData(Count).COMMENT1
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).COMMENT2
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).O_DATE
                Command.CommandText &= "',1,'"
                Command.CommandText &= DateTime.Today.ToString("yyyy/MM/dd")
                Command.CommandText &= "',"
                Command.CommandText &= R_User
                Command.CommandText &= ",1,"
                Command.CommandText &= I_ID
                Command.CommandText &= ","
                Command.CommandText &= C_ID
                Command.CommandText &= ","
                Command.CommandText &= Place
                Command.CommandText &= ");"

                'OUT_PRTテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next

            '全てのデータをOUT_PRTテーブルにINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示ファイル名を取得する
    ' 用途は出荷指示ファイルを取り込む際、取り込もうとする出荷指示ファイル名称が
    ' すでに登録されているかを確認する。
    ' <引数>
    ' FileName : 出荷指示ファイル名
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function FileName_Check(ByRef FileName As String, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim FNameData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataResult As String

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT DISTINCT(FILE_NAME) AS NAME FROM OUT_PRT WHERE R_STATUS=2 AND FILE_NAME='"
            Command.CommandText &= FileName
            Command.CommandText &= "';"

            'データ取得
            FNameData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If FNameData.Read Then
                ' レコードが取得できた時の処理
                DataResult = FNameData("NAME")

                If DataResult <> "" Then
                    ErrorMessage = "既に取り込まれています。"
                    Result = False
                    Exit Function
                End If
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 伝票番号がDBに存在しているかチェックする。
    ' <引数>
    ' Sheet_No : 取り込む出荷指示ファイルの伝票番号の一覧を格納
    ' <戻り値>
    ' OrderNo_Duplication_Message : 伝票番号を格納したエラーメッセージ
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetSheetNo(ByRef Sheet_No() As SheetNo_List, _
                               ByRef OrderNo_Duplication_Message As String, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SheetData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            For Count = 0 To Sheet_No.Length - 1

                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT DISTINCT(SHEET_NO) AS SHEET_NO FROM OUT_HEADER WHERE SHEET_NO='"
                Command.CommandText &= Sheet_No(Count).SHEET_NO
                Command.CommandText &= "';"

                'データ取得
                SheetData = Command.ExecuteReader(CommandBehavior.SingleRow)
                If SheetData.Read Then
                    ' レコードが存在していたらエラー
                    If OrderNo_Duplication_Message = "" Then
                        OrderNo_Duplication_Message = Sheet_No(Count).SHEET_NO & vbCr
                    Else
                        OrderNo_Duplication_Message &= Sheet_No(Count).SHEET_NO & vbCr
                    End If
                End If
                SheetData.Close()
            Next

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷予定検索 - 出荷予定、出庫済み情報を取得する。
    ' <引数>
    ' File_Name : 出荷指示ファイル名
    ' Sheet_No : ドキュメント№
    ' Order_No : オーダー番号
    ' Date_From : 出荷予定日From
    ' Date_To : 出荷予定日To
    ' Fix_Date_From : 出荷日From
    ' Fix_Date_To : 出荷日To
    ' ItemJan_Flg : 商品コード or JANコードのどちらの検索かを判別するフラグ
    ' ItemJan_Code : 商品コード or JANコード
    ' Comment1 : コメント１（伝票向け）
    ' C_Code : 納品先コード
    ' Claim_Code : 請求先コード
    ' Status1 : ステータス（出荷予定）
    ' Status2 : ステータス（ピッキング済み）
    ' Status3 : ステータス（ピッ金後戻し）
    ' Status4 : ステータス（出荷済み）
    ' Status5 : ステータス（伝票出力のみ）
    ' Category1 : 種別（通常出荷）
    ' Category2 : 種別（返品出荷）
    ' Print_type1 : ピッキングリスト未印刷（未使用）
    ' Print_type2 : ピッキングリスト印刷済（未使用）
    ' I_STATUS1 : 区分（良品）
    ' I_STATUS2 : 区分（不良品）
    ' I_STATUS3 : 区分（保管品）
    ' Place : 倉庫
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOutSearch(ByVal File_Name As String, _
                                 ByVal Sheet_No As String, _
                                 ByVal Order_No As String, _
                                 ByVal Date_From As String, _
                                 ByVal Date_To As String, _
                                 ByVal Fix_Date_From As String, _
                                 ByVal Fix_Date_To As String, _
                                 ByVal ItemJan_Flg As Integer, _
                                 ByVal ItemJan_Code As String, _
                                 ByVal Comment1 As String, _
                                 ByVal C_Code As String, _
                                 ByVal Claim_Code As String, _
                                 ByVal Status1 As String, _
                                 ByVal Status2 As String, _
                                 ByVal Status3 As String, _
                                 ByVal Status4 As String, _
                                 ByVal Status5 As String, _
                                 ByVal Category1 As String, _
                                 ByVal Category2 As String, _
                                 ByVal Print_Type1 As String, _
                                 ByVal Print_Type2 As String, _
                                 ByVal I_Status1 As String, _
                                 ByVal I_Status2 As String, _
                                 ByVal I_Status3 As String, _
                                 ByVal Place As Integer, _
                                 ByRef SearchResult() As Out_Search_List, _
                                 ByRef Data_Total As Integer, _
                                 ByRef Data_Num_Total As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing
        Dim I_StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = ""
            '検索条件よりWhereの作成

            '倉庫（必須条件）
            WhereSql = " OUT_TBL.PLACE_ID ="
            WhereSql &= Place
            WhereSql &= " "

            '出荷指示ファイル名
            If File_Name <> "" Then
                WhereSql &= "AND OUT_TBL.FILE_NAME Like '"
                WhereSql &= File_Name
                WhereSql &= "'"
            End If

            '伝票番号
            If Sheet_No <> "" Then
                WhereSql &= " AND OUT_TBL.SHEET_NO ='"
                WhereSql &= Sheet_No
                WhereSql &= "'"
            End If

            'オーダー番号
            If Order_No <> "" Then
                WhereSql &= " AND OUT_TBL.ORDER_NO ='"
                WhereSql &= Order_No
                WhereSql &= "' "
            End If

            '納品先コード
            If C_Code <> "" Then
                WhereSql &= " AND M_CUSTOMER.C_CODE ='"
                WhereSql &= C_Code
                WhereSql &= "' "
            End If

            '請求先コード
            If Claim_Code <> "" Then
                WhereSql &= " AND M_CUSTOMER.CLAIM_CODE ='"
                WhereSql &= Claim_Code
                WhereSql &= "' "
            End If

            '出荷予定日From
            If Date_From <> "" Then
                WhereSql &= " AND OUT_TBL.DATE >= '"
                WhereSql &= Date_From
                WhereSql &= "'"
            End If

            '出荷予定日To
            If Date_To <> "" Then
                'Toのみ入力されている場合
                WhereSql &= " AND OUT_TBL.DATE <= '"
                WhereSql &= Date_To
                WhereSql &= "'"
            End If

            '出荷日From
            If Fix_Date_From <> "" Then
                WhereSql &= " AND DATE_FORMAT(OUT_TBL.FIX_DATE, '%Y/%m/%d') >= '"
                WhereSql &= Fix_Date_From
                WhereSql &= "'"
            End If

            '出荷日To
            If Fix_Date_To <> "" Then
                WhereSql &= " AND DATE_FORMAT(OUT_TBL.FIX_DATE, '%Y/%m/%d') <= '"
                WhereSql &= Fix_Date_To
                WhereSql &= "'"
            End If

            'ItemJan_Flgが1なら商品コードを検索する。
            If ItemJan_Flg = 1 And ItemJan_Code <> "" Then
                WhereSql &= " AND M_ITEM.I_CODE ='"
                WhereSql &= ItemJan_Code
                WhereSql &= "' "
            ElseIf ItemJan_Flg = 2 And ItemJan_Code <> "" Then
                WhereSql &= " AND M_ITEM.JAN ='"
                WhereSql &= ItemJan_Code
                WhereSql &= "' "
            End If

            'コメント１（備考） : あいまい検索
            If Comment1 <> "" Then
                WhereSql &= " AND OUT_TBL.COMMENT1 like '%"
                WhereSql &= Comment1
                WhereSql &= "%'"
            End If

            'ステータスの出荷予定、ピッキング済み、ピッキング戻し、出荷済みの全てもチェックが入っているか
            '全てチェックなしの場合は条件に入れない。
            If Status1 = "True" And Status2 = "True" And Status3 = "True" And Status4 = "True" And Status5 = "True" Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
            Else
                StatusWhere = ""
                'どれかにチェックが入っていれば、以下の処理を行う。
                If Status1 = True Or Status2 = True Or Status3 = True Or Status4 = True Or Status5 = True Then
                    If Status1 = True Then
                        StatusWhere &= "OUT_TBL.STATUS in (1"
                    End If
                    If Status2 = True And StatusWhere = "" Then
                        StatusWhere &= "OUT_TBL.STATUS in (2"
                    ElseIf Status2 = True And StatusWhere <> "" Then
                        StatusWhere &= ",2"
                    End If
                    If Status3 = True And StatusWhere = "" Then
                        StatusWhere &= "OUT_TBL.STATUS in (3"
                    ElseIf Status3 = True And StatusWhere <> "" Then
                        StatusWhere &= ",3"
                    End If
                    If Status4 = True And StatusWhere = "" Then
                        StatusWhere &= "OUT_TBL.STATUS in (4"
                    ElseIf Status4 = True And StatusWhere <> "" Then
                        StatusWhere &= ",4"
                    End If
                    If Status5 = True And StatusWhere = "" Then
                        StatusWhere &= "OUT_TBL.STATUS in (5"
                    ElseIf Status5 = True And StatusWhere <> "" Then
                        StatusWhere &= ",5"
                    End If
                    StatusWhere &= ")"

                    WhereSql &= " AND " & StatusWhere
                End If
            End If

            '種別の通常出荷、返品出荷のどちらにもチェックが入っている
            If Category1 = "True" And Category2 = "True" Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
            ElseIf Category1 = "True" And Category2 = "False" Then
                WhereSql &= " AND OUT_TBL.CATEGORY = 1"
            ElseIf Category1 = "False" And Category2 = "True" Then
                WhereSql &= " AND OUT_TBL.CATEGORY = 2"
            End If

            '種別の通常出荷、返品出荷のどちらにもチェックが入っている
            If I_Status1 = "True" And I_Status2 = "True" And I_Status3 = "True" Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
            Else
                I_StatusWhere = ""
                'どれかにチェックが入っていれば、以下の処理を行う。
                If I_Status1 = True Or I_Status2 = True Or I_Status3 = True Then
                    If I_Status1 = True Then
                        I_StatusWhere &= "OUT_TBL.I_STATUS in (1"
                    End If
                    If I_Status2 = True And I_StatusWhere = "" Then
                        I_StatusWhere &= "OUT_TBL.I_STATUS in (2"
                    ElseIf I_Status2 = True And I_StatusWhere <> "" Then
                        I_StatusWhere &= ",2"
                    End If
                    If I_Status3 = True And I_StatusWhere = "" Then
                        I_StatusWhere &= "OUT_TBL.I_STATUS in (3"
                    ElseIf I_Status3 = True And I_StatusWhere <> "" Then
                        I_StatusWhere &= ",3"
                    End If
                    I_StatusWhere &= ")"
                    WhereSql &= " AND " & I_StatusWhere
                End If
            End If

            If WhereSql <> "" Then
                WhereSql = "WHERE " & WhereSql
            End If

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= WhereSql
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(SearchData("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select SUM(NUM) AS TOTAL "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= WhereSql

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select OUT_TBL.ID,OUT_TBL.C_ID,OUT_TBL.SHEET_NO,OUT_TBL.ORDER_NO,OUT_TBL.I_ID,"
            Command.CommandText &= "M_ITEM.I_CODE,M_ITEM.I_NAME,OUT_TBL.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME AS C_NAME,"
            Command.CommandText &= "OUT_TBL.NUM,OUT_TBL.FIX_NUM,OUT_TBL.DATE,OUT_TBL.FIX_DATE,OUT_TBL.INQUIRY_NO,"
            Command.CommandText &= "OUT_TBL.FILE_NAME,OUT_TBL.STATUS,OUT_TBL.CATEGORY,OUT_TBL.I_STATUS,M_CUSTOMER.D_ZIP,M_CUSTOMER.D_ADDRESS,M_CUSTOMER.D_TEL,"
            Command.CommandText &= "OUT_TBL.UNIT_PRICE, OUT_TBL.UNIT_COST,DATE_FORMAT(OUT_TBL.PRT_DATE, '%Y/%m/%d') AS PRT_DATE,"
            Command.CommandText &= "OUT_TBL.COMMENT1, OUT_TBL.COMMENT2, OUT_TBL.REMARKS,OUT_TBL.PLACE_ID as P_ID,M_PLACE.NAME AS P_NAME "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_TBL.PLACE_ID=M_PLACE.ID "

            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY OUT_TBL.ID"
            'Select実行
            SearchData = Command.ExecuteReader()


            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)

                'OUT.ID
                SearchResult(Count).ID = SearchData("ID")
                'I_ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                '伝票番号
                SearchResult(Count).SHEET_NO = SearchData("SHEET_NO")
                'オーダー番号
                SearchResult(Count).ORDER_NO = SearchData("ORDER_NO")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '納品先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '納品先名
                SearchResult(Count).C_NAME = SearchData("C_NAME")
                '出荷予定数
                SearchResult(Count).NUM = SearchData("NUM")
                '出荷予定日
                SearchResult(Count).O_DATE = SearchData("DATE")
                '出庫数
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '出荷日
                If IsDBNull(SearchData("FIX_DATE")) Then
                    SearchResult(Count).FIX_DATE = ""
                Else
                    SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                End If
                '出荷指示ファイル名
                SearchResult(Count).FILE_NAME = SearchData("FILE_NAME")
                'ステータス
                SearchResult(Count).STATUS = SearchData("STATUS")
                '種別
                SearchResult(Count).CATEGORY = SearchData("CATEGORY")
                '商品ステータス
                SearchResult(Count).DEFECT_TYPE = SearchData("I_STATUS")
                '納入単価
                SearchResult(Count).PRICE = SearchData("UNIT_PRICE")
                '印刷日
                SearchResult(Count).PRT_DATE = CStr(SearchData("PRT_DATE"))
                '売単価
                SearchResult(Count).COST = SearchData("UNIT_COST")
                'コメント１
                SearchResult(Count).COMMENT1 = SearchData("COMMENT1")
                'コメント２
                SearchResult(Count).COMMENT2 = SearchData("COMMENT2")
                '備考
                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(Count).REMARKS = ""
                Else
                    SearchResult(Count).REMARKS = SearchData("REMARKS")
                End If

                '倉庫
                SearchResult(Count).PLACE = SearchData("P_NAME")
                '倉庫ID
                SearchResult(Count).P_ID = SearchData("P_ID")
                '郵便番号
                If IsDBNull(SearchData("D_ZIP")) Then
                    SearchResult(Count).D_ZIP = ""
                Else
                    SearchResult(Count).D_ZIP = SearchData("D_ZIP")
                End If
                '住所
                If IsDBNull(SearchData("D_ADDRESS")) Then
                    SearchResult(Count).D_ADDRESS = ""
                Else
                    SearchResult(Count).D_ADDRESS = SearchData("D_ADDRESS")
                End If
                'TEL
                If IsDBNull(SearchData("D_TEL")) Then
                    SearchResult(Count).D_TEL = ""
                Else
                    SearchResult(Count).D_TEL = SearchData("D_TEL")
                End If
                '問い合わせ番号
                If IsDBNull(SearchData("INQUIRY_NO")) Then
                    SearchResult(Count).INQUIRY_NO = ""
                Else
                    SearchResult(Count).INQUIRY_NO = SearchData("INQUIRY_NO")
                End If

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 選択された商品を削除する。
    ' <引数>
    ' Del_List : 削除する商品IDが格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Out_Del_Item(ByRef Del_List() As Item_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        'Dim Tran As MySqlTransaction = Nothing
        'Dim Connection As New MySqlConnection
        'Dim Command As MySqlCommand = Nothing
        'Dim Count As Integer
        'Dim Item_Check As MySqlDataReader
        'Dim Item_Count As Integer

        'Try
        '    '接続文字列を設定
        '    Connection.ConnectionString = Constring
        '    'オープン
        '    Connection.Open()
        '    'begin
        '    Tran = Connection.BeginTransaction

        '    'DETAILテーブル 削除
        '    For Count = 0 To Del_List.Length - 1
        '        Command = Connection.CreateCommand
        '        Command.CommandText = "DELETE FROM IN_DETAIL WHERE DETAIL_ID="
        '        Command.CommandText &= Del_List(Count).DETAIL_ID
        '        Command.CommandText &= ";"
        '        'Delete実行
        '        Command.ExecuteNonQuery()
        '    Next

        '    '上記で削除したDETAIL_IDのINテーブルの情報が0件だったらINテーブルの情報も削除する。
        '    For Count = 0 To Del_List.Length - 1
        '        Command = Connection.CreateCommand
        '        Command.CommandText = "SELECT COUNT(*) AS COUNT FROM IN_DETAIL WHERE ID="
        '        Command.CommandText &= Del_List(Count).ID
        '        Command.CommandText &= ";"
        '        'Select実行
        '        Item_Check = Command.ExecuteReader(CommandBehavior.SingleRow)

        '        If Item_Check.Read Then
        '            Item_Count = Item_Check("COUNT")
        '            'Close
        '            Item_Check.Close()
        '            'Item_Countの結果が0件ならINを削除する。
        '            If Item_Count = 0 Then
        '                'データを削除する。
        '                Command = Connection.CreateCommand
        '                Command.CommandText = "DELETE FROM `IN` WHERE ID="
        '                Command.CommandText &= Del_List(Count).ID
        '                Command.CommandText &= ";"
        '                'Delete実行
        '                Command.ExecuteNonQuery()
        '            End If
        '        End If
        '    Next
        '    '全SQLの発行が完了したらコミットを行う。
        '    Tran.Commit()
        'Catch ex As Exception
        '    'エラーが発生した場合、ロールバックを行う。
        '    Tran.Rollback()
        '    ErrorMessage = ex.Message
        '    Result = False
        'Finally
        '    If Connection IsNot Nothing Then
        '        Connection.Close()
        '        Connection.Dispose()
        '    End If
        '    If Command IsNot Nothing Then Command.Dispose()
        'End Try
        'Return Result
    End Function

    '***********************************************
    ' プロダクトライン情報を取得する。
    ' <引数>
    ' 
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPLList(ByRef ListData() As PL_List, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim PL_ListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID ,NAME,SHEET_TYPE FROM M_PLINE WHERE DEL_FLG=0 ORDER BY ID"

            'データ取得
            PL_ListData = Command.ExecuteReader

            Do While (PL_ListData.Read)
                ReDim Preserve ListData(0 To Count)
                ListData(Count).ID = PL_ListData("ID")
                ListData(Count).NAME = PL_ListData("NAME")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "プロダクトラインコードのデータがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' プロダクトラインの検品チェックシート情報を取得する。
    ' <引数>
    ' Type : 伝票タイプ
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPLSheetTypeList(ByRef ListData() As PL_List, _
                                    ByRef Type As String, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim PL_ListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID,NAME FROM M_PLINE WHERE SHEET_TYPE='"
            Command.CommandText &= Type
            Command.CommandText &= "' AND  DEL_FLG=0 ORDER BY ID"

            'データ取得
            PL_ListData = Command.ExecuteReader

            Do While (PL_ListData.Read)
                ReDim Preserve ListData(0 To Count)
                ListData(Count).ID = PL_ListData("ID")
                ListData(Count).NAME = PL_ListData("NAME")
                Count += 1
            Loop
            PL_ListData.Close()

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "プロダクトラインコードのデータがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示ファイルからDBに登録した情報をOUT_TBLのデータとチェックを行う。
    ' <引数>
    ' FileName : 出荷指示ファイル名
    ' Key : 出荷指示ファイルごとのキー
    ' <戻り値>
    ' CheckRsult : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetCheckData(ByRef FileName As String, _
                                 ByRef Key As String, _
                                 ByRef PrtOnlyData As Boolean, _
                                 ByRef CheckRsult() As Check_Rsult_List, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean
        Dim ListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim Wheresql As String = Nothing


        If PrtOnlyData = True Then
            Wheresql = "AND OUT_PRT.FILE_NAME=OUT_TBL.FILE_NAME "
        End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT OUT_PRT.ID,M_ITEM.ID AS I_ID,OUT_TBL.ID AS OUT_ID,M_ITEM.PRICE,OUT_PRT.TOTAL_AMOUNT,"
            Command.CommandText &= " OUT_PRT.I_CODE, M_ITEM.I_NAME, OUT_PRT.NUM AS OUTPRT_NUM, OUT_TBL.NUM AS OUTTBL_NUM,"
            Command.CommandText &= " M_CUSTOMER.ID AS C_ID,OUT_PRT.C_CODE,M_CUSTOMER.C_NAME,OUT_PRT.SHEET_NO,OUT_PRT.ORDER_NO,"
            Command.CommandText &= " OUT_PRT.UNIT_COST,OUT_PRT.COMMENT1,OUT_PRT.COMMENT2,OUT_PRT.O_DATE,OUT_TBL.STATUS FROM "
            Command.CommandText &= " (OUT_PRT) LEFT JOIN M_ITEM ON OUT_PRT.I_CODE=M_ITEM.I_CODE "
            Command.CommandText &= " LEFT JOIN OUT_TBL ON OUT_TBL.SHEET_NO=OUT_PRT.SHEET_NO AND"
            Command.CommandText &= " OUT_TBL.ORDER_NO=OUT_PRT.ORDER_NO AND OUT_TBL.I_ID = M_ITEM.ID "
            Command.CommandText &= Wheresql
            Command.CommandText &= " LEFT JOIN M_CUSTOMER ON M_CUSTOMER.C_CODE = OUT_PRT.C_CODE "
            Command.CommandText &= " WHERE   OUT_PRT.FILE_NAME ='"
            Command.CommandText &= FileName
            Command.CommandText &= "' AND OUT_PRT.UNIQUE_KEY='"
            Command.CommandText &= Key
            Command.CommandText &= "';"


            'データ取得
            ListData = Command.ExecuteReader

            Do While (ListData.Read)
                ReDim Preserve CheckRsult(0 To Count)
                CheckRsult(Count).ID = ListData("ID")

                'OUT_TBL.IDがnullの場合、0を入れる。
                If IsDBNull(ListData("I_ID")) Then
                    CheckRsult(Count).I_ID = 0
                    CheckRsult(Count).I_NAME = ""
                    CheckRsult(Count).PRICE = 0
                Else
                    CheckRsult(Count).I_ID = ListData("I_ID")
                    CheckRsult(Count).I_NAME = ListData("I_NAME")
                    CheckRsult(Count).PRICE = ListData("PRICE")
                End If

                'OUT_TBL.IDがnullの場合、0を入れる。
                If IsDBNull(ListData("OUT_ID")) Then
                    CheckRsult(Count).OUT_ID = 0
                Else
                    CheckRsult(Count).OUT_ID = ListData("OUT_ID")
                End If
                CheckRsult(Count).I_CODE = ListData("I_CODE")

                'C_IDのステータス
                If IsDBNull(ListData("C_ID")) Then
                    CheckRsult(Count).C_ID = 0
                    CheckRsult(Count).C_NAME = ""
                Else
                    CheckRsult(Count).C_ID = ListData("C_ID")
                    CheckRsult(Count).C_NAME = ListData("C_NAME")
                End If
                CheckRsult(Count).C_CODE = ListData("C_CODE")


                If IsDBNull(ListData("OUTTBL_NUM")) Then
                    CheckRsult(Count).OUTTBL_NUM = 0
                Else
                    CheckRsult(Count).OUTTBL_NUM = ListData("OUTTBL_NUM")
                End If
                CheckRsult(Count).OUTPRT_NUM = ListData("OUTPRT_NUM")
                CheckRsult(Count).TOTAL_AMOUNT = ListData("TOTAL_AMOUNT")
                CheckRsult(Count).SHEET_NO = ListData("SHEET_NO")
                CheckRsult(Count).ORDER_NO = ListData("ORDER_NO")
                CheckRsult(Count).UNIT_COST = ListData("UNIT_COST")
                CheckRsult(Count).COMMENT1 = ListData("COMMENT1")
                CheckRsult(Count).COMMENT2 = ListData("COMMENT2")
                CheckRsult(Count).O_DATE = ListData("O_DATE")
                'OUT_TBLのステータス
                If IsDBNull(ListData("STATUS")) Then
                    CheckRsult(Count).STATUS = 0
                Else
                    CheckRsult(Count).STATUS = ListData("STATUS")
                End If
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示ファイル取り込み画面のDataGridViewのデータをOUT_TBL、
    ' OUT_PRTにデータを登録。
    ' <引数>
    ' Out_Data : 登録データ格納配列
    ' File_Name : 出荷指示ファイル名
    ' Place : 倉庫ID
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_Out_Instructions(ByRef Out_Data() As Out_Regist_List, _
                                         ByRef File_Name As String, _
                                         ByRef Place As Integer, _
                                         ByRef Result As String, _
                                         ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Stock_Num As MySqlDataReader
        Dim StockNum As Integer = 0
        Dim NUM As Integer = 0
        Dim TotalNum As Integer = 0
        Dim Out_Num As MySqlDataReader
        Dim OutNum As Integer = 0
        Dim TotalOutNum As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'OUT_TBLテーブル用　コマンド作成
            For Count = 0 To Out_Data.Length - 1
                '計算用変数のクリア
                StockNum = 0
                NUM = 0
                TotalNum = 0
                OutNum = 0

                'ステータスにより処理をかえる。
                '通常出荷データか、重複出荷指示データなら、渡されたデータをOUT_TBLにInsertし
                'OUT_PRTテーブルの該当データのR_STATUSを2にUPDATEする。
                If Out_Data(Count).STATUS = 1 Or Out_Data(Count).STATUS = 2 Then
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO OUT_TBL(FILE_NAME,SHEET_NO,ORDER_NO,I_ID,C_ID,UNIT_PRICE,"
                    Command.CommandText &= " UNIT_COST,NUM,FIX_NUM,COMMENT1,COMMENT2,STATUS,DATE,U_USER,PLACE_ID,PRT_DATE,U_DATE)VALUES('"
                    Command.CommandText &= File_Name
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).SHEET_NO
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).ORDER_NO
                    Command.CommandText &= "',"
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).C_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).PRICE
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).UNIT_COST
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).NUM
                    Command.CommandText &= ",0,'"
                    Command.CommandText &= Out_Data(Count).COMMENT1
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).COMMENT2
                    Command.CommandText &= "',"
                    Command.CommandText &= Out_Data(Count).STATUS
                    Command.CommandText &= ",'"
                    Command.CommandText &= Out_Data(Count).O_DATE
                    Command.CommandText &= "',"
                    Command.CommandText &= R_User
                    Command.CommandText &= ","
                    Command.CommandText &= Place
                    Command.CommandText &= ",'0000/00/00',Current_Timestamp);"
                    'Insert実行
                    Command.ExecuteNonQuery()

                    'OUT_PRTのR_STAUTSを2（OUT_TBLにデータ登録済み）にUPDATEする。
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_PRT SET R_STATUS=2,U_DATE=Current_Timestamp,U_USER="
                    Command.CommandText &= R_User
                    Command.CommandText &= " WHERE ID="
                    Command.CommandText &= Out_Data(Count).ID
                    Command.CommandText &= ";"

                    'Update実行
                    Command.ExecuteNonQuery()

                ElseIf Out_Data(Count).STATUS = 3 Then
                    '出荷予定データに対して数量をひく。
                    '数量をひいて数量がマイナスになる場合は０を入れる。

                    '出荷予定データの予定数量を取得
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT NUM FROM OUT_TBL WHERE OUT_TBL.ID="
                    Command.CommandText &= Out_Data(Count).OUT_ID
                    Command.CommandText &= ";"

                    'Select実行
                    Out_Num = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If Out_Num.Read Then
                        ' レコードが取得できた時の処理
                        OutNum = Out_Num("NUM")
                    Else
                        ' レコードが取得できなかった時の処理
                        ErrorMessage = "出荷指示キャンセルを行うデータがみつかりません。"
                        Exit Function
                    End If
                    '予定数量を取得できたのでClose
                    Out_Num.Close()

                    TotalOutNum = OutNum + (Out_Data(Count).NUM)

                    '予定数量 - 出荷指示キャンセル数量の結果が0より小さかったら0を入れる。
                    If TotalOutNum < 0 Then
                        TotalOutNum = 0
                    End If

                    '数量UPDATEコマンドを作成
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_TBL SET NUM="
                    Command.CommandText &= TotalOutNum
                    Command.CommandText &= " WHERE OUT_TBL.ID="
                    Command.CommandText &= Out_Data(Count).OUT_ID
                    Command.CommandText &= ";"

                    'Update実行
                    Command.ExecuteNonQuery()

                    'OUT_PRTのR_STAUTSを2（OUT_TBLにデータ登録済み）にUPDATEする。
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_PRT SET R_STATUS=2,U_DATE=Current_Timestamp,U_USER="
                    Command.CommandText &= R_User
                    Command.CommandText &= " WHERE ID="
                    Command.CommandText &= Out_Data(Count).ID
                    Command.CommandText &= ";"

                    'Update実行
                    Command.ExecuteNonQuery()

                    'STOCKLOGに出荷指示キャンセルの履歴データの登録（I_FLG=7）
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,U_USER)VALUES("
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).NUM
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).I_STATUS
                    Command.CommandText &= ",7,"
                    Command.CommandText &= R_User
                    Command.CommandText &= ");"

                    'Insert実行
                    Command.ExecuteNonQuery()

                ElseIf Out_Data(Count).STATUS = 4 Then
                    'ステータスが4の時は、キャンセルするデータのステータスが”ピッキング済み”なので
                    '”ピッキング済み”を”ピッキング戻し”にUPDATEし、
                    '在庫数をキャンセル数量分、元に戻して
                    'STOCK_LOGに履歴を書き込む。

                    '出荷指示データに対してステータスを"ピッキング済み"から"ピッキング戻し"へUPDATE
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_TBL SET STATUS=3 WHERE ID="
                    Command.CommandText &= Out_Data(Count).OUT_ID
                    Command.CommandText &= " AND I_STATUS="
                    Command.CommandText &= Out_Data(Count).I_STATUS
                    Command.CommandText &= ";"
                    'Update実行
                    Command.ExecuteNonQuery()

                    '在庫数をキャンセル数量分、元に戻す。
                    '現状の在庫数を取得
                    'データ件数の数量合計を取得するコマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT NUM FROM STOCK WHERE"
                    Command.CommandText &= " I_ID="
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= " AND I_STATUS="
                    Command.CommandText &= Out_Data(Count).I_STATUS
                    Command.CommandText &= ";"
                    'Select実行
                    Stock_Num = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If Stock_Num.Read Then
                        If IsDBNull(Stock_Num("NUM")) Then
                            StockNum = 0
                        Else
                            StockNum = Stock_Num("NUM")
                        End If
                    Else
                        StockNum = 0
                    End If

                    'COUNT値を取得できたのでClose
                    Stock_Num.Close()

                    TotalNum = StockNum - (Out_Data(Count).NUM)

                    '在庫テーブルの数量を元に戻すSQL
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE STOCK SET NUM="
                    Command.CommandText &= TotalNum
                    Command.CommandText &= " WHERE I_ID="
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= " AND I_STATUS="
                    Command.CommandText &= Out_Data(Count).I_STATUS
                    Command.CommandText &= ";"
                    'Update実行
                    Command.ExecuteNonQuery()

                    'OUT_PRTのR_STAUTSを2（OUT_TBLにデータ登録済み）にUPDATEする。
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_PRT SET R_STATUS=2,U_DATE=Current_Timestamp,U_USER="
                    Command.CommandText &= R_User
                    Command.CommandText &= " WHERE ID="
                    Command.CommandText &= Out_Data(Count).ID
                    Command.CommandText &= ";"

                    'Update実行
                    Command.ExecuteNonQuery()

                    'STOCKLOGにピッキング戻しの履歴データの登録
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,U_USER)VALUES("
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).NUM
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).I_STATUS
                    Command.CommandText &= ",6,"
                    Command.CommandText &= R_User
                    Command.CommandText &= ");"

                    'Insert実行
                    Command.ExecuteNonQuery()

                    '伝票出力のみなら
                ElseIf Out_Data(Count).STATUS = 5 Then
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO OUT_TBL(FILE_NAME,SHEET_NO,ORDER_NO,I_ID,C_ID,UNIT_PRICE,"
                    Command.CommandText &= " UNIT_COST,NUM,FIX_NUM,COMMENT1,COMMENT2,STATUS,DATE,U_USER,PLACE,PRT_DATE,U_DATE)VALUES('"
                    Command.CommandText &= File_Name
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).SHEET_NO
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).ORDER_NO
                    Command.CommandText &= "',"
                    Command.CommandText &= Out_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).C_ID
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).PRICE
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).UNIT_COST
                    Command.CommandText &= ","
                    Command.CommandText &= Out_Data(Count).NUM
                    Command.CommandText &= ",0,'"
                    Command.CommandText &= Out_Data(Count).COMMENT1
                    Command.CommandText &= "','"
                    Command.CommandText &= Out_Data(Count).COMMENT2
                    Command.CommandText &= "',"
                    Command.CommandText &= Out_Data(Count).STATUS
                    Command.CommandText &= ",'"
                    Command.CommandText &= Out_Data(Count).O_DATE
                    Command.CommandText &= "',"
                    Command.CommandText &= R_User
                    Command.CommandText &= ","
                    Command.CommandText &= Place
                    Command.CommandText &= ",'0000/00/00',Current_Timestamp);"
                    'Insert実行
                    Command.ExecuteNonQuery()

                    'OUT_PRTのR_STAUTSを2（OUT_TBLにデータ登録済み）にUPDATEする。
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE OUT_PRT SET R_STATUS=2,U_DATE=Current_Timestamp,U_USER="
                    Command.CommandText &= R_User
                    Command.CommandText &= " WHERE ID="
                    Command.CommandText &= Out_Data(Count).ID
                    Command.CommandText &= ";"

                    'Update実行
                    Command.ExecuteNonQuery()

                End If
            Next

            'OUT_TBLテーブル、OUT_DETAILテーブルに全てINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示ファイル内の同一伝票番号内に数量がプラス、マイナスが混在しているかチェック
    ' <引数>
    ' File_Name : 出荷指示ファイル名
    ' Key : ファイル取り込み単位でのユニークキー
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Sheet_No_Duplication_Check(ByVal File_Name As String, _
                                               ByVal Key As String, _
                                               ByRef ListData() As Duplication_Num_List, _
                                               ByRef Result As String, _
                                               ByRef ErrorMessage As String) As Boolean
        Dim SheetNo_ListData As MySqlDataReader
        Dim Num_ListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim FileName_List() As String = Nothing
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim PlusFlg As Boolean = False
        Dim MinusFlg As Boolean = False
        Dim ZeroFlg As Boolean = False

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT DISTINCT(SHEET_NO) AS SHEET_NO"
            Command.CommandText &= " FROM"
            Command.CommandText &= " OUT_PRT WHERE FILE_NAME='"
            Command.CommandText &= File_Name
            Command.CommandText &= "' AND UNIQUE_KEY='"
            Command.CommandText &= Key
            Command.CommandText &= "';"

            'Select実行
            SheetNo_ListData = Command.ExecuteReader()

            Do While (SheetNo_ListData.Read)
                ReDim Preserve FileName_List(0 To Count)
                FileName_List(Count) = SheetNo_ListData("SHEET_NO")
                Count += 1
            Loop
            'MAX値を取得できたのでClose
            SheetNo_ListData.Close()
            i = 0
            '取得してきた伝票番号ごとに数量を取得しチェックを行う。

            For i = 0 To Count - 1
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT NUM"
                Command.CommandText &= " FROM"
                Command.CommandText &= " OUT_PRT WHERE FILE_NAME='"
                Command.CommandText &= File_Name
                Command.CommandText &= "' AND SHEET_NO='"
                Command.CommandText &= FileName_List(i)
                Command.CommandText &= "' AND UNIQUE_KEY='"
                Command.CommandText &= Key
                Command.CommandText &= "';"

                'Select実行
                Num_ListData = Command.ExecuteReader()
                j = 0
                PlusFlg = False
                MinusFlg = False
                ZeroFlg = False
                Do While (Num_ListData.Read)
                    If Num_ListData("NUM") > 0 Then
                        PlusFlg = True
                    ElseIf Num_ListData("NUM") < 0 Then
                        MinusFlg = True
                    ElseIf Num_ListData("NUM") = 0 Then
                        ZeroFlg = True
                    End If
                    j += 1
                Loop
                ReDim Preserve ListData(0 To i)
                If PlusFlg = True And MinusFlg = True Then
                    ListData(i).SHEET_NO = FileName_List(i)
                    ListData(i).RESULT = 2
                ElseIf ZeroFlg = True Then
                    ListData(i).SHEET_NO = FileName_List(i)
                    ListData(i).RESULT = 3
                Else
                    ListData(i).SHEET_NO = FileName_List(i)
                    ListData(i).RESULT = 1
                End If
                Num_ListData.Close()
            Next

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出庫データのコメントを修正する。
    ' <引数>
    ' Upd_List : 更新するデータが格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Out_Update(ByRef Upd_List() As Out_Upd_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'OUT_TBLテーブル UPDATE
            For Count = 0 To Upd_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_TBL SET COMMENT1='"
                Command.CommandText &= Upd_List(Count).COMMENT1
                Command.CommandText &= "',COMMENT2 = '"
                Command.CommandText &= Upd_List(Count).COMMENT2
                Command.CommandText &= "', U_DATE=Current_Timestamp ,U_USER="
                Command.CommandText &= R_User
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Upd_List(Count).ID
                Command.CommandText &= ";"
                'UPDATE実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function
    '***********************************************
    ' 出荷確定を行う（OUT_TBLの数量、出荷日、ステータスの更新のみ）
    ' <引数>
    ' Out_List : 出荷確定をするデータを格納
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function OutDefinition(ByRef Out_List() As OutDefinition_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Out_List.Length - 1
                'OUT_TBLをUPDATEするためのコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_TBL SET STATUS ='4',FIX_NUM ="
                Command.CommandText &= Out_List(Count).FIX_NUM
                Command.CommandText &= " ,FIX_DATE ='"
                Command.CommandText &= Out_List(Count).FIX_DATE
                Command.CommandText &= "',U_DATE=Current_Timestamp,U_USER="
                Command.CommandText &= R_User
                Command.CommandText &= " WHERE ID = "
                Command.CommandText &= Out_List(Count).ID
                Command.CommandText &= ";"

                'SQLを実行しOUT_TBLテーブルをUPDATEする。
                Command.ExecuteNonQuery()

                'STOCK_LOGに出荷済み情報をInsertする。
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,U_USER,PLACE_ID)VALUES("
                Command.CommandText &= Out_List(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= Out_List(Count).FIX_NUM
                Command.CommandText &= ",'"
                Command.CommandText &= Out_List(Count).I_STATUS
                Command.CommandText &= "',11,"
                Command.CommandText &= R_User
                Command.CommandText &= ","
                Command.CommandText &= Out_List(Count).PLACE
                Command.CommandText &= ");"
                'INSERTを実行しSTOCK_LOGテーブルへINSERTする。
                Command.ExecuteNonQuery()

            Next

            'OUT_TBLに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()

        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 在庫数を取得する（１件）
    ' <引数>
    ' I_ID : 商品ID
    ' <戻り値>
    ' NUM : 在庫数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetStockNum(ByVal I_ID As Integer, _
                                ByRef NUM As Integer, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim StockData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT NUM FROM STOCK WHERE I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= ";"

            'データ取得
            StockData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If StockData.Read Then
                ' レコードが取得できた時の処理
                NUM = StockData("NUM")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "商品ID：" & I_ID & "の在庫データがみつかりません。"
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function
    '***********************************************
    ' 出荷予定データのステータスをピッキング済みに更新し
    ' 在庫テーブルから在庫数をひき、
    ' STOCK_LOGテーブルにログを登録。
    ' <引数>
    ' File_Name : 出荷指示ファイル名
    ' PLACE_ID：倉庫
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Out_Picking_Data(ByVal File_Name As String, _
                                     ByVal PLACE_ID As Integer, _
                                     ByRef Result As String, _
                                     ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Stock_Num As MySqlDataReader
        Dim StockNum As Integer = 0
        Dim TotalNum As Integer = 0
        Dim ListCount As Integer = 0
        Dim UpDateCheck As Boolean = True
        Dim LocationData As MySqlDataReader
        Dim LOCATION As String = Nothing
        Dim OUT_List As MySqlDataReader
        Dim OUt_List_Data() As Picking_Prt_List = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'OUT_TBLの該当データのステータスをピッキング済み、印刷日をUPDATEする。
            Command = Connection.CreateCommand
            Command.CommandText = "UPDATE OUT_TBL SET STATUS=2,PRT_DATE=Current_Timestamp,U_DATE=Current_Timestamp,U_USER="
            Command.CommandText &= R_User
            Command.CommandText &= " WHERE FILE_NAME ='"
            Command.CommandText &= File_Name
            Command.CommandText &= "';"
            'UPDATE実行
            Command.ExecuteNonQuery()

            '在庫から出荷数をひくために該当データの一覧をOUT_TBLから取得する。
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT I_ID,I_STATUS,NUM,PLACE_ID FROM OUT_TBL WHERE"
            Command.CommandText &= " FILE_NAME='"
            Command.CommandText &= File_Name
            Command.CommandText &= "';"
            'Select実行
            OUT_List = Command.ExecuteReader()

            '格納
            Do While (OUT_List.Read)
                ReDim Preserve OUt_List_Data(0 To ListCount)
                OUt_List_Data(ListCount).I_ID = OUT_List("I_ID")
                OUt_List_Data(ListCount).I_STATUS = OUT_List("I_STATUS")
                OUt_List_Data(ListCount).NUM = OUT_List("NUM")
                OUt_List_Data(ListCount).PLACE_ID = OUT_List("PLACE_ID")
                ListCount += 1
            Loop
            '値を取得できたのでClose
            OUT_List.Close()

            For Count = 0 To OUt_List_Data.Length - 1
                '在庫から引くために、現在の在庫数を取得
                '在庫数をキャンセル数量分、元に戻す。
                '現状の在庫数を取得
                'データ件数の数量合計を取得するコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT NUM FROM STOCK WHERE"
                Command.CommandText &= " I_ID="
                Command.CommandText &= OUt_List_Data(Count).I_ID
                Command.CommandText &= " AND I_STATUS='"
                Command.CommandText &= OUt_List_Data(Count).I_STATUS
                Command.CommandText &= "' AND PLACE_ID="
                Command.CommandText &= OUt_List_Data(Count).PLACE_ID
                Command.CommandText &= ";"
                'Select実行
                Stock_Num = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Stock_Num.Read Then
                    ' レコードが取得できた時の処理
                    StockNum = Stock_Num("NUM")
                    UpDateCheck = True
                Else
                    StockNum = 0
                    UpDateCheck = False
                End If
                'COUNT値を取得できたのでClose
                Stock_Num.Close()

                TotalNum = StockNum - OUt_List_Data(Count).NUM

                If UpDateCheck = True Then
                    'Trueなら在庫があるのでUPDATE
                    '在庫Tにデータからピッキング後の数量を入れる（UPDATEする）
                    Command = Connection.CreateCommand
                    Command.CommandText = "UPDATE STOCK SET NUM="
                    Command.CommandText &= TotalNum
                    Command.CommandText &= ",U_DATE=Current_Timestamp WHERE I_ID="
                    Command.CommandText &= OUt_List_Data(Count).I_ID
                    Command.CommandText &= " AND I_STATUS='"
                    Command.CommandText &= OUt_List_Data(Count).I_STATUS
                    Command.CommandText &= "' AND PLACE_ID="
                    Command.CommandText &= OUt_List_Data(Count).PLACE_ID
                    Command.CommandText &= ";"
                    'SQLを実行しSTOCKテーブルへUPDATEする。
                    Command.ExecuteNonQuery()
                ElseIf UpDateCheck = False Then
                    'FalseならレコードがないのでINSERT

                    'Insertするのにロケーションが必要となるので、
                    'M_ITEMから基礎ロケーションを取得する。
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT LOCATION FROM M_ITEM WHERE"
                    Command.CommandText &= " ID="
                    Command.CommandText &= OUt_List_Data(Count).I_ID
                    Command.CommandText &= ";"
                    'Select実行
                    LocationData = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If LocationData.Read Then
                        ' レコードが取得できた時の処理
                        If IsDBNull(LocationData("LOCATION")) Then
                            LOCATION = ""
                        Else
                            LOCATION = LocationData("LOCATION")
                        End If
                    Else
                        LOCATION = ""
                    End If
                    'COUNT値を取得できたのでClose
                    LocationData.Close()

                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,PLACE_ID)VALUES("
                    Command.CommandText &= OUt_List_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= TotalNum
                    Command.CommandText &= ",'"
                    Command.CommandText &= OUt_List_Data(Count).I_STATUS
                    Command.CommandText &= "','"
                    Command.CommandText &= LOCATION
                    Command.CommandText &= "',"
                    Command.CommandText &= PLACE_ID
                    Command.CommandText &= ");"

                    'SQLを実行しSTOCKテーブルへUPDATEする。
                    Command.ExecuteNonQuery()

                End If

                'STOCK_LOGにピッキング済み情報をInsertする。
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,U_USER,PLACE_ID)VALUES("
                Command.CommandText &= OUt_List_Data(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= OUt_List_Data(Count).NUM
                Command.CommandText &= ",'"
                Command.CommandText &= OUt_List_Data(Count).I_STATUS
                Command.CommandText &= "',5,"
                Command.CommandText &= R_User
                Command.CommandText &= ","
                Command.CommandText &= PLACE_ID
                Command.CommandText &= ");"
                'INSERTを実行しSTOCK_LOGテーブルへINSERTする。
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 在庫情報を取得する。
    ' <引数>
    ' ItemJanCode : 商品コード or JANコード
    ' ItemJanFlg : 商品コード（1）かJANコードか（2）
    ' PL_Id : プロダクトラインコードID
    ' Defect_type1 : 不良区分（良品）
    ' Defect_type2 : 不良区分（不良品）
    ' ItemName : 商品名
    ' ZeroDataFlg : 数量0のデータを表示するかしないか。(1:表示する、2:表示しない）
    ' Place : 倉庫
    ' <戻り値>
    ' ListData : 結果格納用配列 
    ' Data_Total : 商品数
    ' Data_Num_Total : 総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetStockSearch(ByVal ItemJanCode As String, _
                                   ByVal ItemJanFlg As Integer, _
                                   ByVal PL_Id As Integer, _
                                   ByVal Defect_type1 As String, _
                                   ByVal Defect_type2 As String, _
                                   ByVal Defect_type3 As String, _
                                   ByVal ItemName As String, _
                                   ByVal ZeroDataFlg As Integer, _
                                   ByVal Place As Integer, _
                                   ByRef ListData() As Stock_Search_List, _
                                   ByRef Data_Total As Integer, _
                                   ByRef Data_Num_Total As Integer, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim StockListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Dim WhereSql As String = Nothing

        '検索条件を設定

        '倉庫
        WhereSql = " STOCK.PLACE_ID ="
        WhereSql &= Place
        WhereSql &= " "

        'ItemJan_Flgが1なら商品コードを検索する。
        If ItemJanFlg = 1 And ItemJanCode <> "" Then
            If WhereSql = "" Then
                WhereSql &= " M_ITEM.I_CODE ='"
                WhereSql &= ItemJanCode
                WhereSql &= "' "
            Else
                WhereSql &= " AND M_ITEM.I_CODE ='"
                WhereSql &= ItemJanCode
                WhereSql &= "' "
            End If
        ElseIf ItemJanFlg = 2 And ItemJanCode <> "" Then
            If WhereSql = "" Then
                WhereSql &= " M_ITEM.JAN ='"
                WhereSql &= ItemJanCode
                WhereSql &= "' "
            Else
                WhereSql &= " AND M_ITEM.JAN ='"
                WhereSql &= ItemJanCode
                WhereSql &= "' "
            End If
        End If

        'プロダクトライン
        If PL_Id <> 0 Then
            If WhereSql = "" Then
                WhereSql = " M_PLINE.ID = "
                WhereSql &= PL_Id
            Else
                WhereSql &= " AND M_PLINE.ID="
                WhereSql &= PL_Id
            End If
        End If

        '商品名
        If ItemName <> "" Then
            If WhereSql = "" Then
                WhereSql = " M_ITEM.I_NAME like '%"
                WhereSql &= ItemName
                WhereSql &= "%' "
            Else
                WhereSql &= " AND M_ITEM.I_NAME  like '%"
                WhereSql &= ItemName
                WhereSql &= "%' "
            End If
        End If

        '不良区分の良品、不良品のどちらにもチェックが入っていれば、検索条件には入れない。
        If Defect_type1 = "True" And Defect_type2 = "True" And Defect_type3 = "True" Then
        ElseIf Defect_type1 = "True" Or Defect_type2 = "True" Or Defect_type3 = "True" Then

            If WhereSql = "" Then
                WhereSql = " STOCK.I_STATUS in("
            Else
                WhereSql &= " AND STOCK.I_STATUS in("
            End If


            If Defect_type1 = "True" And Defect_type2 = "True" And Defect_type3 = "False" Then
                WhereSql &= "1,2"
            ElseIf Defect_type1 = "True" And Defect_type2 = "False" And Defect_type3 = "True" Then
                WhereSql &= "1,3"
            ElseIf Defect_type1 = "False" And Defect_type2 = "True" And Defect_type3 = "True" Then
                WhereSql &= "2,3"
            ElseIf Defect_type1 = "True" And Defect_type2 = "False" And Defect_type3 = "False" Then
                WhereSql &= "1"
            ElseIf Defect_type1 = "False" And Defect_type2 = "True" And Defect_type3 = "False" Then
                WhereSql &= "2"
            ElseIf Defect_type1 = "False" And Defect_type2 = "False" And Defect_type3 = "True" Then
                WhereSql &= "3"
            End If
            WhereSql &= ")"
        End If

        'If Defect_type1 = "True" And Defect_type2 = "True" And Defect_type3 = "True" Then
        'ElseIf Defect_type1 = "True" And Defect_type2 = "False" And Defect_type3 = "False" Then
        '    If WhereSql = "" Then
        '        WhereSql = " STOCK.I_STATUS = 1"
        '    Else
        '        WhereSql &= " AND STOCK.I_STATUS = 1"
        '    End If
        'ElseIf Defect_type1 = "False" And Defect_type2 = "True" Then
        '    If WhereSql = "" Then
        '        WhereSql = " STOCK.I_STATUS = 2"
        '    Else
        '        WhereSql &= " AND STOCK.I_STATUS = 2"
        '    End If
        'End If

        '数量0のデータを表示するかしないか。
        '1の場合は数量0のデータを表示するので何もしない。
        '2の場合は0をはぶく条件を追加。
        If ZeroDataFlg = 2 Then
            If WhereSql = "" Then
                WhereSql = " STOCK.NUM <> 0 "
            Else
                WhereSql &= " AND STOCK.NUM <> 0"
            End If
        End If

        If WhereSql <> "" Then
            WhereSql = "WHERE " & WhereSql
        End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT"
            Command.CommandText &= " FROM"
            Command.CommandText &= " STOCK INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID "
            Command.CommandText &= " INNER JOIN M_PLINE ON M_PLINE.ID=M_ITEM.PL_CODE "
            Command.CommandText &= WhereSql

            'Select実行
            StockListData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If StockListData.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(StockListData("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            StockListData.Close()

            Command = Connection.CreateCommand
            Command.CommandText = " SELECT SUM(NUM) AS NUM"
            Command.CommandText &= " FROM"
            Command.CommandText &= " STOCK INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID "
            Command.CommandText &= " INNER JOIN M_PLINE ON M_PLINE.ID=M_ITEM.PL_CODE "
            Command.CommandText &= WhereSql

            'Select実行
            StockListData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If StockListData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(StockListData("NUM")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = StockListData("NUM")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            StockListData.Close()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT STOCK.ID ,STOCK.I_ID, M_ITEM.I_CODE, M_ITEM.I_NAME,M_ITEM.JAN,M_PLINE.NAME AS PL_NAME,"
            Command.CommandText &= " M_ITEM.PACKAGE_FLG,STOCK.NUM, STOCK.LOCATION, STOCK.I_STATUS,STOCK.PLACE_ID,M_PLACE.NAME AS P_NAME, "
            Command.CommandText &= " M_ITEM.PACKAGE_FLG,STOCK.NUM, STOCK.LOCATION, STOCK.I_STATUS,STOCK.PLACE_ID,M_PLACE.NAME AS P_NAME, "
            Command.CommandText &= " (SELECT IFNULL( SUM(PLAN_NUM) , 0 ) FROM OUT_SHIPPING_PLAN WHERE STOCK.I_ID = OUT_SHIPPING_PLAN.I_ID AND OUT_SHIPPING_PLAN.PLACE_ID=STOCK.PLACE_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.STATUS='通常出荷' AND OUT_SHIPPING_PLAN.S_STATUS =  '出荷予定') as SHIPPING_NUM, "
            Command.CommandText &= " (SELECT IFNULL( SUM(NUM) , 0 ) FROM OUT_TBL WHERE OUT_TBL.I_ID = STOCK.I_ID AND OUT_TBL.PLACE_ID=STOCK.PLACE_ID AND OUT_TBL.I_STATUS = STOCK.I_STATUS AND OUT_TBL.STATUS='出荷予定') as OUT_NUM "

            Command.CommandText &= " FROM STOCK INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID "
            Command.CommandText &= " INNER JOIN M_PLACE ON STOCK.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= " INNER JOIN M_PLINE ON M_PLINE.ID=M_ITEM.PL_CODE "
            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY M_ITEM.I_CODE "


            'データ取得
            StockListData = Command.ExecuteReader

            ReDim Preserve ListData(0 To Data_Total - 1)

            Do While (StockListData.Read)
                ListData(Count).ID = StockListData("ID")
                ListData(Count).I_ID = StockListData("I_ID")
                ListData(Count).I_CODE = StockListData("I_CODE")
                ListData(Count).I_NAME = StockListData("I_NAME")
                ListData(Count).JAN = StockListData("JAN")
                ListData(Count).NUM = StockListData("NUM")
                If IsDBNull(StockListData("LOCATION")) Then
                    ListData(Count).LOCATION = ""
                Else
                    ListData(Count).LOCATION = StockListData("LOCATION")
                End If
                ListData(Count).I_STATUS = StockListData("I_STATUS")
                ListData(Count).PL_NAME = StockListData("PL_NAME")
                ListData(Count).PACKAGE_FLG = StockListData("PACKAGE_FLG")
                ListData(Count).P_ID = StockListData("PLACE_ID")
                ListData(Count).PLACE = StockListData("P_NAME")

                ListData(Count).SHIPPING_NUM = StockListData("SHIPPING_NUM")
                ListData(Count).OUT_NUM = StockListData("OUT_NUM")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "在庫データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 棚卸兼在庫調整
    ' 棚卸。数量の修正を行う。
    ' STOCKテーブルの数量を修正し
    ' STOCK_LOGに履歴を書き込む。
    ' <引数>
    ' Tanaoroshi_Data : 棚卸データ格納配列
    ' Type : 処理のタイプ（棚卸：3、在庫調整：4）
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Tanaoroshi(ByRef Tanaoroshi_Data() As Tanaoroshi_List, _
                               ByRef Type As Integer, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim NUM As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Tanaoroshi_Data.Length - 1
                'STOCKテーブルUPDATE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " UPDATE STOCK SET NUM="
                Command.CommandText &= Tanaoroshi_Data(Count).NEW_NUM
                Command.CommandText &= ", U_DATE=Current_Timestamp WHERE ID="
                Command.CommandText &= Tanaoroshi_Data(Count).STOCK_ID
                Command.CommandText &= ";"

                'STOCKテーブルへデータ登録
                Command.ExecuteNonQuery()

                NUM = Tanaoroshi_Data(Count).NEW_NUM - Tanaoroshi_Data(Count).NUM

                'STOCK_LOGテーブルINSERTE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,U_USER,U_DATE,REMARKS,PLACE_ID)VALUES("
                Command.CommandText &= Tanaoroshi_Data(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= NUM
                Command.CommandText &= ","
                Command.CommandText &= Tanaoroshi_Data(Count).I_STATUS
                Command.CommandText &= ","
                Command.CommandText &= Type
                Command.CommandText &= ",NULL,"
                Command.CommandText &= R_User
                Command.CommandText &= ",Current_Timestamp,'"
                Command.CommandText &= Tanaoroshi_Data(Count).REMARKS
                Command.CommandText &= "','"
                Command.CommandText &= Tanaoroshi_Data(Count).PLACE_ID
                Command.CommandText &= "');"

                'STOCK_LOGテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next

            '全てのデータをSTOCK,STOCK_LOGテーブルにUPDATE,INSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' セット組み換えを行う。
    ' 組み換え前商品の数量をひき、パッケージ商品のデータを追加・更新
    ' STOCK_LOGに履歴を書き込む。
    ' <引数>
    ' Set_Data : 棚卸データ格納配列
    ' I_ID : セットにする商品ID
    ' Set_Num : セット商品の数量
    ' Location ： ロケーション情報
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_Set(ByRef Set_Data() As Set_Item_List, _
                            ByRef I_Id As Integer, _
                            ByRef Set_Num As Integer, _
                            ByRef Location As String, _
                            ByRef Place As Integer, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim Package_Data As MySqlDataReader
        Dim Package_I_ID_Count As Integer = 0
        Dim Num As Integer = 0
        Dim TotalNum As Integer = 0
        Dim Stock_Id As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Set_Data.Length - 1

                'STOCKテーブルUPDATE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " UPDATE STOCK SET NUM="
                Command.CommandText &= Set_Data(Count).STOCK_NUM
                Command.CommandText &= ", U_DATE=Current_Timestamp WHERE ID="
                Command.CommandText &= Set_Data(Count).STOCK_ID
                Command.CommandText &= ";"

                'STOCKテーブルへデータ更新
                Command.ExecuteNonQuery()

                'STOCK_LOGテーブルINSERTE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,PLACE_ID,U_USER,U_DATE)VALUES("
                Command.CommandText &= Set_Data(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= (Set_Data(Count).STOCK_NUM - Set_Data(Count).BEFORE_NUM)
                Command.CommandText &= ","
                Command.CommandText &= Set_Data(Count).I_STATUS
                Command.CommandText &= ",2,"
                Command.CommandText &= I_Id
                Command.CommandText &= ","
                Command.CommandText &= Set_Data(Count).PLACE
                Command.CommandText &= ","
                Command.CommandText &= R_User
                Command.CommandText &= ",Current_Timestamp);"

                'STOCK_LOGテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next
            'セット商品の情報を取得する。
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM STOCK WHERE I_ID="
            Command.CommandText &= I_Id
            Command.CommandText &= " AND PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= ";"

            'Select実行
            Package_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

            If Package_Data.Read Then
                ' レコードが取得できた時の処理
                Package_I_ID_Count = Package_Data("COUNT")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'ID値を取得できたのでClose
            Package_Data.Close()

            'Package_I_ID_Countの値が０なら在庫に存在していないのでInsertを行う。
            If Package_I_ID_Count = 0 Then
                'セット商品をSTOCKテーブルに登録する。
                Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,PLACE_ID,U_DATE)VALUES("
                Command.CommandText &= I_Id
                Command.CommandText &= ","
                Command.CommandText &= Set_Num
                Command.CommandText &= ",1,'"
                Command.CommandText &= Location
                Command.CommandText &= "',"
                Command.CommandText &= Place
                Command.CommandText &= ",Current_Timestamp);"

                'STOCKテーブルへデータ登録
                Command.ExecuteNonQuery()

                'STOCK_LOGテーブルINSERTE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,PLACE_ID,U_USER,U_DATE)VALUES("
                Command.CommandText &= I_Id
                Command.CommandText &= ","
                Command.CommandText &= Set_Num
                Command.CommandText &= ",1,2,NULL,"
                Command.CommandText &= Place
                Command.CommandText &= ","
                Command.CommandText &= R_User
                Command.CommandText &= ",Current_Timestamp);"

                'STOCK_LOGテーブルへデータ登録
                Command.ExecuteNonQuery()

            Else
                'データが存在すれば、UPDATEを行う。
                'セット商品のSTOCK.IDと現在の数量を取得
                Command.CommandText = " SELECT ID,NUM FROM STOCK WHERE I_STATUS=1 AND I_ID="
                Command.CommandText &= I_Id
                Command.CommandText &= " AND PLACE_ID="
                Command.CommandText &= Place
                Command.CommandText &= ";"


                'Select実行
                Package_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Package_Data.Read Then
                    ' レコードが取得できた時の処理
                    Num = Package_Data("NUM")
                    Stock_Id = Package_Data("ID")
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                    Exit Function
                End If
                '数量を取得できたのでClose
                Package_Data.Close()

                TotalNum = Num + Set_Num

                Command.CommandText = "UPDATE STOCK SET NUM="
                Command.CommandText &= TotalNum
                Command.CommandText &= ",U_DATE=Current_Timestamp WHERE ID="
                Command.CommandText &= Stock_Id
                Command.CommandText &= ";"

                'STOCKテーブルへデータ更新
                Command.ExecuteNonQuery()

                'STOCK_LOGテーブルINSERTE用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,PLACE_ID,U_USER,U_DATE)VALUES("
                Command.CommandText &= I_Id
                Command.CommandText &= ","
                Command.CommandText &= Set_Num
                Command.CommandText &= ",1,2,NULL,"
                Command.CommandText &= Place
                Command.CommandText &= ","
                Command.CommandText &= R_User
                Command.CommandText &= ",Current_Timestamp);"

                'STOCK_LOGテーブルへデータ登録
                Command.ExecuteNonQuery()

            End If

            '全てのデータをSTOCK,STOCK_LOGテーブルにUPDATE,INSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 返品出荷登録を行う。
    ' OUT_TBLに出荷済みデータとして登録を行い、
    ' STOCKテーブルの数量を更新し、
    ' STOCK_LOGに履歴を登録する。
    ' <引数>
    ' Out_Data : 登録データ格納配列
    ' File_Name : 出荷指示ファイル名
    ' Sheet_No : 伝票番号
    ' Order_No : オーダー番号
    ' C_Id : 企業ID(返品先ID）
    ' O_Date : 出荷日
    ' Remarks : メモ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_Return_Out(ByRef Return_Out_Data() As Tanaoroshi_List, _
                                   ByRef File_Name As String, _
                                   ByRef Sheet_No As String, _
                                   ByRef Order_No As String, _
                                   ByRef C_Id As String, _
                                   ByRef O_Date As String, _
                                   ByRef Remarks As String, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim StockNum As Integer = 0
        Dim NUM As Integer = 0
        Dim TotalNum As Integer = 0
        Dim OutNum As Integer = 0
        Dim TotalOutNum As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            '件数分ループ
            For Count = 0 To Return_Out_Data.Length - 1

                '計算用変数のクリア
                StockNum = 0
                NUM = 0
                TotalNum = 0
                OutNum = 0
                'OUT_TBLに出荷済みデータとしてInsertする。

                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO OUT_TBL(FILE_NAME,SHEET_NO,ORDER_NO,I_ID,C_ID,UNIT_PRICE,"
                Command.CommandText &= " UNIT_COST,NUM,FIX_NUM,COMMENT1,COMMENT2,STATUS,DATE,FIX_DATE,PLACE_ID,U_USER,PRT_DATE,U_DATE,CATEGORY,I_STATUS)VALUES('"
                Command.CommandText &= File_Name
                Command.CommandText &= "','"
                Command.CommandText &= Sheet_No
                Command.CommandText &= "','"
                Command.CommandText &= Order_No
                Command.CommandText &= "',"
                Command.CommandText &= Return_Out_Data(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= C_Id
                Command.CommandText &= ",0,0,"
                Command.CommandText &= Return_Out_Data(Count).NEW_NUM
                Command.CommandText &= ","
                Command.CommandText &= Return_Out_Data(Count).NEW_NUM
                Command.CommandText &= ",'','',4,'"
                Command.CommandText &= O_Date
                Command.CommandText &= "','"
                Command.CommandText &= O_Date
                Command.CommandText &= "','"
                Command.CommandText &= Return_Out_Data(Count).PLACE_ID
                Command.CommandText &= "',"
                Command.CommandText &= R_User
                Command.CommandText &= ",'0000/00/00',Current_Timestamp,2,2);"
                'Insert実行
                Command.ExecuteNonQuery()

                'STOCKの数量を更新

                TotalOutNum = Return_Out_Data(Count).NUM - Return_Out_Data(Count).NEW_NUM

                '数量UPDATEコマンドを作成
                Command.CommandText = "UPDATE STOCK SET NUM ="
                Command.CommandText &= TotalOutNum
                Command.CommandText &= " WHERE I_ID ="
                Command.CommandText &= Return_Out_Data(Count).I_ID
                Command.CommandText &= " AND PLACE_ID='"
                Command.CommandText &= Return_Out_Data(Count).PLACE_ID
                Command.CommandText &= "' AND I_STATUS=2;"

                'Update実行
                Command.ExecuteNonQuery()

                'STOCK_LOGに返品出荷の履歴データの登録（I_FLG=8）
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,U_USER,REMARKS,PLACE_ID)VALUES("
                Command.CommandText &= Return_Out_Data(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= Return_Out_Data(Count).NEW_NUM
                Command.CommandText &= ",2,8,"
                Command.CommandText &= R_User
                Command.CommandText &= ",'"
                Command.CommandText &= Remarks
                Command.CommandText &= "','"
                Command.CommandText &= Return_Out_Data(Count).PLACE_ID
                Command.CommandText &= "');"
                'Insert実行
                Command.ExecuteNonQuery()
            Next

            'OUT_TBLテーブル、STOCKテーブル、STOCK_LOGテーブルに全てINSERT,UPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 良品変更を行う。
    '
    ' 現在の在庫数から変更後数量を引き、
    ' 変更データがSTOCKに存在すれば数量を加算、
    ' 変更データがSTOCKになければ新規データ登録し
    ' 作業履歴をSTOCK_LOGに書き込む。
    ' <引数>
    ' IStatus_Change_Data : 変更データ格納配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function IStatus_Change(ByRef IStatus_Change_Data() As IStatus_Change_List, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim Check_Data As MySqlDataReader
        Dim Get_Stock_Num As MySqlDataReader
        '現在個数 - 変更数量の結果を格納
        Dim Num As Integer = 0
        Dim DataCount As Integer = 0
        '変更データが存在する場合の現在数量を格納
        Dim UpdateNum As Integer = 0
        Dim UpdateStock_Id As Integer = 0
        '取得してきた在庫数を格納
        Dim StockNum As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'データ件数ループ
            For Count = 0 To IStatus_Change_Data.Length - 1

                '計算用変数のクリア
                Num = 0

                '在庫から変更数量をひく（STOCKをUPDATE)
                Num = IStatus_Change_Data(Count).NUM - IStatus_Change_Data(Count).CHANGE_NUM

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE STOCK SET NUM="
                Command.CommandText &= Num
                Command.CommandText &= ",U_DATE=Current_Timestamp WHERE ID="
                Command.CommandText &= IStatus_Change_Data(Count).STOCK_ID
                Command.CommandText &= ";"

                'UPDATE実行
                Command.ExecuteNonQuery()

                '変更データがSTOCKテーブルに既に存在しているか確認。（商品ID（I_ID)と商品ステータス（I_STATUS)で確認）
                Command.CommandText = " SELECT COUNT(*) AS COUNT FROM STOCK WHERE I_STATUS="
                Command.CommandText &= IStatus_Change_Data(Count).CHANGE_I_STATUS
                Command.CommandText &= " AND I_ID="
                Command.CommandText &= IStatus_Change_Data(Count).I_ID
                Command.CommandText &= " AND PLACE_ID='"
                Command.CommandText &= IStatus_Change_Data(Count).PLACE_ID
                Command.CommandText &= "';"

                'Select実行
                Check_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Check_Data.Read Then
                    ' レコードが取得できた時の処理
                    DataCount = Check_Data("COUNT")
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "在庫データが取得できませんでした。"
                    Exit Function
                End If
                '数量を取得できたのでClose
                Check_Data.Close()

                'DataCountの値が0なら新規データ追加
                If DataCount = 0 Then

                    'STOCKへINSERTするコマンド作成
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,PLACE_ID,U_DATE)VALUES("
                    Command.CommandText &= IStatus_Change_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= IStatus_Change_Data(Count).CHANGE_NUM
                    Command.CommandText &= ","
                    Command.CommandText &= IStatus_Change_Data(Count).CHANGE_I_STATUS
                    Command.CommandText &= ",'"
                    Command.CommandText &= IStatus_Change_Data(Count).NEW_LOCATION
                    Command.CommandText &= "','"
                    Command.CommandText &= IStatus_Change_Data(Count).PLACE_ID
                    Command.CommandText &= "',Current_Timestamp);"
                    'Insert実行
                    Command.ExecuteNonQuery()

                    '保管品に変更の場合はログに書き込まない（原さんに確認済み）
                    If IStatus_Change_Data(Count).CHANGE_I_STATUS <> 3 Then

                        'STOCK_LOGへ書き込み。
                        Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PLACE_ID,U_USER,U_DATE)VALUES("
                        Command.CommandText &= IStatus_Change_Data(Count).I_ID
                        Command.CommandText &= ","
                        Command.CommandText &= IStatus_Change_Data(Count).CHANGE_NUM
                        Command.CommandText &= ","
                        Command.CommandText &= IStatus_Change_Data(Count).CHANGE_I_STATUS
                        Command.CommandText &= ",9,'"
                        Command.CommandText &= IStatus_Change_Data(Count).PLACE_ID
                        Command.CommandText &= "',"
                        Command.CommandText &= R_User
                        Command.CommandText &= ",Current_Timestamp);"
                        'Insert実行
                        Command.ExecuteNonQuery()

                    End If

                Else
                    'DataCountが0じゃなければ（データがあれば）現在のデータにUPDATEを行う。

                    '現在の数量を取得
                    Command.CommandText = " SELECT ID,NUM FROM STOCK WHERE I_STATUS="
                    Command.CommandText &= IStatus_Change_Data(Count).CHANGE_I_STATUS
                    Command.CommandText &= " AND I_ID="
                    Command.CommandText &= IStatus_Change_Data(Count).I_ID
                    Command.CommandText &= " AND PLACE_ID='"
                    Command.CommandText &= IStatus_Change_Data(Count).PLACE_ID
                    Command.CommandText &= "';"
                    'Select実行
                    Get_Stock_Num = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If Get_Stock_Num.Read Then
                        ' レコードが取得できた時の処理
                        UpdateStock_Id = Get_Stock_Num("ID")
                        UpdateNum = Get_Stock_Num("NUM")
                    Else
                        ' レコードが取得できなかった時の処理
                        ErrorMessage = "在庫データが取得できませんでした。"
                        Exit Function
                    End If
                    '数量を取得できたのでClose
                    Get_Stock_Num.Close()

                    StockNum = UpdateNum + IStatus_Change_Data(Count).CHANGE_NUM

                    'STOCKへUPDATEするコマンド作成
                    Command.CommandText = "UPDATE STOCK SET NUM ="
                    Command.CommandText &= StockNum
                    Command.CommandText &= ", U_DATE = Current_Timestamp WHERE ID="
                    Command.CommandText &= UpdateStock_Id
                    Command.CommandText &= ";"
                    'UPDATE実行
                    Command.ExecuteNonQuery()

                    If IStatus_Change_Data(Count).CHANGE_I_STATUS <> 3 Then

                        'STOCK_LOGへ書き込み。
                        Command.CommandText = "INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PLACE_ID,U_USER,U_DATE)VALUES("
                        Command.CommandText &= IStatus_Change_Data(Count).I_ID
                        Command.CommandText &= ","
                        Command.CommandText &= IStatus_Change_Data(Count).CHANGE_NUM
                        Command.CommandText &= ","
                        Command.CommandText &= IStatus_Change_Data(Count).CHANGE_I_STATUS
                        Command.CommandText &= ",9,'"
                        Command.CommandText &= IStatus_Change_Data(Count).PLACE_ID
                        Command.CommandText &= "',"
                        Command.CommandText &= R_User
                        Command.CommandText &= ",Current_Timestamp);"
                        'Insert実行
                        Command.ExecuteNonQuery()
                    End If
                End If
            Next

            'STOCK,STOCK_LOGテーブルに全てINSERT,UPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' ピッキングリスト作成に必要な情報を取得する。
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' PL_ID : プロダクトラインID
    ' <戻り値>
    ' Picking_Prt_List_Data : ピッキングリスト出力情報を格納
    ' DataCount : データ件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPickingPrtData(ByVal FILE_NAME As String, _
                                      ByVal PL_ID As Integer, _
                                      ByRef Picking_Prt_List_Data() As Picking_Prt_List, _
                                      ByRef ReutrnDataCount As Integer, _
                                      ByRef Result As String, _
                                      ByRef ErrorMessage As String) As Boolean
        Dim PickingListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataCount As Integer = 0
        Dim NumCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.FILE_NAME,M_PLINE.ID AS PL_ID,M_PLINE.NAME AS PL_NAME,STOCK.LOCATION,"
            Command.CommandText &= "M_ITEM.ID AS I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,STOCK.NUM AS STOCK_NUM,OUT_TBL.I_STATUS,"
            Command.CommandText &= "(SELECT SUM(OUT_TBL.NUM) FROM OUT_TBL INNER JOIN M_ITEM AS ITEM ON OUT_TBL.I_ID = ITEM.ID "
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME ='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND ITEM.ID = M_ITEM.ID ) AS NUM "
            Command.CommandText &= " FROM (OUT_TBL) LEFT JOIN STOCK ON OUT_TBL.I_ID=STOCK.I_ID AND OUT_TBL.I_STATUS=STOCK.I_STATUS "
            Command.CommandText &= " INNER JOIN M_ITEM ON OUT_TBL.I_ID =M_ITEM.ID "
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND M_PLINE.ID="
            Command.CommandText &= PL_ID
            Command.CommandText &= " GROUP BY M_ITEM.I_CODE ORDER BY M_ITEM.I_CODE"

            'データ取得
            PickingListData = Command.ExecuteReader()

            Do While (PickingListData.Read)
                ReDim Preserve Picking_Prt_List_Data(0 To Count)
                Picking_Prt_List_Data(DataCount).ID = PickingListData("ID")
                Picking_Prt_List_Data(DataCount).FILE_NAME = PickingListData("FILE_NAME")
                Picking_Prt_List_Data(DataCount).PL_ID = PickingListData("PL_ID")
                Picking_Prt_List_Data(DataCount).PL_NAME = PickingListData("PL_NAME")
                Picking_Prt_List_Data(DataCount).I_ID = PickingListData("I_ID")
                Picking_Prt_List_Data(DataCount).I_CODE = PickingListData("I_CODE")
                Picking_Prt_List_Data(DataCount).I_NAME = PickingListData("I_NAME")
                Picking_Prt_List_Data(DataCount).NUM = PickingListData("NUM")
                If IsDBNull(PickingListData("STOCK_NUM")) Then
                    Picking_Prt_List_Data(DataCount).STOCK_NUM = 0
                Else
                    Picking_Prt_List_Data(DataCount).STOCK_NUM = PickingListData("STOCK_NUM")
                End If

                Picking_Prt_List_Data(DataCount).I_STATUS = PickingListData("I_STATUS")

                NumCount += PickingListData("NUM")

                If IsDBNull(PickingListData("LOCATION")) Then
                    Picking_Prt_List_Data(DataCount).LOCATION = ""
                Else
                    Picking_Prt_List_Data(DataCount).LOCATION = PickingListData("LOCATION")
                End If
                DataCount += 1
                Count += 1
            Loop

            'MAX値を取得できたのでClose
            PickingListData.Close()
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        ReutrnDataCount = NumCount
        Return Result
    End Function

    '***********************************************
    ' 伝票作成に必要な情報を取得する。
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル
    ' Sheet_Type : 伝票の種類
    ' <戻り値>
    ' Picking_Prt_List_Data : ピッキングリスト出力情報を格納
    'DataCount : データ件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetSlipData(ByVal FILE_NAME As String, _
                                ByVal Sheet_Type As Integer, _
                                ByRef Slip_List_Data() As Slip_List, _
                                ByRef ReutrnDataCount As Integer, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim SlipListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataCount As Integer = 0
        Dim SHEET_NO_List() As Slip_List = Nothing
        Dim End_SlipListData As MySqlDataReader
        Dim COUNT_List() As Slip_List = Nothing
        Dim DataNUM As Integer = 0
        Dim i As Integer = 0
        Dim MAXPAGE As Integer = 0
        Dim PrtPage As Integer = 0
        Dim LoopCount As Integer = 0
        Dim DataCountCheck As MySqlDataReader
        Dim DataCountCheckNum As Integer = 0
        Dim SlipCount As Integer = 0
        Dim DataStartNo As Integer = 0
        Dim DataEndNo As Integer = 0

        '伝票タイプごとの１ページあたりの表示件数を設定
        If Sheet_Type = 1 Then
            DataNUM = 6
        ElseIf Sheet_Type = 2 Then
            DataNUM = 10
        ElseIf Sheet_Type = 3 Then
            DataNUM = 10
        End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'データ数を取得する。
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT COUNT(*) AS COUNT"
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID = M_CUSTOMER.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND M_CUSTOMER.SHEET_TYPE="
            Command.CommandText &= Sheet_Type
            Command.CommandText &= ";"

            'データ取得
            DataCountCheck = Command.ExecuteReader(CommandBehavior.SingleRow)

            If DataCountCheck.Read Then
                ' レコードが取得できた時の処理
                DataCountCheckNum = DataCountCheck("COUNT")
            Else
                DataCountCheckNum = 0
            End If
            DataCountCheck.Close()
            If DataCountCheckNum <> 0 Then

                Command = Connection.CreateCommand
                Command.CommandText = "SELECT DISTINCT(OUT_TBL.SHEET_NO) AS SHEET_NO"
                Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID = M_CUSTOMER.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
                Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
                Command.CommandText &= FILE_NAME
                Command.CommandText &= "' AND M_CUSTOMER.SHEET_TYPE="
                Command.CommandText &= Sheet_Type
                Command.CommandText &= " ORDER BY M_CUSTOMER.C_CODE,OUT_TBL.SHEET_NO;"

                'データ取得
                SlipListData = Command.ExecuteReader()

                Do While (SlipListData.Read)
                    ReDim Preserve SHEET_NO_List(0 To Count)
                    SHEET_NO_List(Count).SHEET_NO = SlipListData("SHEET_NO")
                    Count += 1
                Loop
                SlipListData.Close()

                '取得した伝票番号件数分ループする。
                For i = 0 To SHEET_NO_List.Length - 1
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT OUT_TBL.ID,M_ITEM.ID AS I_ID,OUT_TBL.SHEET_NO,OUT_TBL.ORDER_NO,OUT_TBL.DATE AS O_DATE,M_CUSTOMER.D_ADDRESS,"
                    Command.CommandText &= " M_CUSTOMER.D_NAME,M_CUSTOMER.C_NAME,M_CUSTOMER.C_CODE,M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,"
                    Command.CommandText &= " OUT_TBL.NUM,OUT_TBL.UNIT_COST,OUT_TBL.UNIT_PRICE AS PRICE,OUT_TBL.COMMENT1,COMMENT2"
                    Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
                    Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID = M_CUSTOMER.ID"
                    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
                    Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
                    Command.CommandText &= FILE_NAME
                    Command.CommandText &= "' AND OUT_TBL.SHEET_NO ='"
                    Command.CommandText &= SHEET_NO_List(i).SHEET_NO
                    Command.CommandText &= "' AND  M_CUSTOMER.SHEET_TYPE ="
                    Command.CommandText &= Sheet_Type
                    Command.CommandText &= ";"

                    'データ取得
                    End_SlipListData = Command.ExecuteReader()

                    MAXPAGE = 1
                    PrtPage = 1
                    LoopCount = 1

                    DataStartNo = SlipCount
                    '伝票種類別の一覧の取得ができたら、データ件数分ループ(伝票番号の商品件数分のデータが入る）
                    Do While (End_SlipListData.Read)
                        ReDim Preserve Slip_List_Data(0 To SlipCount)
                        Slip_List_Data(SlipCount).ID = End_SlipListData("ID")
                        Slip_List_Data(SlipCount).I_ID = End_SlipListData("I_ID")
                        Slip_List_Data(SlipCount).SHEET_NO = End_SlipListData("SHEET_NO")
                        Slip_List_Data(SlipCount).ORDER_NO = End_SlipListData("ORDER_NO")
                        Slip_List_Data(SlipCount).O_DATE = End_SlipListData("O_DATE")
                        Slip_List_Data(SlipCount).D_ADDRESS = End_SlipListData("D_ADDRESS")
                        Slip_List_Data(SlipCount).D_NAME = End_SlipListData("D_NAME")
                        Slip_List_Data(SlipCount).C_NAME = End_SlipListData("C_NAME")
                        Slip_List_Data(SlipCount).C_CODE = End_SlipListData("C_CODE")
                        Slip_List_Data(SlipCount).I_CODE = End_SlipListData("I_CODE")
                        Slip_List_Data(SlipCount).I_NAME = End_SlipListData("I_NAME")
                        Slip_List_Data(SlipCount).JAN = End_SlipListData("JAN")
                        Slip_List_Data(SlipCount).NUM = End_SlipListData("NUM")
                        Slip_List_Data(SlipCount).UNIT_COST = End_SlipListData("UNIT_COST")
                        Slip_List_Data(SlipCount).PRICE = End_SlipListData("PRICE")
                        Slip_List_Data(SlipCount).COMMENT1 = End_SlipListData("COMMENT1")
                        Slip_List_Data(SlipCount).COMMENT2 = End_SlipListData("COMMENT2")

                        If LoopCount > DataNUM Then
                            PrtPage += 1
                            LoopCount = 1
                        End If
                        Slip_List_Data(SlipCount).PAGE = PrtPage
                        LoopCount += 1

                        SlipCount += 1
                        DataCount += 1
                    Loop
                    DataEndNo = SlipCount

                    If DataEndNo - DataStartNo <= DataNUM Then
                        MAXPAGE = 1
                    Else
                        If (DataEndNo - DataStartNo) Mod DataNUM = 0 Then
                            MAXPAGE = (DataEndNo - DataStartNo) / DataNUM
                        Else
                            MAXPAGE = (DataEndNo - DataStartNo) \ DataNUM + 1
                        End If
                    End If

                    For k = DataStartNo To DataEndNo - 1
                        Slip_List_Data(k).MAXPAGE = MAXPAGE
                    Next

                    End_SlipListData.Close()
                Next
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        ReutrnDataCount = DataCount
        Return Result
    End Function

    '***********************************************
    ' 伝票作成に必要な情報を取得する。
    ' キャンセルデータの存在するデータで出荷指示ファイル取り込み画面から行う
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル
    ' Sheet_Type : 伝票の種類
    ' <戻り値>
    ' Slip_List_Data : 出荷伝票情報を格納
    ' ReutrnDataCount : データ件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetSlipCancelData(ByVal FILE_NAME As String, _
                                ByVal Sheet_Type As Integer, _
                                ByRef Slip_List_Data() As Slip_List, _
                                ByRef ReutrnDataCount As Integer, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim SlipListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataCount As Integer = 0
        Dim SHEET_NO_List() As Slip_List = Nothing
        Dim End_SlipListData As MySqlDataReader
        Dim COUNT_List() As Slip_List = Nothing
        Dim DataNUM As Integer = 0
        Dim i As Integer = 0
        Dim MAXPAGE As Integer = 0
        Dim PrtPage As Integer = 0
        Dim LoopCount As Integer = 0
        Dim DataCountCheck As MySqlDataReader
        Dim DataCountCheckNum As Integer = 0
        Dim SlipCount As Integer = 0
        Dim DataStartNo As Integer = 0
        Dim DataEndNo As Integer = 0

        '伝票タイプごとの１ページあたりの表示件数を設定
        If Sheet_Type = 1 Then
            DataNUM = 6
        ElseIf Sheet_Type = 2 Then
            DataNUM = 10
        ElseIf Sheet_Type = 3 Then
            DataNUM = 10
        End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'データ数を取得する。
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT COUNT(*) AS COUNT"
            Command.CommandText &= " FROM (OUT_PRT) INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID = M_CUSTOMER.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
            Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND M_CUSTOMER.SHEET_TYPE="
            Command.CommandText &= Sheet_Type
            Command.CommandText &= ";"

            'データ取得
            DataCountCheck = Command.ExecuteReader(CommandBehavior.SingleRow)

            If DataCountCheck.Read Then
                ' レコードが取得できた時の処理
                DataCountCheckNum = DataCountCheck("COUNT")
            Else
                DataCountCheckNum = 0
            End If
            DataCountCheck.Close()
            If DataCountCheckNum <> 0 Then

                Command = Connection.CreateCommand
                Command.CommandText = "SELECT DISTINCT(OUT_PRT.SHEET_NO) AS SHEET_NO"
                Command.CommandText &= " FROM (OUT_PRT) INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID = M_CUSTOMER.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
                Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
                Command.CommandText &= FILE_NAME
                Command.CommandText &= "' AND M_CUSTOMER.SHEET_TYPE="
                Command.CommandText &= Sheet_Type
                Command.CommandText &= " ORDER BY M_CUSTOMER.C_CODE,OUT_PRT.SHEET_NO;"

                'データ取得
                SlipListData = Command.ExecuteReader()

                Do While (SlipListData.Read)
                    ReDim Preserve SHEET_NO_List(0 To Count)
                    SHEET_NO_List(Count).SHEET_NO = SlipListData("SHEET_NO")
                    Count += 1
                Loop
                SlipListData.Close()

                '取得した伝票番号件数分ループする。
                For i = 0 To SHEET_NO_List.Length - 1
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT OUT_PRT.ID,M_ITEM.ID AS I_ID,OUT_PRT.SHEET_NO,OUT_PRT.ORDER_NO,OUT_PRT.O_DATE,M_CUSTOMER.D_ADDRESS,"
                    Command.CommandText &= " M_CUSTOMER.D_NAME,M_CUSTOMER.C_NAME,M_CUSTOMER.C_CODE,M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,"
                    Command.CommandText &= " OUT_PRT.NUM,OUT_PRT.UNIT_COST,M_ITEM.PRICE,OUT_PRT.COMMENT1,OUT_PRT.COMMENT2"
                    Command.CommandText &= " FROM (OUT_PRT) INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
                    Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID = M_CUSTOMER.ID"
                    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID "
                    Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
                    Command.CommandText &= FILE_NAME
                    Command.CommandText &= "' AND OUT_PRT.SHEET_NO ='"
                    Command.CommandText &= SHEET_NO_List(i).SHEET_NO
                    Command.CommandText &= "' AND  M_CUSTOMER.SHEET_TYPE ="
                    Command.CommandText &= Sheet_Type
                    Command.CommandText &= ";"

                    'データ取得
                    End_SlipListData = Command.ExecuteReader()

                    MAXPAGE = 1
                    PrtPage = 1
                    LoopCount = 1

                    DataStartNo = SlipCount
                    '伝票種類別の一覧の取得ができたら、データ件数分ループ(伝票番号の商品件数分のデータが入る）
                    Do While (End_SlipListData.Read)
                        ReDim Preserve Slip_List_Data(0 To SlipCount)
                        Slip_List_Data(SlipCount).ID = End_SlipListData("ID")
                        Slip_List_Data(SlipCount).I_ID = End_SlipListData("I_ID")
                        Slip_List_Data(SlipCount).SHEET_NO = End_SlipListData("SHEET_NO")
                        Slip_List_Data(SlipCount).ORDER_NO = End_SlipListData("ORDER_NO")
                        Slip_List_Data(SlipCount).O_DATE = End_SlipListData("O_DATE")
                        Slip_List_Data(SlipCount).D_ADDRESS = End_SlipListData("D_ADDRESS")
                        Slip_List_Data(SlipCount).D_NAME = End_SlipListData("D_NAME")
                        Slip_List_Data(SlipCount).C_NAME = End_SlipListData("C_NAME")
                        Slip_List_Data(SlipCount).C_CODE = End_SlipListData("C_CODE")
                        Slip_List_Data(SlipCount).I_CODE = End_SlipListData("I_CODE")
                        Slip_List_Data(SlipCount).I_NAME = End_SlipListData("I_NAME")
                        Slip_List_Data(SlipCount).JAN = End_SlipListData("JAN")
                        Slip_List_Data(SlipCount).NUM = End_SlipListData("NUM")
                        Slip_List_Data(SlipCount).UNIT_COST = End_SlipListData("UNIT_COST")
                        Slip_List_Data(SlipCount).PRICE = End_SlipListData("PRICE")
                        Slip_List_Data(SlipCount).COMMENT1 = End_SlipListData("COMMENT1")
                        Slip_List_Data(SlipCount).COMMENT2 = End_SlipListData("COMMENT2")

                        If LoopCount > DataNUM Then
                            PrtPage += 1
                            LoopCount = 1
                        End If
                        Slip_List_Data(SlipCount).PAGE = PrtPage
                        LoopCount += 1

                        SlipCount += 1
                        DataCount += 1
                    Loop
                    DataEndNo = SlipCount

                    If DataEndNo - DataStartNo <= DataNUM Then
                        MAXPAGE = 1
                    Else
                        If (DataEndNo - DataStartNo) Mod DataNUM = 0 Then
                            MAXPAGE = (DataEndNo - DataStartNo) / DataNUM
                        Else
                            MAXPAGE = (DataEndNo - DataStartNo) \ DataNUM + 1
                        End If
                    End If

                    For k = DataStartNo To DataEndNo - 1
                        Slip_List_Data(k).MAXPAGE = MAXPAGE
                    Next

                    End_SlipListData.Close()
                Next
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        ReutrnDataCount = DataCount
        Return Result
    End Function

    '***********************************************
    ' 納品先の一覧を取得する（納品先別出荷リスト向け）
    ' 請求先コードでサマリーするよう対応（2015/4/22）
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' <戻り値>
    ' Delivery_List_Data : 納品先情報を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDeliveryList(ByVal FILE_NAME As String, _
                                    ByRef C_CODE As String, _
                                    ByRef Delivery_List_Data() As Delivery_List, _
                                    ByRef WhereSql As String, _
                                    ByRef DataCount As Integer, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim DeliveryGroupListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DeliveryDataCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            '納品先コードが入っていなければ、M_CUSTOMERのDELIVERY_PRT_FLG=1のデータのみ納品先チェックリストを作成する。
            'If C_CODE = "" Then

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,M_CUSTOMER.C_NAME "
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME ='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND M_CUSTOMER.DELIVERY_PRT_FLG=1 GROUP BY OUT_TBL.C_ID;"

            'データ取得
            DeliveryGroupListData = Command.ExecuteReader()
            'Else
            '    'C_CODEにコードが入っていれば、指定のコードの情報のみの納品先チェックリストを作成する。
            '    'コマンド作成
            '    Command = Connection.CreateCommand
            '    Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID"
            '    Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            '    Command.CommandText &= " WHERE OUT_TBL.FILE_NAME ='"
            '    Command.CommandText &= FILE_NAME
            '    Command.CommandText &= "' AND M_CUSTOMER.C_CODE='"
            '    Command.CommandText &= C_CODE
            '    'Command.CommandText &= "' GROUP BY OUT_TBL.C_ID;"
            '    Command.CommandText &= "' GROUP BY M_CUSTOMER.CLAIM_CODE;"

            '    'データ取得
            '    DeliveryGroupListData = Command.ExecuteReader()
            'End If

            Do While (DeliveryGroupListData.Read)
                ReDim Preserve Delivery_List_Data(0 To DataCount)
                Delivery_List_Data(DataCount).C_ID = DeliveryGroupListData("C_ID")
                Delivery_List_Data(DataCount).C_NAME = DeliveryGroupListData("C_NAME")
                DataCount += 1
            Loop

            '値を取得できたのでClose
            DeliveryGroupListData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 納品先の明細情報を取得する（納品先別出荷リスト向け）
    ' <引数>
    ' C_ID : 納品先ID
    ' FILE_NAME : 出荷指示ファイル名
    ' <戻り値>
    ' Delivery_List_Data : 納品先情報を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDeliveryPrtList(ByVal C_ID As Integer, _
                                       ByVal FILE_NAME As String, _
                                       ByRef Delivery_List_Data() As Delivery_List, _
                                       ByRef TotalNum As Integer, _
                                       ByRef Result As String, _
                                       ByRef ErrorMessage As String) As Boolean
        Dim DeliveryGroupListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataCount As Integer = 0
        Dim DeliveryDataCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,OUT_TBL.FILE_NAME,M_CUSTOMER.D_NAME,M_ITEM.I_CODE,M_ITEM.I_NAME,"
            Command.CommandText &= " (SELECT SUM(OUT_TBL.NUM) FROM OUT_TBL INNER JOIN M_ITEM AS ITEM ON OUT_TBL.I_ID = ITEM.ID "
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME ='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND ITEM.ID = M_ITEM.ID AND OUT_TBL.C_ID ="
            Command.CommandText &= C_ID
            Command.CommandText &= ") AS NUM "
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
            Command.CommandText &= " WHERE OUT_TBL.C_ID="
            Command.CommandText &= C_ID
            Command.CommandText &= " AND OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            '2012.01.23 修正依頼 order byをI_CODEの昇順に変更
            ' Command.CommandText &= "' GROUP BY OUT_TBL.I_ID ORDER BY OUT_TBL.SHEET_NO,OUT_TBL.ORDER_NO,OUT_TBL.ID;"
            Command.CommandText &= "' GROUP BY OUT_TBL.I_ID ORDER BY M_ITEM.I_CODE;"

            'データ取得
            DeliveryGroupListData = Command.ExecuteReader()

            'DeliveryListData
            Do While (DeliveryGroupListData.Read)
                ReDim Preserve Delivery_List_Data(0 To DeliveryDataCount)
                Delivery_List_Data(DeliveryDataCount).ID = DeliveryGroupListData("ID")
                Delivery_List_Data(DeliveryDataCount).C_ID = DeliveryGroupListData("C_ID")
                Delivery_List_Data(DeliveryDataCount).FILE_NAME = DeliveryGroupListData("FILE_NAME")
                Delivery_List_Data(DeliveryDataCount).D_NAME = DeliveryGroupListData("D_NAME")
                Delivery_List_Data(DeliveryDataCount).I_CODE = DeliveryGroupListData("I_CODE")
                Delivery_List_Data(DeliveryDataCount).I_NAME = DeliveryGroupListData("I_NAME")
                Delivery_List_Data(DeliveryDataCount).NUM = DeliveryGroupListData("NUM")
                TotalNum += DeliveryGroupListData("NUM")
                DeliveryDataCount += 1
            Loop

            '値を取得できたのでClose
            DeliveryGroupListData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 納品書チェックリストの一覧を取得する
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' Sheet_Type : 伝票タイプ
    ' <戻り値>
    ' Check_List_Data : チェックリストデータ格納
    ' Total : 総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDliveryCheckData(ByVal FILE_NAME As String, _
                                        ByVal Sheet_Type As Integer, _
                                        ByRef PageMax As Integer, _
                                        ByRef Check_List_Data() As Check_List, _
                                        ByRef Total As Integer, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim DeliveryGroupListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim i As Integer = 0
        Dim OrderNo_Loop As Integer = 0
        Dim DataCount As Integer = 0
        Dim DeliveryDataCount As Integer = 0
        Dim CID_Check_List() As CheckID_List = Nothing
        Dim OrderNoListData As MySqlDataReader
        Dim PageCheck As Integer = 0
        Dim DataTotalCount As Integer = 0
        Dim TMP_Check_List_Data() As Check_List = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID"
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND M_CUSTOMER.SHEET_TYPE="
            Command.CommandText &= Sheet_Type
            Command.CommandText &= " GROUP BY OUT_TBL.C_ID;"

            'データ取得
            DeliveryGroupListData = Command.ExecuteReader()

            Do While (DeliveryGroupListData.Read)
                ReDim Preserve CID_Check_List(0 To DataCount)
                CID_Check_List(DataCount).C_ID = DeliveryGroupListData("C_ID")
                DataCount += 1
            Loop

            '値を取得できたのでClose
            DeliveryGroupListData.Close()

            If DataCount <> 0 Then

                '納品先のデータをループして、伝票番号の一覧を取得する。
                For Count = 0 To CID_Check_List.Length - 1

                    Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,M_CUSTOMER.C_CODE,OUT_TBL.SHEET_NO,M_CUSTOMER.D_NAME,OUT_TBL.FILE_NAME"
                    Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
                    Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
                    Command.CommandText &= FILE_NAME
                    Command.CommandText &= "' AND OUT_TBL.C_ID="
                    Command.CommandText &= CID_Check_List(Count).C_ID
                    Command.CommandText &= " GROUP BY OUT_TBL.SHEET_NO;"
                    i = 0
                    'データ取得
                    OrderNoListData = Command.ExecuteReader()

                    Do While (OrderNoListData.Read)
                        ReDim Preserve TMP_Check_List_Data(0 To i)
                        TMP_Check_List_Data(i).ID = OrderNoListData("ID")
                        TMP_Check_List_Data(i).C_ID = OrderNoListData("C_ID")
                        TMP_Check_List_Data(i).C_CODE = OrderNoListData("C_CODE")
                        TMP_Check_List_Data(i).SHEET_NO = OrderNoListData("SHEET_NO")
                        TMP_Check_List_Data(i).D_NAME = OrderNoListData("D_NAME")
                        TMP_Check_List_Data(i).FILE_NAME = OrderNoListData("FILE_NAME")

                        i += 1
                    Loop
                    '値を取得できたのでClose
                    OrderNoListData.Close()

                    '伝票番号ごとの商品数量を求め、帳票の枚数を算出する。
                    For OrderNo_Loop = 0 To TMP_Check_List_Data.Length - 1

                        Command.CommandText = "SELECT Count(*) AS COUNT"
                        Command.CommandText &= " FROM OUT_TBL "
                        Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
                        Command.CommandText &= FILE_NAME
                        Command.CommandText &= "' AND OUT_TBL.SHEET_NO='"
                        Command.CommandText &= TMP_Check_List_Data(OrderNo_Loop).SHEET_NO
                        Command.CommandText &= "' AND OUT_TBL.C_ID="
                        Command.CommandText &= TMP_Check_List_Data(OrderNo_Loop).C_ID
                        Command.CommandText &= ";"

                        'データ取得
                        OrderNoListData = Command.ExecuteReader(CommandBehavior.SingleRow)

                        If OrderNoListData.Read Then
                            ' レコードが取得できた時の処理

                            'データが伝票一枚のMAX件数以下ならPageMaxに1を設定
                            If OrderNoListData("COUNT") <= PageMax Then
                                TMP_Check_List_Data(OrderNo_Loop).DATA_NUM = 1
                            Else
                                'PageMax件以上なら、データ件数 \ PageMax件数で商を求め、Modで余りが出るなら+1
                                If OrderNoListData("COUNT") Mod PageMax = 0 Then
                                    TMP_Check_List_Data(OrderNo_Loop).DATA_NUM = OrderNoListData("COUNT") \ PageMax
                                Else
                                    TMP_Check_List_Data(OrderNo_Loop).DATA_NUM = OrderNoListData("COUNT") \ PageMax + 1
                                End If
                            End If
                        Else
                            ' レコードが取得できなかった時の処理
                            ErrorMessage = "納品書チェックリストデータの取得に失敗しました。"
                            Exit Function
                        End If

                        '戻し先用の配列に格納する。
                        ReDim Preserve Check_List_Data(0 To DataTotalCount)
                        Check_List_Data(DataTotalCount).ID = TMP_Check_List_Data(OrderNo_Loop).ID
                        Check_List_Data(DataTotalCount).C_ID = TMP_Check_List_Data(OrderNo_Loop).C_ID
                        Check_List_Data(DataTotalCount).C_CODE = TMP_Check_List_Data(OrderNo_Loop).C_CODE
                        Check_List_Data(DataTotalCount).SHEET_NO = TMP_Check_List_Data(OrderNo_Loop).SHEET_NO
                        Check_List_Data(DataTotalCount).D_NAME = TMP_Check_List_Data(OrderNo_Loop).D_NAME
                        Check_List_Data(DataTotalCount).FILE_NAME = TMP_Check_List_Data(OrderNo_Loop).FILE_NAME
                        Check_List_Data(DataTotalCount).DATA_NUM = TMP_Check_List_Data(OrderNo_Loop).DATA_NUM

                        Total += Check_List_Data(DataTotalCount).DATA_NUM

                        DataTotalCount += 1
                        '値を取得できたのでClose
                        OrderNoListData.Close()
                    Next

                Next
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' コメントリストに必要な情報を取得する。
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' <戻り値>
    ' Comment_List_Data : 出力情報を格納
    ' CommentDataCount : 結果件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetCommentData(ByVal FILE_NAME As String, _
                                   ByVal Place_ID As Integer, _
                                   ByRef Comment_List_Data() As Comment_List, _
                                   ByRef CommentDataCount As Integer, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim CoomentListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim DataCount As Integer = 0
        Dim i As Integer = 0
        Dim Wheresql As String = Nothing

        'If Place = "八潮" Then
        '    Wheresql = " AND COMMENT2 <> ''"
        'ElseIf Place = "東久留米" Then
        '    Wheresql = ""
        'End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.SHEET_NO,M_ITEM.ID AS I_ID,M_ITEM.I_CODE,OUT_TBL.FILE_NAME,"
            Command.CommandText &= " OUT_TBL.NUM,M_CUSTOMER.C_CODE,OUT_TBL.COMMENT1,OUT_TBL.COMMENT2"
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID = M_CUSTOMER.ID"
            Command.CommandText &= " WHERE FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND COMMENT2 IS NULL "
            Command.CommandText &= Wheresql
            Command.CommandText &= " ORDER BY OUT_TBL.SHEET_NO,OUT_TBL.ORDER_NO,OUT_TBL.ID;"

            'データ取得
            CoomentListData = Command.ExecuteReader

            Do While (CoomentListData.Read)
                ReDim Preserve Comment_List_Data(0 To DataCount)
                Comment_List_Data(DataCount).ID = CoomentListData("ID")
                Comment_List_Data(DataCount).I_ID = CoomentListData("I_ID")
                Comment_List_Data(DataCount).I_CODE = CoomentListData("I_CODE")
                Comment_List_Data(DataCount).SHEET_NO = CoomentListData("SHEET_NO")
                Comment_List_Data(DataCount).FILE_NAME = CoomentListData("FILE_NAME")
                Comment_List_Data(DataCount).NUM = CoomentListData("NUM")
                Comment_List_Data(DataCount).C_CODE = CoomentListData("C_CODE")
                Comment_List_Data(DataCount).COMMENT1 = CoomentListData("COMMENT1")
                Comment_List_Data(DataCount).COMMENT2 = CoomentListData("COMMENT2")
                DataCount += 1
            Loop

            '値を取得できたのでClose
            CoomentListData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        CommentDataCount = DataCount
        Return Result
    End Function

    '***********************************************
    ' 在庫履歴検索
    ' <引数>
    ' Item_Code : 商品コード
    ' Date_From : 入出荷日From
    ' Date_To : 入出荷日To
    ' I_Status1 : 良品
    ' I_Status2 : 不良品
    ' Status1 : ステータス（入荷）
    ' Status2 : ステータス（セット組み換え）
    ' Status3 : ステータス（棚卸）
    ' Status4 : ステータス（在庫調整）
    ' Status5 : ステータス（ピッキング戻し）
    ' Status6 : ステータス（ピッキング済み）
    ' Status7 : ステータス（出荷指示キャンセル）
    ' Status8 : ステータス（返品出荷）
    ' Status9 : ステータス（良品区分変更）
    ' Status10 : ステータス（セットばらし）
    ' Status11 : ステータス（出荷）
    ' Comment : コメント
    ' Place : 倉庫
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetStockLogSeach(ByVal Item_Code As String, _
                                     ByVal Date_From As String, _
                                     ByVal Date_To As String, _
                                     ByVal I_Status1 As String, _
                                     ByVal I_Status2 As String, _
                                     ByVal I_Status3 As String, _
                                     ByVal Status1 As String, _
                                     ByVal Status2 As String, _
                                     ByVal Status3 As String, _
                                     ByVal Status4 As String, _
                                     ByVal Status5 As String, _
                                     ByVal Status6 As String, _
                                     ByVal Status7 As String, _
                                     ByVal Status8 As String, _
                                     ByVal Status9 As String, _
                                     ByVal Status10 As String, _
                                     ByVal Status11 As String, _
                                     ByVal Comment As String, _
                                     ByVal Place As Integer, _
                                     ByRef SearchResult() As StockLog_List, _
                                     ByRef Data_Total As Integer, _
                                     ByRef Data_Num_Total As Integer, _
                                     ByRef Result As String, _
                                     ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = ""

            WhereSql = " STOCK_LOG.PLACE_ID="
            WhereSql &= Place
            WhereSql &= " "


            '検索条件よりWhereの作成
            '出荷指示ファイル名
            If Item_Code <> "" Then
                WhereSql = " M_ITEM.I_CODE='"
                WhereSql &= Item_Code
                WhereSql &= "'"
            End If

            '作業日付
            'Fromが入力されている場合
            'If Date_From <> "" Then
            '    If WhereSql = "" Then
            '        WhereSql = " STOCK_LOG.U_DATE >= '"
            '        WhereSql &= Date_From
            '        WhereSql &= "'"
            '    Else
            '        WhereSql &= " AND STOCK_LOG.U_DATE >= '"
            '        WhereSql &= Date_From
            '        WhereSql &= "'"
            '    End If
            'End If
            If Date_From <> "" Then
                If WhereSql = "" Then
                    WhereSql = " date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= Date_From
                    WhereSql &= "'"
                Else
                    WhereSql &= " AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= Date_From
                    WhereSql &= "'"
                End If
            End If


            'Toが入力されている場合
            'If Date_To <> "" Then
            '    If WhereSql = "" Then
            '        WhereSql = " STOCK_LOG.U_DATE <= '"
            '        WhereSql &= Date_To
            '        WhereSql &= "'"
            '    Else
            '        WhereSql &= " AND STOCK_LOG.U_DATE <= '"
            '        WhereSql &= Date_To
            '        WhereSql &= "'"
            '    End If
            'End If
            If Date_To <> "" Then
                If WhereSql = "" Then
                    WhereSql = " date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) <= '"
                    WhereSql &= Date_To
                    WhereSql &= "'"
                Else
                    WhereSql &= " AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) <= '"
                    WhereSql &= Date_To
                    WhereSql &= "'"
                End If
            End If

            'コメント: あいまい検索
            If Comment <> "" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.REMARKS like '%"
                    WhereSql &= Comment
                    WhereSql &= "%'"
                Else
                    WhereSql &= " AND STOCK_LOG.REMARKS like '%"
                    WhereSql &= Comment
                    WhereSql &= "%'"
                End If
            End If

            '不良区分の良品、不良品、保管品のすべてにチェックが入っている
            If I_Status1 = "True" And I_Status2 = "True" And I_Status3 = "True" Then

            ElseIf I_Status1 = "True" And I_Status2 = "False" And I_Status3 = "False" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS = 1 "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS = 1 "
                End If
            ElseIf I_Status1 = "True" And I_Status2 = "True" And I_Status3 = "False" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS in (1,2) "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS in (1,2) "
                End If
            ElseIf I_Status1 = "True" And I_Status2 = "False" And I_Status3 = "True" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS in (1,3) "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS in (1,3) "
                End If
            ElseIf I_Status1 = "False" And I_Status2 = "True" And I_Status3 = "False" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS = 2 "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS = 2 "
                End If
            ElseIf I_Status1 = "False" And I_Status2 = "True" And I_Status3 = "True" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS in (2,3) "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS in (2,3) "
                End If
            ElseIf I_Status1 = "False" And I_Status2 = "False" And I_Status3 = "True" Then
                If WhereSql = "" Then
                    WhereSql = " STOCK_LOG.I_STATUS = 3 "
                Else
                    WhereSql &= " AND STOCK_LOG.I_STATUS = 3 "
                End If
            End If

            'ステータスの出荷予定、ピッキング済み、ピッキング戻し、出荷済みの全てもチェックが入っているか
            '全てチェックなしの場合は条件に入れない。
            If Status1 = "True" And Status2 = "True" And Status3 = "True" And Status4 = "True" And Status5 = "True" _
            And Status6 = "True" And Status7 = "True" And Status8 = "True" And Status9 = "True" And Status10 = "True" And Status11 = "True" Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
            Else
                StatusWhere = ""
                'どれかにチェックが入っていれば、以下の処理を行う。
                If Status1 = True Or Status2 = True Or Status3 = True Or Status4 = True Or Status5 = True _
                Or Status6 = True Or Status7 = True Or Status8 = True Or Status9 = True Or Status10 = True Or Status11 = True Then
                    If Status1 = True Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (1"
                    End If
                    If Status2 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (2"
                    ElseIf Status2 = True And StatusWhere <> "" Then
                        StatusWhere &= ",2"
                    End If
                    If Status3 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (3"
                    ElseIf Status3 = True And StatusWhere <> "" Then
                        StatusWhere &= ",3"
                    End If
                    If Status4 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (4"
                    ElseIf Status4 = True And StatusWhere <> "" Then
                        StatusWhere &= ",4"
                    End If
                    If Status5 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (5"
                    ElseIf Status5 = True And StatusWhere <> "" Then
                        StatusWhere &= ",5"
                    End If
                    If Status6 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (6"
                    ElseIf Status6 = True And StatusWhere <> "" Then
                        StatusWhere &= ",6"
                    End If
                    If Status7 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (7"
                    ElseIf Status7 = True And StatusWhere <> "" Then
                        StatusWhere &= ",7"
                    End If
                    If Status8 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (8"
                    ElseIf Status8 = True And StatusWhere <> "" Then
                        StatusWhere &= ",8"
                    End If
                    If Status9 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (9"
                    ElseIf Status9 = True And StatusWhere <> "" Then
                        StatusWhere &= ",9"
                    End If
                    If Status10 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (10"
                    ElseIf Status10 = True And StatusWhere <> "" Then
                        StatusWhere &= ",10"
                    End If
                    If Status11 = True And StatusWhere = "" Then
                        StatusWhere &= "STOCK_LOG.I_FLG in (11"
                    ElseIf Status11 = True And StatusWhere <> "" Then
                        StatusWhere &= ",11"
                    End If
                    StatusWhere &= ") "
                    If WhereSql = "" Then
                        WhereSql = StatusWhere
                    Else
                        WhereSql &= " AND " & StatusWhere
                    End If
                End If
            End If

            'もし検索条件が設定されていたら、 Whereを追加。
            If WhereSql <> "" Then
                WhereSql = "WHERE " & WhereSql
            End If

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (STOCK_LOG) INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID "
            Command.CommandText &= WhereSql
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(SearchData("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT SUM(NUM) AS TOTAL "
            Command.CommandText &= "FROM (STOCK_LOG) INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID "
            Command.CommandText &= WhereSql

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT STOCK_LOG.I_ID,STOCK_LOG.NUM,STOCK_LOG.I_STATUS,STOCK_LOG.PACKAGE_NO,STOCK_LOG.REMARKS,"
            Command.CommandText &= "STOCK_LOG.U_DATE,M_ITEM.I_CODE,M_ITEM.I_NAME,STOCK_LOG.I_FLG,STOCK_LOG.PLACE_ID,M_PLACE.NAME AS P_NAME "
            Command.CommandText &= "FROM (STOCK_LOG) INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON STOCK_LOG.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= "ORDER BY STOCK_LOG.I_ID;"
            'Select実行
            SearchData = Command.ExecuteReader()

            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)

                'I_ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                '数量
                SearchResult(Count).NUM = SearchData("NUM")
                '不良区分
                SearchResult(Count).I_STATUS = SearchData("I_STATUS")
                'パッケージ№
                If IsDBNull(SearchData("PACKAGE_NO")) Then
                    SearchResult(Count).PACKAGE_NO = 0
                Else
                    SearchResult(Count).PACKAGE_NO = SearchData("PACKAGE_NO")
                End If
                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(Count).COMMENT = ""
                Else
                    SearchResult(Count).COMMENT = SearchData("REMARKS")
                End If
                '作業日時
                SearchResult(Count).I_FLG = SearchData("I_FLG")
                '作業日時
                SearchResult(Count).U_DATE = SearchData("U_DATE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '倉庫
                SearchResult(Count).PLACE = SearchData("P_NAME")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品マスタに登録する。その後、八潮倉庫に良品の在庫0のデータを作成する。
    ' <引数>
    ' Data_List : 登録商品が格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function MIns_Item(ByVal Data_List() As MIns_Item_List, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim SearchData As MySqlDataReader

        Dim SQL_indata As String = Nothing
        Dim Search_Result() As Integer

        Dim Datacount As Integer = 0

        Dim PlaceData() As Place_List = Nothing
        Dim PlaceDatacount As Integer = 0
        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'M_ITEMテーブル INSERT
            For Count = 0 To Data_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO M_ITEM(I_CODE,I_NAME,JAN,PL_CODE,PRICE,LOCATION,PACKAGE_FLG,R_USER,C_ID,PURCHASE_PRICE,IMMUNITY_PRICE,REPAIR_PRICE)VALUES('"
                Command.CommandText &= Data_List(Count).I_CODE
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).I_NAME
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).JAN
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(Count).PL_CODE
                Command.CommandText &= ","
                Command.CommandText &= Data_List(Count).PRICE
                Command.CommandText &= ",'"
                Command.CommandText &= Data_List(Count).LOCATION
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(Count).SET_FLG
                Command.CommandText &= ","
                Command.CommandText &= R_User
                Command.CommandText &= ",0,"
                Command.CommandText &= Data_List(Count).PURCHASE_PRICE
                Command.CommandText &= ","
                Command.CommandText &= Data_List(Count).IMMUNITY_PRICE
                Command.CommandText &= ","
                Command.CommandText &= Data_List(Count).REPAIR_PRICE
                Command.CommandText &= ");"
                'INSERT実行
                Command.ExecuteNonQuery()

                If Count = Data_List.Length - 1 Then
                    SQL_indata &= "'" & Data_List(Count).I_CODE & "'"
                Else
                    SQL_indata &= "'" & Data_List(Count).I_CODE & "',"
                End If
            Next

            '商品IDを取得し、
            'STOCKテーブルに倉庫分、良品、不良品、保管品の在庫数0というデータを登録する。

            '倉庫データ取得
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID FROM M_PLACE WHERE DEL_FLG=0;"
            'Select実行
            SearchData = Command.ExecuteReader()

            '倉庫IDを取得
            Do While (SearchData.Read)
                ReDim Preserve PlaceData(0 To PlaceDatacount)
                'ID
                PlaceData(PlaceDatacount).ID = SearchData("ID")
                PlaceDatacount += 1
            Loop
            'Close
            SearchData.Close()

            Command = Connection.CreateCommand
            Command.CommandText &= "SELECT ID FROM M_ITEM WHERE I_CODE in("
            Command.CommandText &= SQL_indata
            Command.CommandText &= ") ORDER BY ID;"
            'Select実行
            SearchData = Command.ExecuteReader()

            'さきほど登録したデータの商品IDを取得
            Do While (SearchData.Read)
                ReDim Preserve Search_Result(0 To Datacount)
                'ID
                Search_Result(Datacount) = SearchData("ID")
                Datacount += 1
            Loop
            'Close
            SearchData.Close()

            '在庫テーブルにデータ登録
            For Count = 0 To Search_Result.Length - 1
                For P_Count = 0 To PlaceData.Length - 1

                    '良品データ登録
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,U_DATE,PLACE_ID)VALUES("
                    Command.CommandText &= Search_Result(Count)
                    Command.CommandText &= ",0,'良品','',Current_Timestamp,"
                    Command.CommandText &= PlaceData(P_Count).ID
                    Command.CommandText &= ");"
                    'INSERT実行
                    Command.ExecuteNonQuery()

                    '不良品データ登録
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,U_DATE,PLACE_ID)VALUES("
                    Command.CommandText &= Search_Result(Count)
                    Command.CommandText &= ",0,'不良品','',Current_Timestamp,"
                    Command.CommandText &= PlaceData(P_Count).ID
                    Command.CommandText &= ");"
                    'INSERT実行
                    Command.ExecuteNonQuery()

                    '保管品データ登録
                    Command = Connection.CreateCommand
                    Command.CommandText = "INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,U_DATE,PLACE_ID)VALUES("
                    Command.CommandText &= Search_Result(Count)
                    Command.CommandText &= ",0,'保管品','',Current_Timestamp,"
                    Command.CommandText &= PlaceData(P_Count).ID
                    Command.CommandText &= ");"
                    'INSERT実行
                    Command.ExecuteNonQuery()
                Next
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()

        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業マスタに登録する
    ' <引数>
    ' Data_List : 企業情報が格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function MIns_Customer(ByVal Data_List() As MIns_Customer_List, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'DETAILテーブル UPDATE
            For Count = 0 To Data_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO M_CUSTOMER(C_CODE,C_NAME,SHEET_TYPE,D_NAME,D_ZIP,D_ADDRESS,D_TEL,D_FAX,R_USER,CUSTOMER_TYPE,CLAIM_CODE,DISCOUNT_RATE,DELIVERY_PRT_FLG)VALUES('"
                Command.CommandText &= Data_List(Count).C_CODE
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).C_NAME
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(Count).SHEET_TYPE
                Command.CommandText &= ",'"
                Command.CommandText &= Data_List(Count).D_NAME
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).D_ZIP
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).D_ADDRESS
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).D_TEL
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(Count).D_FAX
                Command.CommandText &= "',"
                Command.CommandText &= R_User
                Command.CommandText &= ","
                Command.CommandText &= Data_List(Count).CUSTOMER_TYPE
                Command.CommandText &= ",'"
                Command.CommandText &= Data_List(Count).CLAIM_CODE
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(Count).DISCOUNT_RATE
                Command.CommandText &= ","
                Command.CommandText &= Data_List(Count).DELIVERY_OUTPUT_FLG
                Command.CommandText &= ");"
                'INSERT実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()

        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' セット商品をばらして通常商品に戻す。
    ' セット商品の数量をひき、通常商品のデータを追加・更新
    ' STOCK_LOGに履歴を書き込む。
    ' <引数>
    ' List_Data : ばらす商品のデータ格納配列
    ' I_ID : セット商品ID
    ' NUM : 現在の数量
    ' UPD_NUM : ばらす数量
    ' STOCK_ID ： セット商品の在庫ID
    ' PLACE :　セット商品の倉庫情報
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_Dismantling(ByVal List_Data() As Dismantling_List, _
                                    ByVal I_ID As Integer, _
                                    ByVal I_STATUS As Integer, _
                                    ByVal NUM As Integer, _
                                    ByVal UPD_NUM As Integer, _
                                    ByVal STOCK_ID As Integer, _
                                    ByVal PLACE As String, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim Item_Data As MySqlDataReader
        Dim I_ID_Count As Integer = 0
        Dim TotalNum As Integer = 0
        Dim Stock_Data As MySqlDataReader
        Dim UPD_Stock_Num As Integer = 0
        Dim UPD_Stock_Id As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'セット商品の数量をSTOCKテーブルからひく
            Command = Connection.CreateCommand
            Command.CommandText = " UPDATE STOCK SET NUM="
            Command.CommandText &= (NUM - UPD_NUM)
            Command.CommandText &= ", U_DATE=Current_Timestamp WHERE ID="
            Command.CommandText &= STOCK_ID
            Command.CommandText &= ";"
            'UPDATE実行
            Command.ExecuteNonQuery()

            'STOCK_LOGにInsert
            Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,PLACE_ID,I_FLG,PACKAGE_NO,U_USER,U_DATE)VALUES("
            Command.CommandText &= I_ID
            Command.CommandText &= ","
            Command.CommandText &= (UPD_NUM * -1)
            Command.CommandText &= ",'"
            Command.CommandText &= I_STATUS
            Command.CommandText &= "','"
            Command.CommandText &= PLACE
            Command.CommandText &= "',10,NULL,"
            Command.CommandText &= R_User
            Command.CommandText &= ",Current_Timestamp);"

            'Insert実行
            Command.ExecuteNonQuery()

            For Count = 0 To List_Data.Length - 1
                '通常商品の情報を取得する。
                Command.CommandText = " SELECT COUNT(*) AS COUNT FROM STOCK WHERE I_ID="
                Command.CommandText &= List_Data(Count).I_ID
                Command.CommandText &= " AND I_STATUS='"
                Command.CommandText &= List_Data(Count).I_STATUS
                Command.CommandText &= "' AND PLACE_ID='"
                Command.CommandText &= List_Data(Count).PLACE_ID
                Command.CommandText &= "';"
                'Select実行
                Item_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Item_Data.Read Then
                    ' レコードが取得できた時の処理
                    I_ID_Count = Item_Data("COUNT")
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                    Item_Data.Close()
                    Exit Function
                End If
                'ID値を取得できたのでClose
                Item_Data.Close()

                '数量があればUPDATE、0ならばInsertをする。
                If I_ID_Count = 0 Then
                    'STOCKにInsert
                    Command.CommandText = " INSERT INTO STOCK(I_ID,NUM,I_STATUS,LOCATION,PLACE_ID,U_DATE)VALUES("
                    Command.CommandText &= List_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= (List_Data(Count).NUM * UPD_NUM)
                    Command.CommandText &= ",'"
                    Command.CommandText &= List_Data(Count).I_STATUS
                    Command.CommandText &= "','"
                    Command.CommandText &= List_Data(Count).LOCATION
                    Command.CommandText &= "','"
                    Command.CommandText &= List_Data(Count).PLACE_ID
                    Command.CommandText &= "',Current_Timestamp);"

                    'Insert実行
                    Command.ExecuteNonQuery()

                    'STOCK_LOGにInsert
                    Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,PLACE_ID,U_USER,U_DATE)VALUES("
                    Command.CommandText &= List_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= (List_Data(Count).NUM * UPD_NUM)
                    Command.CommandText &= ",'"
                    Command.CommandText &= List_Data(Count).I_STATUS
                    Command.CommandText &= "',10,"
                    Command.CommandText &= I_ID
                    Command.CommandText &= ",'"
                    Command.CommandText &= List_Data(Count).PLACE_ID
                    Command.CommandText &= "',"
                    Command.CommandText &= R_User
                    Command.CommandText &= ",Current_Timestamp);"

                    'Insert実行
                    Command.ExecuteNonQuery()
                Else
                    'Update

                    '在庫の数量を求める。
                    Command.CommandText = " SELECT ID,NUM FROM STOCK WHERE I_ID="
                    Command.CommandText &= List_Data(Count).I_ID
                    Command.CommandText &= " AND I_STATUS='"
                    Command.CommandText &= List_Data(Count).I_STATUS
                    Command.CommandText &= "' AND PLACE_ID='"
                    Command.CommandText &= List_Data(Count).PLACE_ID
                    Command.CommandText &= "';"
                    'Select実行
                    Stock_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If Stock_Data.Read Then
                        ' レコードが取得できた時の処理
                        UPD_Stock_Num = Stock_Data("NUM")
                        UPD_Stock_Id = Stock_Data("ID")

                    Else
                        ' レコードが取得できなかった時の処理
                        ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                        Stock_Data.Close()
                        Exit Function
                    End If
                    'ID値を取得できたのでClose
                    Stock_Data.Close()

                    TotalNum = UPD_Stock_Num + (List_Data(Count).NUM * UPD_NUM)

                    'STOCKテーブルUPDATE用コマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " UPDATE STOCK SET NUM="
                    Command.CommandText &= TotalNum
                    Command.CommandText &= ", U_DATE=Current_Timestamp WHERE ID="
                    Command.CommandText &= UPD_Stock_Id
                    Command.CommandText &= ";"

                    'STOCKテーブルへデータ更新
                    Command.ExecuteNonQuery()

                    'STOCK_LOGテーブルINSERT用コマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " INSERT INTO STOCK_LOG(I_ID,NUM,I_STATUS,I_FLG,PACKAGE_NO,PLACE_ID,U_USER,U_DATE)VALUES("
                    Command.CommandText &= List_Data(Count).I_ID
                    Command.CommandText &= ","
                    Command.CommandText &= (List_Data(Count).NUM * UPD_NUM)
                    Command.CommandText &= ",'"
                    Command.CommandText &= List_Data(Count).I_STATUS
                    Command.CommandText &= "',10,"
                    Command.CommandText &= I_ID
                    Command.CommandText &= ",'"
                    Command.CommandText &= List_Data(Count).PLACE_ID
                    Command.CommandText &= "',"
                    Command.CommandText &= R_User
                    Command.CommandText &= ",Current_Timestamp);"

                    'STOCK_LOGテーブルへデータ登録
                    Command.ExecuteNonQuery()
                End If
            Next

            '全てのデータをSTOCK,STOCK_LOGテーブルにUPDATE,INSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 複数の商品IDで商品名、在庫数量を取得する。
    ' <引数>
    ' WhereSql : DataGridViewでチェックされた商品IDを格納(SQL加工済み）
    ' <戻り値>
    ' RackCard_List : 結果格納用配列 
    ' DataCount : 結果件数を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetRackCardList(ByRef RackCard_List() As RackCard_List, _
                                    ByVal WhereSql As String, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean

        Dim RackCardListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Dim DataCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT STOCK.ID,STOCK.I_ID ,M_ITEM.I_CODE,M_ITEM.I_NAME,STOCK.NUM AS STOCK_NUM FROM STOCK "
            Command.CommandText &= " INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID "
            Command.CommandText &= " WHERE STOCK.I_ID in ("
            Command.CommandText &= WhereSql
            Command.CommandText &= ") ORDER BY STOCK.I_ID"

            'データ取得
            RackCardListData = Command.ExecuteReader

            Do While (RackCardListData.Read)
                ReDim Preserve RackCard_List(0 To DataCount)
                RackCard_List(DataCount).ID = RackCardListData("ID")
                RackCard_List(DataCount).I_ID = RackCardListData("I_ID")
                RackCard_List(DataCount).I_CODE = RackCardListData("I_CODE")
                RackCard_List(DataCount).I_NAME = RackCardListData("I_NAME")
                RackCard_List(DataCount).STOCK_NUM = RackCardListData("STOCK_NUM")
                DataCount += 1
            Loop

            '0件ならエラー
            If DataCount = 0 Then
                ErrorMessage = "在庫データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 棚卸票出力
    ' 
    ' <引数>
    ' 
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Tanaoroshi_Prt(ByRef SQL As String, _
                                   ByRef Tanaoroshi_PrtData() As Tanaoroshi_PrtList, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean

        Dim TanaoroshiList As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT STOCK.LOCATION,M_PLINE.NAME AS PL_NAME, M_ITEM.I_CODE,M_ITEM.I_NAME,STOCK.NUM"
            Command.CommandText &= " FROM STOCK INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
            Command.CommandText &= " WHERE STOCK.ID IN("
            Command.CommandText &= SQL
            Command.CommandText &= ") ORDER BY STOCK.LOCATION,M_ITEM.I_CODE;"

            'データ取得
            TanaoroshiList = Command.ExecuteReader

            Do While (TanaoroshiList.Read)
                ReDim Preserve Tanaoroshi_PrtData(0 To Count)
                Tanaoroshi_PrtData(Count).LOCATION = TanaoroshiList("LOCATION")
                Tanaoroshi_PrtData(Count).PL_NAME = TanaoroshiList("PL_NAME")
                Tanaoroshi_PrtData(Count).I_CODE = TanaoroshiList("I_CODE")
                Tanaoroshi_PrtData(Count).I_NAME = TanaoroshiList("I_NAME")
                Tanaoroshi_PrtData(Count).NUM = TanaoroshiList("NUM")
                Count += 1
            Loop

            'Command.CommandText &= " ORDER BY STOCK.I_ID"

            '件数分ループ
            'For Count = 0 To Stock_Data.Length - 1

            '    'オープン
            '    Connection.Open()
            '    'コマンド作成

            '    Command = Connection.CreateCommand
            '    Command.CommandText = "SELECT STOCK.LOCATION,M_PLINE.NAME AS PL_NAME, M_ITEM.I_CODE,M_ITEM.I_NAME,STOCK.NUM"
            '    Command.CommandText &= " FROM STOCK INNER JOIN M_ITEM ON STOCK.I_ID=M_ITEM.ID"
            '    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
            '    Command.CommandText &= " WHERE STOCK.ID="
            '    Command.CommandText &= Stock_Data(Count).STOCK_ID
            '    Command.CommandText &= ";"


            '    'Select実行
            '    TanaoroshiList = Command.ExecuteReader(CommandBehavior.SingleRow)

            '    If TanaoroshiList.Read Then
            '        ' レコードが取得できた時の処理
            '        ReDim Preserve Tanaoroshi_PrtData(0 To Count)
            '        Tanaoroshi_PrtData(Count).LOCATION = TanaoroshiList("LOCATION")
            '        Tanaoroshi_PrtData(Count).PL_NAME = TanaoroshiList("PL_NAME")
            '        Tanaoroshi_PrtData(Count).I_CODE = TanaoroshiList("I_CODE")
            '        Tanaoroshi_PrtData(Count).I_NAME = TanaoroshiList("I_NAME")
            '        Tanaoroshi_PrtData(Count).NUM = TanaoroshiList("NUM")

            '        Count += 1
            '    End If
            '    'ID値を取得できたのでClose
            '    TanaoroshiList.Close()
            'Next



            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "在庫データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 在庫商品を取得する（１件）
    ' <引数>
    ' I_Code : 商品コード
    ' Place : 倉庫（1:八潮、2:東久留米）
    ' <戻り値>
    ' Stock_Item : 結果格納配列
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetStockItem(ByVal I_Code As String, _
                                 ByVal Place As Integer, _
                                 ByRef Stock_Item() As Stock_Item_List, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean
        Dim StockData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT STOCK.ID AS STOCK_ID,STOCK.I_ID,M_ITEM.I_NAME,M_ITEM.JAN,M_PLINE.NAME AS PL_NAME,STOCK.LOCATION,"
            Command.CommandText &= " STOCK.I_STATUS,STOCK.NUM"
            Command.CommandText &= " FROM STOCK INNER JOIN M_ITEM ON STOCK.I_ID = M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID"
            Command.CommandText &= " WHERE M_ITEM.I_CODE='"
            Command.CommandText &= I_Code
            Command.CommandText &= "' AND STOCK.PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= " AND STOCK.I_STATUS=1 AND M_ITEM.DEL_FLG=0;"

            'データ取得
            StockData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If StockData.Read Then
                ' レコードが取得できた時の処理
                ReDim Preserve Stock_Item(0 To 0)
                Stock_Item(0).STOCK_ID = StockData("STOCK_ID")
                Stock_Item(0).I_ID = StockData("I_ID")
                Stock_Item(0).I_NAME = StockData("I_NAME")
                Stock_Item(0).JAN = StockData("JAN")
                Stock_Item(0).PL_NAME = StockData("PL_NAME")
                Stock_Item(0).LOCATION = StockData("LOCATION")
                Stock_Item(0).I_STATUS = StockData("I_STATUS")
                Stock_Item(0).NUM = StockData("NUM")

            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した商品コードに該当する在庫データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷伝票検索
    ' OUT_PRTテーブルに登録されているデータを取得する。
    ' <引数>
    ' FileName : 出荷指示ファイル名
    ' PLACE : 倉庫（1:八潮、2:東久留米）
    ' <戻り値>
    ' SlipList : 結果格納配列
    ' Data_Total : 商品数
    ' Data_Num_Total : 数量の総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetSlipSearch(ByVal FileName As String, _
                                  ByVal Place As Integer, _
                                  ByRef SlipLitData() As Slip_List, _
                                  ByRef Data_Total As Integer, _
                                  ByRef Data_Num_Total As Integer, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean
        Dim DataTotalSQL As MySqlDataReader
        Dim DataNumTotalSQL As MySqlDataReader
        Dim SlipList As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM OUT_PRT INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
            Command.CommandText &= FileName
            Command.CommandText &= "' AND OUT_PRT.PLACE_ID="
            Command.CommandText &= Place


            'Select実行
            DataTotalSQL = Command.ExecuteReader(CommandBehavior.SingleRow)

            If DataTotalSQL.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(DataTotalSQL("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            DataTotalSQL.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select SUM(NUM) AS TOTAL "
            Command.CommandText &= " FROM OUT_PRT INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
            Command.CommandText &= FileName
            Command.CommandText &= "' AND OUT_PRT.PLACE_ID="
            Command.CommandText &= Place

            'Select実行
            DataNumTotalSQL = Command.ExecuteReader(CommandBehavior.SingleRow)

            If DataNumTotalSQL.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(DataNumTotalSQL("TOTAL")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = DataNumTotalSQL("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'TOTAL値を取得できたのでClose
            DataNumTotalSQL.Close()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT OUT_PRT.ID,OUT_PRT.FILE_NAME,OUT_PRT.SHEET_NO,OUT_PRT.ORDER_NO,M_ITEM.ID AS I_ID,OUT_PRT.I_CODE,M_ITEM.I_NAME,"
            Command.CommandText &= " M_CUSTOMER.ID AS C_ID,OUT_PRT.C_CODE,M_CUSTOMER.C_NAME,OUT_PRT.UNIT_COST,OUT_PRT.NUM, "
            Command.CommandText &= " OUT_PRT.COMMENT1,OUT_PRT.COMMENT2,OUT_PRT.O_DATE,OUT_PRT.PRT_STATUS "
            Command.CommandText &= " FROM OUT_PRT INNER JOIN M_ITEM ON OUT_PRT.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_CUSTOMER ON OUT_PRT.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_PRT.R_STATUS=2 AND OUT_PRT.FILE_NAME='"
            Command.CommandText &= FileName
            Command.CommandText &= "' AND OUT_PRT.PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= " ORDER BY OUT_PRT.ID"

            'データ取得
            SlipList = Command.ExecuteReader

            Do While (SlipList.Read)
                ReDim Preserve SlipLitData(0 To Count)
                SlipLitData(Count).ID = SlipList("ID")
                SlipLitData(Count).FILE_NAME = SlipList("FILE_NAME")
                SlipLitData(Count).SHEET_NO = SlipList("SHEET_NO")
                SlipLitData(Count).ORDER_NO = SlipList("ORDER_NO")
                SlipLitData(Count).I_ID = SlipList("I_ID")
                SlipLitData(Count).I_CODE = SlipList("I_CODE")
                SlipLitData(Count).I_NAME = SlipList("I_NAME")
                SlipLitData(Count).C_ID = SlipList("C_ID")
                SlipLitData(Count).C_CODE = SlipList("C_CODE")
                SlipLitData(Count).C_NAME = SlipList("C_NAME")
                SlipLitData(Count).UNIT_COST = SlipList("UNIT_COST")
                SlipLitData(Count).NUM = SlipList("NUM")
                SlipLitData(Count).COMMENT1 = SlipList("COMMENT1")
                SlipLitData(Count).COMMENT2 = SlipList("COMMENT2")
                SlipLitData(Count).O_DATE = SlipList("O_DATE")
                SlipLitData(Count).PRT_STATUS = SlipList("PRT_STATUS")
                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 在庫履歴画面　サマリーデータファイル作成
    ' 作業日付を元に日別、プロダクトラインコード別に入庫、出庫数を出力
    ' <引数>
    ' Date_From : 作業日付From
    ' Date_To : 作業日付To
    ' <戻り値>
    ' Summary_List : 結果格納配列
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetHistorySummary(ByVal Date_From As Date, _
                                  ByRef Date_To As Date, _
                                  ByRef Place As Integer, _
                                  ByRef Summary_INList() As Summary_List, _
                                  ByRef Summary_OUTList() As Summary_List, _
                                  ByRef Summary_TanaoroshiList() As Summary_List, _
                                  ByRef PL_List() As PL_List, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean
        Dim IN_Data As MySqlDataReader
        Dim OUT_Data As MySqlDataReader
        Dim TANAOROSHI_Data As MySqlDataReader
        Dim M_PLine As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim PLCount As Integer = 0

        Dim MPLine As String = Nothing
        Dim SQL As String = Nothing

        Dim Search_Date As Date = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'プロダクトラインの一覧を取得
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT ID,NAME FROM M_PLINE WHERE DEL_FLG=0 AND ID <> 77 ORDER BY ID"

            'Select実行
            M_PLine = Command.ExecuteReader

            Do While (M_PLine.Read)
                ReDim Preserve PL_List(0 To PLCount)
                PL_List(PLCount).ID = M_PLine("ID")
                PL_List(PLCount).NAME = M_PLine("NAME")
                PLCount += 1
            Loop
            '取得できたのでClose
            M_PLine.Close()

            Search_Date = Date_From
            '取得したプロダクトラインを元にSQLを作成。

            'FromとToが同じなら
            If Date_From = Date_To Then
                SQL = ""
                Search_Date = Date_From

                'プロダクトラインごとの数量をSUMするようSQLを作成。
                'LINE用
                SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS LINE_NUM,"
                'Bait用
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bait_NUM,"
                'Rods
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Rods_NUM,"
                'Baitcast_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS BReels_NUM,"
                'Spinning_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS SReels_NUM,"
                'Combos
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Combos_NUM,"
                'Accessory
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Accessory_NUM,"
                'Accessory2
                'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                'SQL &= Search_Date
                'SQL &= "') AS Accessory2_NUM,"
                'Hard_Lure
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Hard_Lure_NUM,"
                'BagApparel
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bag_NUM,"
                'RodOEM
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodOEM_NUM,"
                'RodParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodParts_NUM,"
                'ReelParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS ReelParts_NUM"

                'データを取得するコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT "
                Command.CommandText &= SQL
                Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                Command.CommandText &= " WHERE STOCK_LOG.I_FLG=1 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                Command.CommandText &= Search_Date
                Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                'Select実行
                IN_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If IN_Data.Read Then
                    ' レコードが取得できた時の処理
                    ReDim Preserve Summary_INList(0 To Count)
                    Summary_INList(Count).WorkDate = Search_Date
                    If IsDBNull(IN_Data("Line_NUM")) Then
                        Summary_INList(Count).Line = 0
                    Else
                        Summary_INList(Count).Line = IN_Data("Line_NUM")
                    End If

                    If IsDBNull(IN_Data("Bait_NUM")) Then
                        Summary_INList(Count).Bait = 0
                    Else
                        Summary_INList(Count).Bait = IN_Data("Bait_NUM")
                    End If

                    If IsDBNull(IN_Data("Rods_NUM")) Then
                        Summary_INList(Count).Rods = 0
                    Else
                        Summary_INList(Count).Rods = IN_Data("Rods_NUM")
                    End If

                    If IsDBNull(IN_Data("BReels_NUM")) Then
                        Summary_INList(Count).B_Reels = 0
                    Else
                        Summary_INList(Count).B_Reels = IN_Data("BReels_NUM")
                    End If

                    If IsDBNull(IN_Data("SReels_NUM")) Then
                        Summary_INList(Count).S_Reels = 0
                    Else
                        Summary_INList(Count).S_Reels = IN_Data("SReels_NUM")
                    End If

                    If IsDBNull(IN_Data("Combos_NUM")) Then
                        Summary_INList(Count).Combos = 0
                    Else
                        Summary_INList(Count).Combos = IN_Data("Combos_NUM")
                    End If

                    If IsDBNull(IN_Data("Accessory_NUM")) Then
                        Summary_INList(Count).Accessory = 0
                    Else
                        Summary_INList(Count).Accessory = IN_Data("Accessory_NUM")
                    End If
                    'If IsDBNull(IN_Data("Accessory2_NUM")) Then
                    '    Summary_INList(Count).Accessory2 = 0
                    'Else
                    '    Summary_INList(Count).Accessory2 = IN_Data("Accessory2_NUM")
                    'End If

                    If IsDBNull(IN_Data("Hard_Lure_NUM")) Then
                        Summary_INList(Count).Hard_Lure = 0
                    Else
                        Summary_INList(Count).Hard_Lure = IN_Data("Hard_Lure_NUM")
                    End If

                    If IsDBNull(IN_Data("Bag_NUM")) Then
                        Summary_INList(Count).Bag = 0
                    Else
                        Summary_INList(Count).Bag = IN_Data("Bag_NUM")
                    End If

                    If IsDBNull(IN_Data("RodOEM_NUM")) Then
                        Summary_INList(Count).Rod_OEM = 0
                    Else
                        Summary_INList(Count).Rod_OEM = IN_Data("RodOEM_NUM")
                    End If

                    If IsDBNull(IN_Data("RodParts_NUM")) Then
                        Summary_INList(Count).Rod_Parts = 0
                    Else
                        Summary_INList(Count).Rod_Parts = IN_Data("RodParts_NUM")
                    End If
                    If IsDBNull(IN_Data("ReelParts_NUM")) Then
                        Summary_INList(Count).ReelParts = 0
                    Else
                        Summary_INList(Count).ReelParts = IN_Data("ReelParts_NUM")
                    End If

                Else
                    'データがなかった場合は全てに0を入れる。
                    ReDim Preserve Summary_INList(0 To Count)
                    Summary_INList(Count).WorkDate = Search_Date
                    Summary_INList(Count).Line = 0
                    Summary_INList(Count).Bait = 0
                    Summary_INList(Count).Rods = 0
                    Summary_INList(Count).B_Reels = 0
                    Summary_INList(Count).S_Reels = 0
                    Summary_INList(Count).Combos = 0
                    Summary_INList(Count).Accessory = 0
                    'Summary_INList(Count).Accessory2 = 0
                    Summary_INList(Count).Hard_Lure = 0
                    Summary_INList(Count).Bag = 0
                    Summary_INList(Count).Rod_OEM = 0
                    Summary_INList(Count).Rod_Parts = 0
                    Summary_INList(Count).ReelParts = 0

                End If

                '入庫情報を取得できたのでClose
                IN_Data.Close()


                'プロダクトラインごとの数量をSUMするようSQLを作成。
                'LINE用
                SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS LINE_NUM,"
                'Bait用
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bait_NUM,"
                'Rods
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Rods_NUM,"
                'Baitcast_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS BReels_NUM,"
                'Spinning_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS SReels_NUM,"
                'Combos
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Combos_NUM,"
                'Accessory
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Accessory_NUM,"
                'Accessory2
                'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                'SQL &= Search_Date
                'SQL &= "') AS Accessory2_NUM,"
                'Hard_Lure
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Hard_Lure_NUM,"
                'BagApparel
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bag_NUM,"
                'RodOEM
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodOEM_NUM,"
                'RodParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodParts_NUM,"
                'ReelParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS ReelParts_NUM"

                'データを取得するコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT "
                Command.CommandText &= SQL
                Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                Command.CommandText &= " WHERE STOCK_LOG.I_FLG=11 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                Command.CommandText &= Search_Date
                Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                'Select実行
                OUT_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If OUT_Data.Read Then
                    ' レコードが取得できた時の処理
                    ReDim Preserve Summary_OUTList(0 To Count)
                    Summary_OUTList(Count).WorkDate = Search_Date
                    If IsDBNull(OUT_Data("Line_NUM")) Then
                        Summary_OUTList(Count).Line = 0
                    Else
                        Summary_OUTList(Count).Line = OUT_Data("Line_NUM")
                    End If

                    If IsDBNull(OUT_Data("Bait_NUM")) Then
                        Summary_OUTList(Count).Bait = 0
                    Else
                        Summary_OUTList(Count).Bait = OUT_Data("Bait_NUM")
                    End If

                    If IsDBNull(OUT_Data("Rods_NUM")) Then
                        Summary_OUTList(Count).Rods = 0
                    Else
                        Summary_OUTList(Count).Rods = OUT_Data("Rods_NUM")
                    End If

                    If IsDBNull(OUT_Data("BReels_NUM")) Then
                        Summary_OUTList(Count).B_Reels = 0
                    Else
                        Summary_OUTList(Count).B_Reels = OUT_Data("BReels_NUM")
                    End If

                    If IsDBNull(OUT_Data("SReels_NUM")) Then
                        Summary_OUTList(Count).S_Reels = 0
                    Else
                        Summary_OUTList(Count).S_Reels = OUT_Data("SReels_NUM")
                    End If

                    If IsDBNull(OUT_Data("Combos_NUM")) Then
                        Summary_OUTList(Count).Combos = 0
                    Else
                        Summary_OUTList(Count).Combos = OUT_Data("Combos_NUM")
                    End If

                    If IsDBNull(OUT_Data("Accessory_NUM")) Then
                        Summary_OUTList(Count).Accessory = 0
                    Else
                        Summary_OUTList(Count).Accessory = OUT_Data("Accessory_NUM")
                    End If

                    'If IsDBNull(OUT_Data("Accessory2_NUM")) Then
                    '    Summary_OUTList(Count).Accessory2 = 0
                    'Else
                    '    Summary_OUTList(Count).Accessory2 = OUT_Data("Accessory2_NUM")
                    'End If

                    If IsDBNull(OUT_Data("Hard_Lure_NUM")) Then
                        Summary_OUTList(Count).Hard_Lure = 0
                    Else
                        Summary_OUTList(Count).Hard_Lure = OUT_Data("Hard_Lure_NUM")
                    End If

                    If IsDBNull(OUT_Data("Bag_NUM")) Then
                        Summary_OUTList(Count).Bag = 0
                    Else
                        Summary_OUTList(Count).Bag = OUT_Data("Bag_NUM")
                    End If

                    If IsDBNull(OUT_Data("RodOEM_NUM")) Then
                        Summary_OUTList(Count).Rod_OEM = 0
                    Else
                        Summary_OUTList(Count).Rod_OEM = OUT_Data("RodOEM_NUM")
                    End If

                    If IsDBNull(OUT_Data("RodParts_NUM")) Then
                        Summary_OUTList(Count).Rod_Parts = 0
                    Else
                        Summary_OUTList(Count).Rod_Parts = OUT_Data("RodParts_NUM")
                    End If
                    If IsDBNull(OUT_Data("ReelParts_NUM")) Then
                        Summary_OUTList(Count).ReelParts = 0
                    Else
                        Summary_OUTList(Count).ReelParts = OUT_Data("ReelParts_NUM")
                    End If

                Else
                    'データがなかった場合は全てに0を入れる。
                    ReDim Preserve Summary_OUTList(0 To Count)
                    Summary_OUTList(Count).WorkDate = Search_Date
                    Summary_OUTList(Count).Line = 0
                    Summary_OUTList(Count).Bait = 0
                    Summary_OUTList(Count).Rods = 0
                    Summary_OUTList(Count).B_Reels = 0
                    Summary_OUTList(Count).S_Reels = 0
                    Summary_OUTList(Count).Combos = 0
                    Summary_OUTList(Count).Accessory = 0
                    'Summary_OUTList(Count).Accessory2 = 0
                    Summary_OUTList(Count).Hard_Lure = 0
                    Summary_OUTList(Count).Bag = 0
                    Summary_OUTList(Count).Rod_OEM = 0
                    Summary_OUTList(Count).Rod_Parts = 0
                    Summary_OUTList(Count).ReelParts = 0
                End If


                '出庫情報を取得できたのでClose
                OUT_Data.Close()

                SQL = ""
                '棚卸情報取得


                'プロダクトラインごとの数量をSUMするようSQLを作成。
                'LINE用
                SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS LINE_NUM,"
                'Bait用
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bait_NUM,"
                'Rods
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Rods_NUM,"
                'Baitcast_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS BReels_NUM,"
                'Spinning_Reels
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS SReels_NUM,"
                'Combos
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Combos_NUM,"
                'Accessory
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Accessory_NUM,"
                'Accessory
                'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                'SQL &= Search_Date
                'SQL &= "') AS Accessory2_NUM,"
                'Hard_Lure
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Hard_Lure_NUM,"
                'BagApparel
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS Bag_NUM,"
                'RodOEM
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodOEM_NUM,"
                'RodParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS RodParts_NUM,"
                'ReelParts
                SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                SQL &= Search_Date
                SQL &= "' AND STOCK_LOG.PLACE_ID ="
                SQL &= Place
                SQL &= ") AS ReelParts_NUM"


                'データを取得するコマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT "
                Command.CommandText &= SQL
                Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                Command.CommandText &= " WHERE STOCK_LOG.I_FLG=3 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                Command.CommandText &= Search_Date
                Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                'Select実行
                TANAOROSHI_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If TANAOROSHI_Data.Read Then
                    ' レコードが取得できた時の処理
                    ReDim Preserve Summary_TanaoroshiList(0 To Count)
                    Summary_TanaoroshiList(Count).WorkDate = Search_Date
                    If IsDBNull(TANAOROSHI_Data("Line_NUM")) Then
                        Summary_TanaoroshiList(Count).Line = 0
                    Else
                        Summary_TanaoroshiList(Count).Line = TANAOROSHI_Data("Line_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("Bait_NUM")) Then
                        Summary_TanaoroshiList(Count).Bait = 0
                    Else
                        Summary_TanaoroshiList(Count).Bait = TANAOROSHI_Data("Bait_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("Rods_NUM")) Then
                        Summary_TanaoroshiList(Count).Rods = 0
                    Else
                        Summary_TanaoroshiList(Count).Rods = TANAOROSHI_Data("Rods_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("BReels_NUM")) Then
                        Summary_TanaoroshiList(Count).B_Reels = 0
                    Else
                        Summary_TanaoroshiList(Count).B_Reels = TANAOROSHI_Data("BReels_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("SReels_NUM")) Then
                        Summary_TanaoroshiList(Count).S_Reels = 0
                    Else
                        Summary_TanaoroshiList(Count).S_Reels = TANAOROSHI_Data("SReels_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("Combos_NUM")) Then
                        Summary_TanaoroshiList(Count).Combos = 0
                    Else
                        Summary_TanaoroshiList(Count).Combos = TANAOROSHI_Data("Combos_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("Accessory_NUM")) Then
                        Summary_TanaoroshiList(Count).Accessory = 0
                    Else
                        Summary_TanaoroshiList(Count).Accessory = TANAOROSHI_Data("Accessory_NUM")
                    End If
                    'If IsDBNull(TANAOROSHI_Data("Accessory2_NUM")) Then
                    '    Summary_TanaoroshiList(Count).Accessory2 = 0
                    'Else
                    '    Summary_TanaoroshiList(Count).Accessory2 = TANAOROSHI_Data("Accessory2_NUM")
                    'End If

                    If IsDBNull(TANAOROSHI_Data("Hard_Lure_NUM")) Then
                        Summary_TanaoroshiList(Count).Hard_Lure = 0
                    Else
                        Summary_TanaoroshiList(Count).Hard_Lure = TANAOROSHI_Data("Hard_Lure_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("Bag_NUM")) Then
                        Summary_TanaoroshiList(Count).Bag = 0
                    Else
                        Summary_TanaoroshiList(Count).Bag = TANAOROSHI_Data("Bag_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("RodOEM_NUM")) Then
                        Summary_TanaoroshiList(Count).Rod_OEM = 0
                    Else
                        Summary_TanaoroshiList(Count).Rod_OEM = TANAOROSHI_Data("RodOEM_NUM")
                    End If

                    If IsDBNull(TANAOROSHI_Data("RodParts_NUM")) Then
                        Summary_TanaoroshiList(Count).Rod_Parts = 0
                    Else
                        Summary_TanaoroshiList(Count).Rod_Parts = TANAOROSHI_Data("RodParts_NUM")
                    End If
                    If IsDBNull(TANAOROSHI_Data("ReelParts_NUM")) Then
                        Summary_TanaoroshiList(Count).ReelParts = 0
                    Else
                        Summary_TanaoroshiList(Count).ReelParts = TANAOROSHI_Data("ReelParts_NUM")
                    End If

                Else
                    'データがなかった場合は全てに0を入れる。
                    ReDim Preserve Summary_TanaoroshiList(0 To Count)
                    Summary_TanaoroshiList(Count).WorkDate = Search_Date
                    Summary_TanaoroshiList(Count).Line = 0
                    Summary_TanaoroshiList(Count).Bait = 0
                    Summary_TanaoroshiList(Count).Rods = 0
                    Summary_TanaoroshiList(Count).B_Reels = 0
                    Summary_TanaoroshiList(Count).S_Reels = 0
                    Summary_TanaoroshiList(Count).Combos = 0
                    Summary_TanaoroshiList(Count).Accessory = 0
                    'Summary_TanaoroshiList(Count).Accessory2 = 0
                    Summary_TanaoroshiList(Count).Hard_Lure = 0
                    Summary_TanaoroshiList(Count).Bag = 0
                    Summary_TanaoroshiList(Count).Rod_OEM = 0
                    Summary_TanaoroshiList(Count).Rod_Parts = 0
                    Summary_TanaoroshiList(Count).ReelParts = 0

                End If

                '出庫情報を取得できたのでClose
                TANAOROSHI_Data.Close()


            Else
                Do While Date_To.AddDays(1) <> Search_Date

                    '入庫情報の取得
                    SQL = ""

                    'プロダクトラインごとの数量をSUMするようSQLを作成。
                    'LINE用
                    SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                    SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS LINE_NUM,"
                    'Bait用
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                    SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bait_NUM,"
                    'Rods
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                    SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Rods_NUM,"
                    'Baitcast_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                    SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS BReels_NUM,"
                    'Spinning_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                    SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS SReels_NUM,"
                    'Combos
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                    SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Combos_NUM,"
                    'Accessory
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Accessory_NUM,"
                    'Accessory
                    'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    'SQL &= Search_Date
                    'SQL &= "') AS Accessory2_NUM,"
                    'Hard_Lure
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                    SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Hard_Lure_NUM,"
                    'BagApparel
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                    SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bag_NUM,"
                    'RodOEM
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                    SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodOEM_NUM,"
                    'RodParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                    SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodParts_NUM,"
                    'ReelParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                    SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=1 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS ReelParts_NUM"

                    'データを取得するコマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " SELECT "
                    Command.CommandText &= SQL
                    Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                    Command.CommandText &= " WHERE STOCK_LOG.I_FLG=1 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    Command.CommandText &= Search_Date
                    Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                    'Select実行
                    IN_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If IN_Data.Read Then
                        ' レコードが取得できた時の処理
                        ReDim Preserve Summary_INList(0 To Count)
                        Summary_INList(Count).WorkDate = Search_Date
                        If IsDBNull(IN_Data("Line_NUM")) Then
                            Summary_INList(Count).Line = 0
                        Else
                            Summary_INList(Count).Line = IN_Data("Line_NUM")
                        End If

                        If IsDBNull(IN_Data("Bait_NUM")) Then
                            Summary_INList(Count).Bait = 0
                        Else
                            Summary_INList(Count).Bait = IN_Data("Bait_NUM")
                        End If

                        If IsDBNull(IN_Data("Rods_NUM")) Then
                            Summary_INList(Count).Rods = 0
                        Else
                            Summary_INList(Count).Rods = IN_Data("Rods_NUM")
                        End If

                        If IsDBNull(IN_Data("BReels_NUM")) Then
                            Summary_INList(Count).B_Reels = 0
                        Else
                            Summary_INList(Count).B_Reels = IN_Data("BReels_NUM")
                        End If

                        If IsDBNull(IN_Data("SReels_NUM")) Then
                            Summary_INList(Count).S_Reels = 0
                        Else
                            Summary_INList(Count).S_Reels = IN_Data("SReels_NUM")
                        End If

                        If IsDBNull(IN_Data("Combos_NUM")) Then
                            Summary_INList(Count).Combos = 0
                        Else
                            Summary_INList(Count).Combos = IN_Data("Combos_NUM")
                        End If

                        If IsDBNull(IN_Data("Accessory_NUM")) Then
                            Summary_INList(Count).Accessory = 0
                        Else
                            Summary_INList(Count).Accessory = IN_Data("Accessory_NUM")
                        End If

                        'If IsDBNull(IN_Data("Accessory2_NUM")) Then
                        '    Summary_INList(Count).Accessory2 = 0
                        'Else
                        '    Summary_INList(Count).Accessory2 = IN_Data("Accessory2_NUM")
                        'End If

                        If IsDBNull(IN_Data("Hard_Lure_NUM")) Then
                            Summary_INList(Count).Hard_Lure = 0
                        Else
                            Summary_INList(Count).Hard_Lure = IN_Data("Hard_Lure_NUM")
                        End If

                        If IsDBNull(IN_Data("Bag_NUM")) Then
                            Summary_INList(Count).Bag = 0
                        Else
                            Summary_INList(Count).Bag = IN_Data("Bag_NUM")
                        End If

                        If IsDBNull(IN_Data("RodOEM_NUM")) Then
                            Summary_INList(Count).Rod_OEM = 0
                        Else
                            Summary_INList(Count).Rod_OEM = IN_Data("RodOEM_NUM")
                        End If

                        If IsDBNull(IN_Data("RodParts_NUM")) Then
                            Summary_INList(Count).Rod_Parts = 0
                        Else
                            Summary_INList(Count).Rod_Parts = IN_Data("RodParts_NUM")
                        End If
                        If IsDBNull(IN_Data("ReelParts_NUM")) Then
                            Summary_INList(Count).ReelParts = 0
                        Else
                            Summary_INList(Count).ReelParts = IN_Data("ReelParts_NUM")
                        End If

                    Else
                        'データがなかった場合は全てに0を入れる。
                        ReDim Preserve Summary_INList(0 To Count)
                        Summary_INList(Count).WorkDate = Search_Date
                        Summary_INList(Count).Line = 0
                        Summary_INList(Count).Bait = 0
                        Summary_INList(Count).Rods = 0
                        Summary_INList(Count).B_Reels = 0
                        Summary_INList(Count).S_Reels = 0
                        Summary_INList(Count).Combos = 0
                        Summary_INList(Count).Accessory = 0
                        'Summary_INList(Count).Accessory2 = 0
                        Summary_INList(Count).Hard_Lure = 0
                        Summary_INList(Count).Bag = 0
                        Summary_INList(Count).Rod_OEM = 0
                        Summary_INList(Count).Rod_Parts = 0
                        Summary_INList(Count).ReelParts = 0

                    End If

                    '入庫情報を取得できたのでClose
                    IN_Data.Close()

                    '出庫情報の取得
                    SQL = ""


                    'プロダクトラインごとの数量をSUMするようSQLを作成。
                    'LINE用
                    SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                    SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS LINE_NUM,"
                    'Bait用
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                    SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bait_NUM,"
                    'Rods
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                    SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Rods_NUM,"
                    'Baitcast_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                    SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS BReels_NUM,"
                    'Spinning_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                    SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS SReels_NUM,"
                    'Combos
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                    SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Combos_NUM,"
                    'Accessory
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Accessory_NUM,"
                    'Accessory
                    'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    'SQL &= Search_Date
                    'SQL &= "') AS Accessory2_NUM,"
                    'Hard_Lure
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                    SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Hard_Lure_NUM,"
                    'BagApparel
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                    SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bag_NUM,"
                    'RodOEM
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                    SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodOEM_NUM,"
                    'RodParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                    SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodParts_NUM,"
                    'ReelParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                    SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=11 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS ReelParts_NUM"

                    'データを取得するコマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " SELECT "
                    Command.CommandText &= SQL
                    Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                    Command.CommandText &= " WHERE STOCK_LOG.I_FLG=11 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    Command.CommandText &= Search_Date
                    Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                    'Select実行
                    OUT_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If OUT_Data.Read Then
                        ' レコードが取得できた時の処理
                        ReDim Preserve Summary_OUTList(0 To Count)
                        Summary_OUTList(Count).WorkDate = Search_Date
                        If IsDBNull(OUT_Data("Line_NUM")) Then
                            Summary_OUTList(Count).Line = 0
                        Else
                            Summary_OUTList(Count).Line = OUT_Data("Line_NUM")
                        End If

                        If IsDBNull(OUT_Data("Bait_NUM")) Then
                            Summary_OUTList(Count).Bait = 0
                        Else
                            Summary_OUTList(Count).Bait = OUT_Data("Bait_NUM")
                        End If

                        If IsDBNull(OUT_Data("Rods_NUM")) Then
                            Summary_OUTList(Count).Rods = 0
                        Else
                            Summary_OUTList(Count).Rods = OUT_Data("Rods_NUM")
                        End If

                        If IsDBNull(OUT_Data("BReels_NUM")) Then
                            Summary_OUTList(Count).B_Reels = 0
                        Else
                            Summary_OUTList(Count).B_Reels = OUT_Data("BReels_NUM")
                        End If

                        If IsDBNull(OUT_Data("SReels_NUM")) Then
                            Summary_OUTList(Count).S_Reels = 0
                        Else
                            Summary_OUTList(Count).S_Reels = OUT_Data("SReels_NUM")
                        End If

                        If IsDBNull(OUT_Data("Combos_NUM")) Then
                            Summary_OUTList(Count).Combos = 0
                        Else
                            Summary_OUTList(Count).Combos = OUT_Data("Combos_NUM")
                        End If

                        If IsDBNull(OUT_Data("Accessory_NUM")) Then
                            Summary_OUTList(Count).Accessory = 0
                        Else
                            Summary_OUTList(Count).Accessory = OUT_Data("Accessory_NUM")
                        End If

                        'If IsDBNull(OUT_Data("Accessory2_NUM")) Then
                        '    Summary_OUTList(Count).Accessory2 = 0
                        'Else
                        '    Summary_OUTList(Count).Accessory2 = OUT_Data("Accessory2_NUM")
                        'End If

                        If IsDBNull(OUT_Data("Hard_Lure_NUM")) Then
                            Summary_OUTList(Count).Hard_Lure = 0
                        Else
                            Summary_OUTList(Count).Hard_Lure = OUT_Data("Hard_Lure_NUM")
                        End If

                        If IsDBNull(OUT_Data("Bag_NUM")) Then
                            Summary_OUTList(Count).Bag = 0
                        Else
                            Summary_OUTList(Count).Bag = OUT_Data("Bag_NUM")
                        End If

                        If IsDBNull(OUT_Data("RodOEM_NUM")) Then
                            Summary_OUTList(Count).Rod_OEM = 0
                        Else
                            Summary_OUTList(Count).Rod_OEM = OUT_Data("RodOEM_NUM")
                        End If

                        If IsDBNull(OUT_Data("RodParts_NUM")) Then
                            Summary_OUTList(Count).Rod_Parts = 0
                        Else
                            Summary_OUTList(Count).Rod_Parts = OUT_Data("RodParts_NUM")
                        End If
                        If IsDBNull(OUT_Data("ReelParts_NUM")) Then
                            Summary_OUTList(Count).ReelParts = 0
                        Else
                            Summary_OUTList(Count).ReelParts = OUT_Data("ReelParts_NUM")
                        End If

                    Else
                        'データがなかった場合は全てに0を入れる。
                        ReDim Preserve Summary_OUTList(0 To Count)
                        Summary_OUTList(Count).WorkDate = Search_Date
                        Summary_OUTList(Count).Line = 0
                        Summary_OUTList(Count).Bait = 0
                        Summary_OUTList(Count).Rods = 0
                        Summary_OUTList(Count).B_Reels = 0
                        Summary_OUTList(Count).S_Reels = 0
                        Summary_OUTList(Count).Combos = 0
                        Summary_OUTList(Count).Accessory = 0
                        'Summary_OUTList(Count).Accessory2 = 0
                        Summary_OUTList(Count).Hard_Lure = 0
                        Summary_OUTList(Count).Bag = 0
                        Summary_OUTList(Count).Rod_OEM = 0
                        Summary_OUTList(Count).Rod_Parts = 0
                        Summary_OUTList(Count).ReelParts = 0
                    End If


                    '出庫情報を取得できたのでClose
                    OUT_Data.Close()

                    SQL = ""
                    '棚卸情報取得

                    'プロダクトラインごとの数量をSUMするようSQLを作成。
                    'LINE用
                    SQL = " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS LINE ON M_ITEM.PL_CODE =LINE.ID AND"
                    SQL &= " LINE.ID=1 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS LINE_NUM,"
                    'Bait用
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BAIT ON M_ITEM.PL_CODE =BAIT.ID AND"
                    SQL &= " BAIT.ID=2 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bait_NUM,"
                    'Rods
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Rods ON M_ITEM.PL_CODE =Rods.ID AND"
                    SQL &= " Rods.ID=3 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Rods_NUM,"
                    'Baitcast_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS BReels ON M_ITEM.PL_CODE =BReels.ID AND"
                    SQL &= " BReels.ID=4 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS BReels_NUM,"
                    'Spinning_Reels
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS SReels ON M_ITEM.PL_CODE =SReels.ID AND"
                    SQL &= " SReels.ID=5 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS SReels_NUM,"
                    'Combos
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Combos ON M_ITEM.PL_CODE =Combos.ID AND"
                    SQL &= " Combos.ID=6 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Combos_NUM,"
                    'Accessory
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    SQL &= " Accessory.ID=7 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Accessory_NUM,"
                    'Accessory
                    'SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Accessory ON M_ITEM.PL_CODE =Accessory.ID AND"
                    'SQL &= " Accessory.ID=9 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    'SQL &= Search_Date
                    'SQL &= "') AS Accessory2_NUM,"
                    'Hard_Lure
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Hard_Lure ON M_ITEM.PL_CODE =Hard_Lure.ID AND"
                    SQL &= " Hard_Lure.ID=11 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Hard_Lure_NUM,"
                    'BagApparel
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS Bag ON M_ITEM.PL_CODE =Bag.ID AND"
                    SQL &= " Bag.ID=12 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS Bag_NUM,"
                    'RodOEM
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodOEM ON M_ITEM.PL_CODE =RodOEM.ID AND"
                    SQL &= " RodOEM.ID=21 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodOEM_NUM,"
                    'RodParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS RodParts ON M_ITEM.PL_CODE =RodParts.ID AND"
                    SQL &= " RodParts.ID=24 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS RodParts_NUM,"
                    'ReelParts
                    SQL &= " (SELECT SUM( NUM ) FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID = M_ITEM.ID INNER JOIN M_PLINE AS ReelParts ON M_ITEM.PL_CODE =ReelParts.ID AND"
                    SQL &= " ReelParts.ID=99 AND STOCK_LOG.I_FLG=3 WHERE date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    SQL &= Search_Date
                    SQL &= "' AND STOCK_LOG.PLACE_ID ="
                    SQL &= Place
                    SQL &= ") AS ReelParts_NUM"


                    'データを取得するコマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " SELECT "
                    Command.CommandText &= SQL
                    Command.CommandText &= "  FROM STOCK_LOG INNER JOIN M_ITEM ON STOCK_LOG.I_ID=M_ITEM.ID"
                    Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
                    Command.CommandText &= " WHERE STOCK_LOG.I_FLG=3 AND date_format( STOCK_LOG.U_DATE, '%Y/%m/%d' ) = '"
                    Command.CommandText &= Search_Date
                    Command.CommandText &= "' GROUP BY STOCK_LOG.U_DATE"
                    'Select実行
                    TANAOROSHI_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                    If TANAOROSHI_Data.Read Then
                        ' レコードが取得できた時の処理
                        ReDim Preserve Summary_TanaoroshiList(0 To Count)
                        Summary_TanaoroshiList(Count).WorkDate = Search_Date
                        If IsDBNull(TANAOROSHI_Data("Line_NUM")) Then
                            Summary_TanaoroshiList(Count).Line = 0
                        Else
                            Summary_TanaoroshiList(Count).Line = TANAOROSHI_Data("Line_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("Bait_NUM")) Then
                            Summary_TanaoroshiList(Count).Bait = 0
                        Else
                            Summary_TanaoroshiList(Count).Bait = TANAOROSHI_Data("Bait_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("Rods_NUM")) Then
                            Summary_TanaoroshiList(Count).Rods = 0
                        Else
                            Summary_TanaoroshiList(Count).Rods = TANAOROSHI_Data("Rods_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("BReels_NUM")) Then
                            Summary_TanaoroshiList(Count).B_Reels = 0
                        Else
                            Summary_TanaoroshiList(Count).B_Reels = TANAOROSHI_Data("BReels_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("SReels_NUM")) Then
                            Summary_TanaoroshiList(Count).S_Reels = 0
                        Else
                            Summary_TanaoroshiList(Count).S_Reels = TANAOROSHI_Data("SReels_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("Combos_NUM")) Then
                            Summary_TanaoroshiList(Count).Combos = 0
                        Else
                            Summary_TanaoroshiList(Count).Combos = TANAOROSHI_Data("Combos_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("Accessory_NUM")) Then
                            Summary_TanaoroshiList(Count).Accessory = 0
                        Else
                            Summary_TanaoroshiList(Count).Accessory = TANAOROSHI_Data("Accessory_NUM")
                        End If

                        'If IsDBNull(TANAOROSHI_Data("Accessory2_NUM")) Then
                        '    Summary_TanaoroshiList(Count).Accessory2 = 0
                        'Else
                        '    Summary_TanaoroshiList(Count).Accessory2 = TANAOROSHI_Data("Accessory2_NUM")
                        'End If

                        If IsDBNull(TANAOROSHI_Data("Hard_Lure_NUM")) Then
                            Summary_TanaoroshiList(Count).Hard_Lure = 0
                        Else
                            Summary_TanaoroshiList(Count).Hard_Lure = TANAOROSHI_Data("Hard_Lure_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("Bag_NUM")) Then
                            Summary_TanaoroshiList(Count).Bag = 0
                        Else
                            Summary_TanaoroshiList(Count).Bag = TANAOROSHI_Data("Bag_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("RodOEM_NUM")) Then
                            Summary_TanaoroshiList(Count).Rod_OEM = 0
                        Else
                            Summary_TanaoroshiList(Count).Rod_OEM = TANAOROSHI_Data("RodOEM_NUM")
                        End If

                        If IsDBNull(TANAOROSHI_Data("RodParts_NUM")) Then
                            Summary_TanaoroshiList(Count).Rod_Parts = 0
                        Else
                            Summary_TanaoroshiList(Count).Rod_Parts = TANAOROSHI_Data("RodParts_NUM")
                        End If
                        If IsDBNull(TANAOROSHI_Data("ReelParts_NUM")) Then
                            Summary_TanaoroshiList(Count).ReelParts = 0
                        Else
                            Summary_TanaoroshiList(Count).ReelParts = TANAOROSHI_Data("ReelParts_NUM")
                        End If

                    Else
                        'データがなかった場合は全てに0を入れる。
                        ReDim Preserve Summary_TanaoroshiList(0 To Count)
                        Summary_TanaoroshiList(Count).WorkDate = Search_Date
                        Summary_TanaoroshiList(Count).Line = 0
                        Summary_TanaoroshiList(Count).Bait = 0
                        Summary_TanaoroshiList(Count).Rods = 0
                        Summary_TanaoroshiList(Count).B_Reels = 0
                        Summary_TanaoroshiList(Count).S_Reels = 0
                        Summary_TanaoroshiList(Count).Combos = 0
                        Summary_TanaoroshiList(Count).Accessory = 0
                        'Summary_TanaoroshiList(Count).Accessory2 = 0
                        Summary_TanaoroshiList(Count).Hard_Lure = 0
                        Summary_TanaoroshiList(Count).Bag = 0
                        Summary_TanaoroshiList(Count).Rod_OEM = 0
                        Summary_TanaoroshiList(Count).Rod_Parts = 0
                        Summary_TanaoroshiList(Count).ReelParts = 0

                    End If

                    '出庫情報を取得できたのでClose
                    TANAOROSHI_Data.Close()

                    '入庫、出庫、棚卸の３データの取得が完了したら日付を+1日する。
                    Search_Date = Search_Date.AddDays(1)

                    Count += 1

                Loop
            End If


        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' TMP_P_ORDERテーブルを作成する。
    ' <引数>
    ' 
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function CREATE_TMP_P_ORDER(ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean


        Dim TableCheck As MySqlDataReader
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Dim DBString As String = Constring

        'DB接続用文字列からDB名を取得
        Dim DBArrayData1 As String() = DBString.Split(";"c)
        Dim DBArrayData2 As String() = DBArrayData1(0).Split("="c)

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SHOW TABLES FROM `"
            Command.CommandText &= DBArrayData2(1)
            Command.CommandText &= "` like 'TMP_P_ORDER'"
            'Select実行
            TableCheck = Command.ExecuteReader(CommandBehavior.SingleRow)

            If TableCheck.Read Then
                TableCheck.Close()
                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " DROP TABLE TMP_P_ORDER;"
                '実行
                Command.ExecuteNonQuery()
            Else
                TableCheck.Close()
            End If


            'OUT_PRTテーブル用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " CREATE TABLE TMP_P_ORDER("
            Command.CommandText &= " I_ID			INTEGER NOT NULL		COMMENT '商品ID',"
            Command.CommandText &= " I_CODE			VARCHAR(24) NOT NULL		COMMENT '商品コード',"
            Command.CommandText &= " PO_CODE			VARCHAR(24) NOT NULL		COMMENT '発注先コード',"
            Command.CommandText &= " PO_NO			VARCHAR(20) NOT NULL		COMMENT '発注No',"
            Command.CommandText &= " PO_NUM			INTEGER NOT NULL		COMMENT '発注数',"
            Command.CommandText &= " PO_DATE			DATE NOT NULL			COMMENT '希望納期',"
            Command.CommandText &= " ORDER_DATE		DATE NOT NULL			COMMENT '発注日',"
            Command.CommandText &= " REMARKS			VARCHAR(500)  NULL		COMMENT '備考',"
            Command.CommandText &= " INDEX INDEX_1 (I_ID)"
            '八潮のサーバーでは以下で設定
            Command.CommandText &= " ) DEFAULT CHARSET=utf8;"
            '開発では以下で設定
            'Command.CommandText &= " ) TYPE = InnoDB, ENGINE = InnoDB, ROW_FORMAT = COMPACT, DEFAULT CHARSET=utf8, CHARACTER SET UTF8;"

            'OUT_PRTテーブルへデータ登録
            Command.ExecuteNonQuery()


            'コミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' TMP_P_ORDERテーブルを削除する。
    ' <引数>
    ' 
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function DROP_TMP_P_ORDER(ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " DROP TABLE TMP_P_ORDER;"
            '実行
            Command.ExecuteNonQuery()

            'コミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 発注データエクセルデータを登録。
    ' <引数>
    ' ExcelData : DataGridViewに取り込まれた発注データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_TMP_P_ORDER(ByRef ExcelData() As PO_List, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean


        Dim ItemID As MySqlDataReader
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim FileName_List() As String = Nothing
        Dim Count As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim PlusFlg As Boolean = False
        Dim MinusFlg As Boolean = False
        Dim ZeroFlg As Boolean = False


        Dim I_ID As Integer = 0
        Dim C_ID As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To ExcelData.Length - 1

                I_ID = 0
                C_ID = 0

                '商品IDを取得する。
                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " SELECT ID,C_ID FROM M_ITEM WHERE I_CODE='"
                Command.CommandText &= ExcelData(Count).I_CODE
                Command.CommandText &= "';"

                'データ取得
                ItemID = Command.ExecuteReader(CommandBehavior.SingleRow)
                If ItemID.Read Then
                    'レコードが取得できた時の処理()
                    If IsDBNull(ItemID("ID")) Then
                        I_ID = 0
                        C_ID = 0
                    Else
                        I_ID = ItemID("ID")
                        C_ID = ItemID("C_ID")

                    End If

                End If
                ItemID.Close()

                'TMP_P_ORDERテーブル用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO TMP_P_ORDER(I_ID,I_CODE,PO_CODE,PO_NO,PO_NUM,PO_DATE,ORDER_DATE,REMARKS)VALUES("
                Command.CommandText &= I_ID
                Command.CommandText &= ",'"
                Command.CommandText &= ExcelData(Count).I_CODE
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).PO_CODE
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).PO_NO
                Command.CommandText &= "',"
                Command.CommandText &= ExcelData(Count).PO_NUM
                Command.CommandText &= ",'"
                Command.CommandText &= ExcelData(Count).PO_DATE
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).ORDER_DATE
                Command.CommandText &= "','"
                Command.CommandText &= ExcelData(Count).REMARKS
                Command.CommandText &= "');"

                'TMP_P_ORDERテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next

            '全てのデータをTMP_P_ORDERテーブルにINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' POデータ取込後のTMPテーブルからデータ取得
    ' <引数>
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GET_TMP_P_ORDER(ByRef SearchResult() As PO_List, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()


            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT TMP_P_ORDER.I_ID,TMP_P_ORDER.I_CODE,TMP_P_ORDER.PO_CODE,M_ORDER.NAME AS PO_NAME,TMP_P_ORDER.PO_NUM,TMP_P_ORDER.PO_NO,"
            Command.CommandText &= "TMP_P_ORDER.PO_DATE,TMP_P_ORDER.ORDER_DATE,M_CUSTOMER.ID AS C_ID,M_CUSTOMER.C_CODE,TMP_P_ORDER.REMARKS,M_ITEM.I_NAME,M_ORDER.ID AS PO_M_ID, "
            Command.CommandText &= "(SELECT COUNT(*) FROM P_ORDER WHERE TMP_P_ORDER.I_ID = P_ORDER.I_ID AND TMP_P_ORDER.PO_NO = P_ORDER.PO_NO) AS R_CHECK, "
            Command.CommandText &= "(SELECT COUNT(*) FROM TMP_P_ORDER AS CHECK_TABLE WHERE CHECK_TABLE.PO_NO = TMP_P_ORDER.PO_NO AND CHECK_TABLE.I_CODE = TMP_P_ORDER.I_CODE) AS DUPLICATE_CHECK "
            Command.CommandText &= " FROM (TMP_P_ORDER) INNER JOIN M_ORDER ON TMP_P_ORDER.PO_CODE = M_ORDER.CODE"
            Command.CommandText &= " LEFT JOIN M_ITEM ON TMP_P_ORDER.I_ID=M_ITEM.ID"
            Command.CommandText &= " LEFT JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID "
            Command.CommandText &= " LEFT JOIN M_CUSTOMER ON M_ITEM.C_ID = M_CUSTOMER.ID;"

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)

                SearchResult(Count).I_ID = SearchData("I_ID")
                If SearchData("I_ID") = 0 Then

                End If

                If IsDBNull(SearchData("I_CODE")) Then
                    SearchResult(Count).I_CODE = 0
                Else
                    SearchResult(Count).I_CODE = SearchData("I_CODE")
                End If
                If IsDBNull(SearchData("I_NAME")) Then
                    SearchResult(Count).I_NAME = ""
                Else
                    SearchResult(Count).I_NAME = SearchData("I_NAME")
                End If

                If IsDBNull(SearchData("PO_CODE")) Then
                    SearchResult(Count).PO_CODE = 0
                Else
                    SearchResult(Count).PO_CODE = SearchData("PO_CODE")
                End If

                SearchResult(Count).PO_NAME = SearchData("PO_NAME")

                SearchResult(Count).PO_NO = SearchData("PO_NO")

                SearchResult(Count).PO_NUM = SearchData("PO_NUM")

                SearchResult(Count).PO_DATE = SearchData("PO_DATE")

                SearchResult(Count).ORDER_DATE = SearchData("ORDER_DATE")
                If IsDBNull(SearchData("C_ID")) Then
                    SearchResult(Count).C_ID = 0
                    SearchResult(Count).C_CODE = "0"
                Else
                    SearchResult(Count).C_ID = SearchData("C_ID")
                    SearchResult(Count).C_CODE = SearchData("C_CODE")
                End If

                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(Count).REMARKS = ""
                Else
                    SearchResult(Count).REMARKS = SearchData("REMARKS")
                End If
                SearchResult(Count).R_CHECK = SearchData("R_CHECK")

                SearchResult(Count).DUPLICATE_CHECK = SearchData("DUPLICATE_CHECK")

                SearchResult(Count).PO_M_ID = SearchData("PO_M_ID")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 発注データ取り込みデータをP_ORDERに登録する。
    ' <引数>
    ' Data_List : 登録データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_P_ORDER(ByRef Data_List() As PO_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For i = 0 To Data_List.Length - 1

                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO P_ORDER(I_ID,PO_NO,PO_NUM,PO_DATE,REMARKS,ORDER_DATE,PO_M_ID,STATUS,IN_NUM,FIX_NUM,U_DATE,U_USER,PLACE_ID)VALUES("
                Command.CommandText &= Data_List(i).I_ID
                Command.CommandText &= ",'"
                Command.CommandText &= Data_List(i).PO_NO
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(i).PO_NUM
                Command.CommandText &= ",'"
                Command.CommandText &= Data_List(i).PO_DATE
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(i).REMARKS
                Command.CommandText &= "','"
                Command.CommandText &= Data_List(i).ORDER_DATE
                Command.CommandText &= "',"
                Command.CommandText &= Data_List(i).PO_M_ID
                Command.CommandText &= ",1,0,0,Current_Timestamp,"
                Command.CommandText &= R_User
                Command.CommandText &= ",1);"

                'P_ORDERテーブルへデータ登録
                Command.ExecuteNonQuery()

            Next

            'IN_HEADERテーブル、IN_DETAILテーブルにINSERT、P_ORDERテーブルにUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 発注関連検索
    ' <引数>
    ' PO_No : 発注№
    ' PO_Status1 : 発注ステータス（発注済み）
    ' PO_Status2 : 発注ステータス（一部納品確定）
    ' PO_Status3 : 発注ステータス（全納品確定）
    ' PO_Status4 : 発注ステータス（一部キャンセル）
    ' PO_Status5 : 発注ステータス（全キャンセル）
    ' Vender_Code : ベンダーコード
    ' PL_ID : プロダクトラインID
    ' Item_Jan_Type : 1:商品コード、2:JANコード
    ' Item_Jan_Code : 商品コード or JANコード
    ' PO_DateFrom : 希望納期From
    ' PO_DateTo : 希望納期To
    ' INS_DateFrom : 取込日From
    ' INS_DateTo : 取込日To
    ' Remarks : 備考
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPoSeach(ByRef PO_No As String, _
                               ByRef PO_Status1 As String, _
                               ByRef PO_Status2 As String, _
                               ByRef PO_Status3 As String, _
                               ByRef PO_Status4 As String, _
                               ByRef PO_Status5 As Integer, _
                               ByRef Vender_Code As String, _
                               ByRef PL_ID As Integer, _
                               ByRef ItemJanType As Integer, _
                               ByRef ItemJanCode As String, _
                               ByRef PO_DateFrom As String, _
                               ByRef PO_DateTo As String, _
                               ByRef INS_DateFrom As String, _
                               ByRef INS_DateTo As String, _
                               ByRef Remarks As String, _
                               ByRef SearchResult() As PO_Search_List, _
                               ByRef Data_Total As Integer, _
                               ByRef Data_Num_Total As Integer, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            '検索条件よりWhereの作成
            '発注No
            If PO_No <> "" Then
                If WhereSql <> "" Then
                    WhereSql = " AND P_ORDER.PO_NO ='"
                    WhereSql &= PO_No
                    WhereSql &= "'"
                Else
                    WhereSql = " P_ORDER.PO_NO ='"
                    WhereSql &= PO_No
                    WhereSql &= "'"
                End If
            End If

            '発注ステータスの発注済、一部納品確定、全納品確定、一部キャンセル、全キャンセルにチェックが入っていたら
            '検索条件に加える。
            If PO_Status1 = True Or PO_Status2 = True Or PO_Status3 = True Or PO_Status4 = True Or PO_Status5 = True Then
                If PO_Status1 = True Then
                    StatusWhere &= "P_ORDER.STATUS in (1"
                End If

                If PO_Status2 = True And StatusWhere = "" Then
                    StatusWhere &= "P_ORDER.STATUS in (2"
                ElseIf PO_Status2 = True And StatusWhere <> "" Then
                    StatusWhere &= ",2"
                End If

                If PO_Status3 = True And StatusWhere = "" Then
                    StatusWhere &= "P_ORDER.STATUS in (3"
                ElseIf PO_Status3 = True And StatusWhere <> "" Then
                    StatusWhere &= ",3"
                End If

                If PO_Status4 = True And StatusWhere = "" Then
                    StatusWhere &= "P_ORDER.STATUS in (4"
                ElseIf PO_Status4 = True And StatusWhere <> "" Then
                    StatusWhere &= ",4"
                End If

                If PO_Status5 = True And StatusWhere = "" Then
                    StatusWhere &= "P_ORDER.STATUS in (5"
                ElseIf PO_Status5 = True And StatusWhere <> "" Then
                    StatusWhere &= ",5"
                End If

                StatusWhere &= ") "

                If WhereSql = "" Then
                    WhereSql = StatusWhere
                Else
                    WhereSql &= " AND " & StatusWhere
                End If
            End If

            'ベンダーコード
            If Vender_Code <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND M_ORDER.CODE ='"
                    WhereSql &= Vender_Code
                    WhereSql &= "'"
                Else
                    WhereSql &= " M_ORDER.CODE ='"
                    WhereSql &= Vender_Code
                    WhereSql &= "'"
                End If

            End If

            'プロダクトラインコード
            'プロダクトライン
            If PL_ID <> 0 Then
                If WhereSql = "" Then
                    WhereSql = " M_PLINE.ID = "
                    WhereSql &= PL_ID
                Else
                    WhereSql &= " AND M_PLINE.ID="
                    WhereSql &= PL_ID
                End If
            End If

            '備考
            If Remarks <> "" Then
                If WhereSql = "" Then
                    WhereSql = " P_ORDER.REMARKS LIKE '%"
                    WhereSql &= Remarks
                    WhereSql &= "%'"

                Else
                    WhereSql &= "AND P_ORDER.REMARKS LIKE '%"
                    WhereSql &= Remarks
                    WhereSql &= "%'"
                End If
            End If

            'ItemJan_Flgが1なら商品コードを検索する。
            If ItemJanType = 1 And ItemJanCode <> "" Then
                If WhereSql = "" Then
                    WhereSql &= " M_ITEM.I_CODE ='"
                    WhereSql &= ItemJanCode
                    WhereSql &= "' "
                Else
                    WhereSql &= " AND M_ITEM.I_CODE ='"
                    WhereSql &= ItemJanCode
                    WhereSql &= "' "
                End If
            ElseIf ItemJanType = 2 And ItemJanCode <> "" Then
                If WhereSql = "" Then
                    WhereSql &= " M_ITEM.JAN ='"
                    WhereSql &= ItemJanCode
                    WhereSql &= "' "
                Else
                    WhereSql &= " AND M_ITEM.JAN ='"
                    WhereSql &= ItemJanCode
                    WhereSql &= "' "
                End If
            End If

            '希望納期From
            If PO_DateFrom <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND P_ORDER.PO_DATE >= '"
                    WhereSql &= PO_DateFrom
                    WhereSql &= "'"
                Else
                    WhereSql &= " P_ORDER.PO_DATE >= '"
                    WhereSql &= PO_DateFrom
                    WhereSql &= "'"
                End If

            End If

            '希望納期To
            If PO_DateTo <> "" Then
                If WhereSql <> "" Then
                    WhereSql &= " AND P_ORDER.PO_DATE <= '"
                    WhereSql &= PO_DateTo
                    WhereSql &= "'"
                Else
                    WhereSql &= " P_ORDER.PO_DATE <= '"
                    WhereSql &= PO_DateTo
                    WhereSql &= "'"
                End If
            End If

            '取込日From
            If INS_DateFrom <> "" Then
                If WhereSql = "" Then
                    WhereSql = " date_format( P_ORDER.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= INS_DateFrom
                    WhereSql &= "'"
                Else
                    WhereSql &= " AND date_format( P_ORDER.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= INS_DateFrom
                    WhereSql &= "'"
                End If
            End If

            '取込日To
            If INS_DateTo <> "" Then
                If WhereSql = "" Then
                    WhereSql = " date_format( P_ORDER.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= INS_DateTo
                    WhereSql &= "'"
                Else
                    WhereSql &= " AND date_format( P_ORDER.U_DATE, '%Y/%m/%d' ) >= '"
                    WhereSql &= INS_DateTo
                    WhereSql &= "'"
                End If
            End If

            'もし検索条件が入力されていたらWhere句を追加する。
            If WhereSql <> "" Then
                WhereSql = "Where " & WhereSql
            End If

            'オープン
            Connection.Open()

            'データ件数取得用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT COUNT(*) AS COUNT FROM P_ORDER INNER JOIN M_ITEM ON P_ORDER.I_ID=M_ITEM.ID "
            Command.CommandText &= "LEFT JOIN M_CUSTOMER ON M_ITEM.C_ID = M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID "
            Command.CommandText &= "INNER JOIN M_ORDER ON P_ORDER.PO_M_ID = M_ORDER.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= ";"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                Data_Total = CStr(SearchData("COUNT"))
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "データがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT  SUM(PO_NUM) AS TOTAL FROM P_ORDER INNER JOIN M_ITEM ON P_ORDER.I_ID=M_ITEM.ID "
            Command.CommandText &= "LEFT JOIN M_CUSTOMER ON M_ITEM.C_ID = M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID "
            Command.CommandText &= "INNER JOIN M_ORDER ON P_ORDER.PO_M_ID = M_ORDER.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= ";"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理

                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0

                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "データがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT P_ORDER.ID,P_ORDER.I_ID,P_ORDER.PO_NO,P_ORDER.PO_NUM,P_ORDER.PO_DATE,P_ORDER.REMARKS,P_ORDER.STATUS,"
            Command.CommandText &= " P_ORDER.IN_NUM,P_ORDER.IN_DATE,P_ORDER.FIX_NUM,P_ORDER.FIX_DATE,DATE_FORMAT(P_ORDER.ORDER_DATE, '%Y/%m/%d') AS ORDER_DATE,M_ITEM.I_CODE,M_ORDER.NAME AS PO_M_NAME,"
            Command.CommandText &= " M_ITEM.I_NAME,M_PLINE.NAME AS PL_NAME,M_CUSTOMER.C_CODE,P_ORDER.CANCEL_NUM,P_ORDER.CANCEL_DATE,M_CUSTOMER.D_NAME AS C_NAME, "
            Command.CommandText &= " M_ORDER.CODE AS PO_M_CODE "
            Command.CommandText &= " FROM (P_ORDER) INNER JOIN M_ITEM ON P_ORDER.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID "
            Command.CommandText &= " LEFT JOIN M_CUSTOMER ON M_ITEM.C_ID = M_CUSTOMER.ID "
            Command.CommandText &= " INNER JOIN M_ORDER ON P_ORDER.PO_M_ID = M_ORDER.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " order by P_ORDER.ID"

            'データ取得
            SearchData = Command.ExecuteReader()

            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)

                SearchResult(Count).ID = SearchData("ID")
                SearchResult(Count).I_ID = SearchData("I_ID")
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                If IsDBNull(SearchData("C_CODE")) Then
                    SearchResult(Count).C_CODE = ""
                Else
                    SearchResult(Count).C_CODE = SearchData("C_CODE")
                End If
                SearchResult(Count).PO_NO = SearchData("PO_NO")
                SearchResult(Count).PO_NUM = SearchData("PO_NUM")
                SearchResult(Count).PO_NO = SearchData("PO_NO")
                SearchResult(Count).PO_DATE = SearchData("PO_DATE")

                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(Count).REMARKS = ""
                Else
                    SearchResult(Count).REMARKS = SearchData("REMARKS")
                End If

                SearchResult(Count).STATUS = SearchData("STATUS")
                SearchResult(Count).IN_NUM = SearchData("IN_NUM")

                If IsDBNull(SearchData("IN_DATE")) Then
                    SearchResult(Count).IN_DATE = ""
                Else
                    SearchResult(Count).IN_DATE = SearchData("IN_DATE")
                End If

                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")

                If IsDBNull(SearchData("FIX_DATE")) Then
                    SearchResult(Count).FIX_DATE = ""
                Else
                    SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                End If

                SearchResult(Count).CANCEL_NUM = SearchData("CANCEL_NUM")

                If IsDBNull(SearchData("CANCEL_DATE")) Then
                    SearchResult(Count).CANCEL_DATE = ""
                Else
                    SearchResult(Count).CANCEL_DATE = SearchData("CANCEL_DATE")
                End If

                SearchResult(Count).ORDER_DATE = SearchData("ORDER_DATE")
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                SearchResult(Count).PL_NAME = SearchData("PL_NAME")


                'ベンダー名（工場名）
                If IsDBNull(SearchData("C_NAME")) Then
                    SearchResult(Count).C_NAME = ""
                Else
                    'ベンダー名（工場名）
                    SearchResult(Count).C_NAME = SearchData("C_NAME")
                End If
                '発注先名
                SearchResult(Count).PO_M_NAME = SearchData("PO_M_NAME")
                '発注先コード
                SearchResult(Count).PO_M_CODE = SearchData("PO_M_CODE")
                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 登録するデータのInvoiceNoがすでにIN_HEADERに存在するかチェックを行う。
    ' <引数>
    ' Data_List : 登録データリスト

    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function PO_Invoice_Check(ByRef Data_List() As PO_In_List, _
                                ByRef Check_Result As Boolean, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            For i = 0 To Data_List.Length - 1

                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT COUNT(*) AS COUNT FROM IN_HEADER WHERE SHEET_NO='"
                Command.CommandText &= Data_List(i).INVOICE_NO
                Command.CommandText &= "'"

                'Select実行
                SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

                If SearchData.Read Then
                    ' レコードが取得できた時の処理
                    If SearchData("COUNT") <> 0 Then
                        Check_Result = True
                    End If

                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "データがみつかりません。"
                    Exit Function
                End If
                SearchData.Close()
            Next

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品を登録する（POからの入荷予定入力）
    ' <引数>
    ' Data_List : 登録データ
    ' Defect_Type : 不良区分
    ' Category : 種別
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function PO_InsItem(ByRef Data_List() As PO_In_List, _
                            ByRef Defect_Type As Integer, _
                            ByRef Category As Integer, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Id_Data As MySqlDataReader

        Dim Invoice_Data As MySqlDataReader
        Dim Invoice_Count As Integer
        Dim Header_ID As Integer

        Dim P_Order_Data As MySqlDataReader
        Dim C_ID As Integer
        Dim MAX_ID As Integer

        Dim TMP_IN_NUM As Integer
        Dim TMP_IN_DATE As String
        Dim TMP_PO_NUM As Integer
        Dim TMP_STATUS As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For i = 0 To Data_List.Length - 1
                '登録する商品のC_ID情報を取得
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT C_ID FROM M_ITEM WHERE ID="
                Command.CommandText &= Data_List(i).I_ID
                Id_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
                'Command.CommandText = "SELECT ID AS C_ID FROM M_CUSTOMER WHERE C_CODE='"
                'Command.CommandText &= Data_List(i).VENDER_CODE
                'Command.CommandText &= "';"
                'Id_Data = Command.ExecuteReader(CommandBehavior.SingleRow)

                If Id_Data.Read Then
                    ' レコードが取得できた時の処理
                    C_ID = Id_Data("C_ID")
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "入力した商品コードに該当する企業データがみつかりません。"
                    Exit Function
                End If
                'MAX値を取得できたのでClose
                Id_Data.Close()

                'IN_HEADERテーブルにインボイスNOが存在するかチェック。
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT COUNT(*) AS COUNT,ID FROM IN_HEADER WHERE SHEET_NO='"
                Command.CommandText &= Data_List(i).INVOICE_NO
                Command.CommandText &= "'"
                Invoice_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
                If Invoice_Data.Read Then

                    Invoice_Count = Invoice_Data("COUNT")
                    If IsDBNull(Invoice_Data("ID")) Then

                    Else
                        Header_ID = Invoice_Data("ID")
                    End If

                    Invoice_Data.Close()
                    ' レコードが取得できた時の処理(DETAILのみInsert)
                    If Invoice_Count <> 0 Then
                        'IN_DETAILテーブル用　コマンド作成
                        Command = Connection.CreateCommand
                        Command.CommandText = "INSERT INTO IN_DETAIL(ID,STATUS,CATEGORY,I_ID,I_STATUS,NUM,FIX_NUM,DATE,PO_ID,U_USER)VALUES("
                        Command.CommandText &= Header_ID
                        Command.CommandText &= ",1,'"
                        Command.CommandText &= Category
                        Command.CommandText &= "','"
                        Command.CommandText &= Data_List(i).I_ID
                        Command.CommandText &= "','"
                        Command.CommandText &= Defect_Type
                        Command.CommandText &= "',"
                        Command.CommandText &= Data_List(i).IN_NUM
                        Command.CommandText &= ",0,'"
                        Command.CommandText &= Data_List(i).IN_DATE
                        Command.CommandText &= "',"
                        Command.CommandText &= Data_List(i).ID
                        Command.CommandText &= ","
                        Command.CommandText &= R_User
                        Command.CommandText &= ")"
                        'Insert実行
                        Command.ExecuteNonQuery()
                    Else
                        'レコードが取得できなかった時の処理（HEADER,DETAILのInsert）

                        'INテーブル用　コマンド作成
                        Command = Connection.CreateCommand
                        Command.CommandText = "INSERT INTO IN_HEADER(SHEET_NO,C_ID,PLACE_ID,U_USER)VALUES('"
                        Command.CommandText &= Data_List(i).INVOICE_NO
                        Command.CommandText &= "',"
                        Command.CommandText &= C_ID
                        Command.CommandText &= ","
                        Command.CommandText &= Data_List(i).PLACE
                        Command.CommandText &= ","
                        Command.CommandText &= R_User
                        Command.CommandText &= ");"

                        'INテーブルへデータ登録
                        Command.ExecuteNonQuery()

                        '上記で登録したデータのIDを取得する。
                        Command = Connection.CreateCommand
                        Command.CommandText = "SELECT MAX(ID) as ID FROM IN_HEADER"
                        Id_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
                        If Id_Data.Read Then
                            ' レコードが取得できた時の処理
                            MAX_ID = CStr(Id_Data("ID"))
                        Else
                            ' レコードが取得できなかった時の処理
                            ErrorMessage = "入力した商品コードに該当する商品がみつかりません。"
                            Exit Function
                        End If
                        'MAX値を取得できたのでClose
                        Id_Data.Close()

                        'IN_DETAILテーブル用　コマンド作成
                        Command = Connection.CreateCommand
                        Command.CommandText = "INSERT INTO IN_DETAIL(ID,STATUS,CATEGORY,I_ID,I_STATUS,NUM,FIX_NUM,DATE,PO_ID,U_USER)VALUES("
                        Command.CommandText &= MAX_ID
                        Command.CommandText &= ",1,'"
                        Command.CommandText &= Category
                        Command.CommandText &= "','"
                        Command.CommandText &= Data_List(i).I_ID
                        Command.CommandText &= "','"
                        Command.CommandText &= Defect_Type
                        Command.CommandText &= "',"
                        Command.CommandText &= Data_List(i).IN_NUM
                        Command.CommandText &= ",0,'"
                        Command.CommandText &= Data_List(i).IN_DATE
                        Command.CommandText &= "',"
                        Command.CommandText &= Data_List(i).ID
                        Command.CommandText &= ","
                        Command.CommandText &= R_User
                        Command.CommandText &= ")"
                        'Insert実行
                        Command.ExecuteNonQuery()
                    End If

                Else
                    Invoice_Data.Close()
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "該当データがみつかりません。"
                    Exit Function
                End If

                'P_ORDERのデータを取得する。
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT PO_NUM,IN_NUM,IN_DATE FROM P_ORDER WHERE ID="
                Command.CommandText &= Data_List(i).ID


                P_Order_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
                If P_Order_Data.Read Then
                    ' レコードが取得できた時の処理
                    TMP_PO_NUM = P_Order_Data("PO_NUM")
                    TMP_IN_NUM = P_Order_Data("IN_NUM")
                    If IsDBNull(P_Order_Data("IN_DATE")) Then
                        TMP_IN_DATE = ""
                    Else
                        TMP_IN_DATE = P_Order_Data("IN_DATE")
                    End If

                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "該当データがみつかりません。"
                    Exit Function
                End If
                'MAX値を取得できたのでClose
                P_Order_Data.Close()


                '取得してきたIN_NUMにIN_NUMを足す。
                TMP_IN_NUM += Data_List(i).IN_NUM
                '取得してきたIN_DATEにIN_DATEを追加。
                If TMP_IN_DATE = "" Then
                    TMP_IN_DATE = Data_List(i).IN_DATE
                Else
                    TMP_IN_DATE = TMP_IN_DATE & "、" & Data_List(i).IN_DATE
                End If

                'もしTMP_PO_NUMとTMP_IN_NUM+Data_List(i).IN_NUMが同じだったら全納品確定のフラグでUPDATEする。
                If TMP_PO_NUM = TMP_IN_NUM Then
                    TMP_STATUS = 3
                Else
                    TMP_STATUS = 2
                End If

                'P_ORDERテーブルをアップデートする。
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE P_ORDER SET IN_NUM="
                Command.CommandText &= TMP_IN_NUM
                Command.CommandText &= ",IN_DATE='"
                Command.CommandText &= TMP_IN_DATE
                Command.CommandText &= "',STATUS="
                Command.CommandText &= TMP_STATUS
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Data_List(i).ID
                'Update実行
                Command.ExecuteNonQuery()
            Next

            'IN_HEADERテーブル、IN_DETAILテーブルにINSERT、P_ORDERテーブルにUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 発注キャンセルを行う。
    ' <引数>
    ' Cancel_List : 削除する商品IDが格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function PO_Upd_Cancel(ByRef Cancel_List() As PO_In_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim P_Order_Data As MySqlDataReader

        Dim dtNow As DateTime = DateTime.Now

        Dim CANCEL_DATE As String
        Dim CANCEL_NUM As Integer

        Dim STATUS As Integer


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            '件数分ループ
            For Count = 0 To Cancel_List.Length - 1
                'P_ORDERのデータを取得する。
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT CANCEL_NUM,CANCEL_DATE FROM P_ORDER WHERE ID="
                Command.CommandText &= Cancel_List(Count).ID

                P_Order_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
                If P_Order_Data.Read Then
                    ' レコードが取得できた時の処理
                    CANCEL_NUM = P_Order_Data("CANCEL_NUM") + Cancel_List(Count).CANCEL_NUM

                    If IsDBNull(P_Order_Data("CANCEL_DATE")) Then
                        CANCEL_DATE = dtNow.ToString("yyyy/MM/dd")
                    Else
                        CANCEL_DATE = P_Order_Data("CANCEL_DATE") & "、" & dtNow.ToString("yyyy/MM/dd")
                    End If
                Else
                    ' レコードが取得できなかった時の処理
                    ErrorMessage = "該当データがみつかりません。"
                    Exit Function
                End If
                'MAX値を取得できたのでClose
                P_Order_Data.Close()

                'もし納品確定（入庫予定）数に数値が入っていれば、一部キャンセル
                '発注数＝今回入力したキャンセル数ならば全キャンセルとしてステータスをUPDATEする。
                If Cancel_List(Count).IN_NUM <> 0 Then
                    '一部キャンセル
                    STATUS = 4
                ElseIf Cancel_List(Count).PO_NUM = Cancel_List(Count).CANCEL_NUM Then
                    '全キャンセル
                    STATUS = 5
                End If

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE P_ORDER SET CANCEL_NUM="
                Command.CommandText &= CANCEL_NUM
                Command.CommandText &= ",CANCEL_DATE='"
                Command.CommandText &= CANCEL_DATE
                Command.CommandText &= "', STATUS="
                Command.CommandText &= STATUS
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Cancel_List(Count).ID
                'Delete実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 選択された発注データを削除する。
    ' <引数>
    ' Del_List : 削除する商品IDが格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Del_Po(ByRef Del_List() As PO_In_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'DETAILテーブル 削除
            For Count = 0 To Del_List.Length - 1
                Command = Connection.CreateCommand
                Command.CommandText = "DELETE FROM P_ORDER WHERE ID="
                Command.CommandText &= Del_List(Count).ID
                Command.CommandText &= ";"
                'Delete実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 月別検索のデータ取得
    ' <引数>
    ' Date_From : 対象年月From
    ' Date_To   : 対象年月To
    ' ItemJan_Flg : 商品コードかJANコードかの判別フラグ
    ' ItemJanCode : 商品コードかJANコードを格納
    ' ItemName : 商品名を格納
    ' PL_Id : プロダクトラインID
    ' ZeroDataFlg : 数量0のデータを表示するかしないか。
    ' Datatype : 検索データ
    ' Place : 倉庫
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Get_Monthly_Report(ByRef Date_From As String, _
                                       ByRef Date_To As String, _
                                       ByRef ItemJan_Flg As Integer, _
                                       ByRef ItemJanCode As String, _
                                       ByRef ItemName As String, _
                                       ByRef PL_Id As Integer, _
                                       ByRef ZeroDataFlg As Integer, _
                                       ByRef Datatype As Integer, _
                                       ByRef Place As Integer, _
                                       ByRef SearchResult() As Monthly_Report_Search_List, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0

        Dim Wheresql As String = Nothing
        Dim Selectsql As String = Nothing
        Dim TMP_Date As Date
        Dim TMP_Date_String As String

        Dim checkflg As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            '商品コード or JANコードが入力されていれば追加
            If ItemJanCode <> "" Then
                '商品コード
                If ItemJan_Flg = 1 Then
                    Wheresql = "AND M_ITEM.I_CODE='"
                    Wheresql &= ItemJanCode
                    Wheresql &= "'"
                ElseIf ItemJan_Flg = 2 Then
                    Wheresql = "AND M_ITEM.JAN='"
                    Wheresql &= ItemJanCode
                    Wheresql &= "'"
                End If
            End If

            '商品名が設定されていれば追加
            If ItemName <> "" Then
                Wheresql = "AND M_ITEM.I_NAME like '%"
                Wheresql &= ItemName
                Wheresql &= "%'"
            End If

            'プロダクトラインコードが設定されていれば追加
            If PL_Id <> 0 Then
                Wheresql = "AND M_PLINE.ID="
                Wheresql &= PL_Id
                Wheresql &= " "
            End If

            '検索データが発注数の場合
            If Datatype = 1 Then
                TMP_Date_String = Date_From
                While checkflg = 0
                    Selectsql &= " sum(field(DATE_FORMAT(P_ORDER.U_DATE, '%Y/%m'), '"
                    Selectsql &= TMP_Date_String
                    Selectsql &= "')*PO_NUM) as ym"
                    Selectsql &= MonthCount

                    TMP_Date = Date.Parse(Date_From).AddMonths(MonthCount)
                    TMP_Date_String = TMP_Date.ToString("yyyy/MM")
                    If TMP_Date_String > Date_To Then
                        checkflg = 1
                    Else
                        Selectsql &= ","
                    End If
                    MonthCount += 1
                End While

                'From
                Wheresql = "AND DATE_FORMAT(P_ORDER.U_DATE,'%Y/%m')>='"
                Wheresql &= Date_From
                Wheresql &= "' "
                'To
                Wheresql &= "AND DATE_FORMAT(P_ORDER.U_DATE,'%Y/%m')<='"
                Wheresql &= Date_To
                Wheresql &= "' "


                '検索データが発注数の場合
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT P_ORDER.I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_PLINE.NAME AS PL_NAME,M_ITEM.JAN, "
                Command.CommandText &= Selectsql
                Command.CommandText &= " FROM P_ORDER INNER JOIN M_ITEM ON P_ORDER.I_ID= M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID"
                Command.CommandText &= " WHERE P_ORDER.STATUS=1 "
                Command.CommandText &= Wheresql
                Command.CommandText &= " GROUP BY P_ORDER.I_ID"


                '検索データが入庫確定数の場合
            ElseIf Datatype = 2 Then

                If ZeroDataFlg = 2 Then
                    Wheresql = "AND IN_DETAIL.NUM <> 0"
                End If

                TMP_Date_String = Date_From
                While checkflg = 0
                    Selectsql &= " sum(field(DATE_FORMAT(FIX_DATE, '%Y/%m'), '"
                    Selectsql &= TMP_Date_String
                    Selectsql &= "')*FIX_NUM) as ym"
                    Selectsql &= MonthCount

                    TMP_Date = Date.Parse(Date_From).AddMonths(MonthCount)
                    TMP_Date_String = TMP_Date.ToString("yyyy/MM")
                    If TMP_Date_String > Date_To Then
                        checkflg = 1
                    Else
                        Selectsql &= ","
                    End If
                    MonthCount += 1
                End While

                'From
                Wheresql = "AND DATE_FORMAT(IN_DETAIL.FIX_DATE,'%Y/%m')>='"
                Wheresql &= Date_From
                Wheresql &= "' "
                'To
                Wheresql &= "AND DATE_FORMAT(IN_DETAIL.FIX_DATE,'%Y/%m')<='"
                Wheresql &= Date_To
                Wheresql &= "' "

                '検索データが入庫確定の場合
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_PLINE.NAME AS PL_NAME,M_ITEM.JAN, "
                Command.CommandText &= Selectsql
                Command.CommandText &= " FROM IN_DETAIL INNER JOIN M_ITEM ON IN_DETAIL.I_ID= M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID"
                Command.CommandText &= " INNER JOIN IN_HEADER ON IN_HEADER.ID = IN_DETAIL.ID"
                Command.CommandText &= " WHERE IN_DETAIL.STATUS=2 AND IN_HEADER.PLACE_ID= "
                Command.CommandText &= Place
                Command.CommandText &= " "
                Command.CommandText &= Wheresql
                Command.CommandText &= " GROUP BY I_ID"

            ElseIf Datatype = 3 Then
                '検索データが出庫確定の場合
                If ZeroDataFlg = 2 Then
                    Wheresql = "AND OUT_TBL.NUM <> 0"
                End If

                TMP_Date_String = Date_From
                While checkflg = 0
                    Selectsql &= " sum(field(DATE_FORMAT(FIX_DATE, '%Y/%m'), '"
                    Selectsql &= TMP_Date_String
                    Selectsql &= "')*FIX_NUM) as ym"
                    Selectsql &= MonthCount

                    TMP_Date = Date.Parse(Date_From).AddMonths(MonthCount)
                    TMP_Date_String = TMP_Date.ToString("yyyy/MM")
                    If TMP_Date_String > Date_To Then
                        checkflg = 1
                    Else
                        Selectsql &= ","
                    End If
                    MonthCount += 1
                End While

                'From
                Wheresql = "AND DATE_FORMAT(OUT_TBL.FIX_DATE,'%Y/%m')>='"
                Wheresql &= Date_From
                Wheresql &= "' "
                'To
                Wheresql &= "AND DATE_FORMAT(OUT_TBL.FIX_DATE,'%Y/%m')<='"
                Wheresql &= Date_To
                Wheresql &= "' "

                '検索データが出庫確定の場合
                Command = Connection.CreateCommand
                Command.CommandText = "SELECT I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_PLINE.NAME AS PL_NAME,M_ITEM.JAN, "
                Command.CommandText &= "(SELECT SUM(PO_NUM) FROM P_ORDER WHERE P_ORDER.I_ID=OUT_TBL.I_ID AND P_ORDER.STATUS=1) AS ORDER_NUM,"
                Command.CommandText &= "(SELECT SUM(NUM) FROM STOCK WHERE STOCK.I_ID=OUT_TBL.I_ID AND I_STATUS=1 AND PLACE_ID="
                Command.CommandText &= Place
                Command.CommandText &= ") AS STOCK_NUM,"
                Command.CommandText &= Selectsql
                Command.CommandText &= " FROM OUT_TBL INNER JOIN M_ITEM ON OUT_TBL.I_ID= M_ITEM.ID"
                Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID"
                Command.CommandText &= " WHERE OUT_TBL.STATUS=4 AND OUT_TBL.PLACE_ID= "
                Command.CommandText &= Place
                Command.CommandText &= " "
                Command.CommandText &= Wheresql
                Command.CommandText &= " GROUP BY I_ID"

            End If
            'コマンド作成

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)

                SearchResult(Count).I_ID = SearchData("I_ID")

                If IsDBNull(SearchData("I_CODE")) Then
                    SearchResult(Count).I_CODE = 0
                Else
                    SearchResult(Count).I_CODE = SearchData("I_CODE")
                End If
                If IsDBNull(SearchData("I_NAME")) Then
                    SearchResult(Count).I_NAME = ""
                Else
                    SearchResult(Count).I_NAME = SearchData("I_NAME")
                End If

                SearchResult(Count).PL_NAME = SearchData("PL_NAME")
                SearchResult(Count).JAN = SearchData("JAN")
                SearchResult(Count).Month1 = SearchData("ym1")
                If Datatype = 1 Then

                    'If IsDBNull(SearchData("IN_NUM")) Then
                    '    SearchResult(Count).IN_NUM = 0
                    'Else
                    '    SearchResult(Count).IN_NUM = SearchData("IN_NUM")
                    'End If

                ElseIf Datatype = 3 Then


                    If IsDBNull(SearchData("ORDER_NUM")) Then
                        SearchResult(Count).ORDER_NUM = 0
                    Else
                        SearchResult(Count).ORDER_NUM = SearchData("ORDER_NUM")
                    End If

                    If IsDBNull(SearchData("STOCK_NUM")) Then
                        SearchResult(Count).STOCK_NUM = 0
                    Else
                        SearchResult(Count).STOCK_NUM = SearchData("STOCK_NUM")
                    End If
                End If
                If MonthCount > 1 Then
                    If IsDBNull(SearchData("ym1")) Then
                        SearchResult(Count).Month1 = "0"
                    Else
                        SearchResult(Count).Month1 = SearchData("ym1")
                    End If
                End If
                If MonthCount > 2 Then
                    If IsDBNull(SearchData("ym2")) Then
                        SearchResult(Count).Month2 = "0"
                    Else
                        SearchResult(Count).Month2 = SearchData("ym2")
                    End If
                End If
                If MonthCount > 3 Then
                    If IsDBNull(SearchData("ym3")) Then
                        SearchResult(Count).Month3 = "0"
                    Else
                        SearchResult(Count).Month3 = SearchData("ym3")
                    End If
                End If
                If MonthCount > 4 Then
                    If IsDBNull(SearchData("ym4")) Then
                        SearchResult(Count).Month4 = "0"
                    Else
                        SearchResult(Count).Month4 = SearchData("ym4")
                    End If
                End If
                If MonthCount > 5 Then
                    If IsDBNull(SearchData("ym5")) Then
                        SearchResult(Count).Month5 = "0"
                    Else
                        SearchResult(Count).Month5 = SearchData("ym5")
                    End If
                End If
                If MonthCount > 6 Then
                    If IsDBNull(SearchData("ym6")) Then
                        SearchResult(Count).Month6 = "0"
                    Else
                        SearchResult(Count).Month6 = SearchData("ym6")
                    End If
                End If
                If MonthCount > 7 Then
                    If IsDBNull(SearchData("ym7")) Then
                        SearchResult(Count).Month7 = "0"
                    Else
                        SearchResult(Count).Month7 = SearchData("ym7")
                    End If
                End If
                If MonthCount > 8 Then
                    If IsDBNull(SearchData("ym8")) Then
                        SearchResult(Count).Month8 = "0"
                    Else
                        SearchResult(Count).Month8 = SearchData("ym8")
                    End If
                End If
                If MonthCount > 9 Then
                    If IsDBNull(SearchData("ym9")) Then
                        SearchResult(Count).Month9 = "0"
                    Else
                        SearchResult(Count).Month9 = SearchData("ym9")
                    End If
                End If
                If MonthCount > 10 Then
                    If IsDBNull(SearchData("ym10")) Then
                        SearchResult(Count).Month10 = "0"
                    Else
                        SearchResult(Count).Month10 = SearchData("ym10")
                    End If
                End If
                If MonthCount > 11 Then
                    If IsDBNull(SearchData("ym11")) Then
                        SearchResult(Count).Month11 = "0"
                    Else
                        SearchResult(Count).Month11 = SearchData("ym11")
                    End If
                End If
                If MonthCount > 12 Then
                    If IsDBNull(SearchData("ym12")) Then
                        SearchResult(Count).Month12 = "0"
                    Else
                        SearchResult(Count).Month12 = SearchData("ym12")
                    End If
                End If

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品データ取得
    ' <引数>
    ' Itemcode : 商品コード
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Get_Mitem_Modify(ByRef Itemcode As String, _
                                       ByRef SearchResult() As M_Item_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0
        Dim Wheresql As String = Nothing
        Dim Itemsql As String = Nothing

        Dim checkflg As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_ITEM.ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,M_ITEM.PRICE,M_ITEM.PACKAGE_FLG, M_ITEM.IN_BOX_NUM,"
            Command.CommandText &= " M_ITEM.MASTER_CARTON_SIZE,M_ITEM.LOCATION,M_ITEM.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.C_NAME,"
            Command.CommandText &= " M_ITEM.PL_CODE AS PL_ID,M_PLINE.NAME AS PL_NAME,M_ITEM.PURCHASE_PRICE,M_ITEM.IMMUNITY_PRICE,M_ITEM.REPAIR_PRICE "
            Command.CommandText &= " FROM M_ITEM INNER JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID "
            Command.CommandText &= " LEFT JOIN M_CUSTOMER ON M_ITEM.C_ID = M_CUSTOMER.ID "
            Command.CommandText &= " WHERE M_ITEM.I_CODE='"
            Command.CommandText &= Itemcode
            Command.CommandText &= "'"


            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '商品ID
                SearchResult(Count).ID = SearchData("ID")

                '商品コード
                If IsDBNull(SearchData("I_CODE")) Then
                    SearchResult(Count).I_CODE = ""
                Else
                    SearchResult(Count).I_CODE = SearchData("I_CODE")
                End If


                '商品名
                If IsDBNull(SearchData("I_NAME")) Then
                    SearchResult(Count).I_NAME = ""
                Else
                    SearchResult(Count).I_NAME = SearchData("I_NAME")
                End If
                'JANコード
                If IsDBNull(SearchData("JAN")) Then
                    SearchResult(Count).JAN = ""
                Else
                    SearchResult(Count).JAN = SearchData("JAN")
                End If

                '価格
                SearchResult(Count).PRICE = SearchData("PRICE")
                '商品フラグ
                SearchResult(Count).PACKAGE_FLG = SearchData("PACKAGE_FLG")
                '入数
                If IsDBNull(SearchData("IN_BOX_NUM")) Then
                    SearchResult(Count).IN_BOX_NUM = 0
                Else
                    SearchResult(Count).IN_BOX_NUM = SearchData("IN_BOX_NUM")
                End If

                'プロダクトライン名
                SearchResult(Count).PL_NAME = SearchData("PL_NAME")
                SearchResult(Count).PL_ID = SearchData("PL_ID")

                'マスターカートーンサイズ
                If IsDBNull(SearchData("MASTER_CARTON_SIZE")) Then
                    SearchResult(Count).MASTER_CARTON_SIZE = ""
                Else
                    SearchResult(Count).MASTER_CARTON_SIZE = SearchData("MASTER_CARTON_SIZE")
                End If

                'ロケーション
                If IsDBNull(SearchData("LOCATION")) Then
                    SearchResult(Count).LOCATION = ""
                Else
                    SearchResult(Count).LOCATION = SearchData("LOCATION")
                End If

                'カスタマーコード
                If IsDBNull(SearchData("C_CODE")) Then
                    SearchResult(Count).C_CODE = ""
                Else
                    SearchResult(Count).C_CODE = SearchData("C_CODE")
                End If

                'カスタマー名
                If IsDBNull(SearchData("C_NAME")) Then
                    SearchResult(Count).C_NAME = ""
                Else
                    SearchResult(Count).C_NAME = SearchData("C_NAME")
                End If

                'カスタマーコード
                'SearchResult(Count).C_CODE = SearchData("C_CODE")
                'カスタマー名
                'SearchResult(Count).C_NAME = SearchData("C_NAME")
                'C_ID
                SearchResult(Count).C_ID = SearchData("C_ID")
                '仕入金額
                SearchResult(Count).PURCHASE_PRICE = SearchData("PURCHASE_PRICE")
                '免責金額
                SearchResult(Count).IMMUNITY_PRICE = SearchData("IMMUNITY_PRICE")
                '修理金額
                SearchResult(Count).REPAIR_PRICE = SearchData("REPAIR_PRICE")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function


    '***********************************************
    ' 商品マスタを修正する。
    ' <引数>
    ' Upd_List : 修正する商品情報が格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_Mitem(ByRef Upd_List() As M_Item_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Dim C_ID As Integer

        Dim C_ID_DATA As MySqlDataReader


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Upd_List.Length - 1

                If Upd_List(Count).C_CODE <> "" Then

                    'ベンダーコードからIDを取得する。
                    Command = Connection.CreateCommand
                    Command.CommandText = "SELECT ID FROM M_CUSTOMER WHERE C_CODE='"
                    Command.CommandText &= Upd_List(Count).C_CODE
                    Command.CommandText &= "'"

                    C_ID_DATA = Command.ExecuteReader(CommandBehavior.SingleRow)
                    If C_ID_DATA.Read Then
                        ' レコードが取得できた時の処理
                        C_ID = C_ID_DATA("ID")
                    Else
                        ' レコードが取得できなかった時の処理
                        ErrorMessage = "該当データがみつかりません。"
                        Exit Function
                    End If
                    'Close
                    C_ID_DATA.Close()
                End If


                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE M_ITEM SET I_CODE='"
                Command.CommandText &= Upd_List(Count).I_CODE
                Command.CommandText &= "',I_NAME='"
                Command.CommandText &= Upd_List(Count).I_NAME
                Command.CommandText &= "',JAN='"
                Command.CommandText &= Upd_List(Count).JAN
                Command.CommandText &= "',PRICE="
                Command.CommandText &= Upd_List(Count).PRICE
                Command.CommandText &= ",IN_BOX_NUM="
                Command.CommandText &= Upd_List(Count).IN_BOX_NUM
                Command.CommandText &= ",MASTER_CARTON_SIZE='"
                Command.CommandText &= Upd_List(Count).MASTER_CARTON_SIZE
                Command.CommandText &= "',LOCATION='"
                Command.CommandText &= Upd_List(Count).LOCATION
                Command.CommandText &= "',PL_CODE="
                Command.CommandText &= Upd_List(Count).PL_ID
                Command.CommandText &= ",C_ID="
                Command.CommandText &= C_ID
                Command.CommandText &= ",PACKAGE_FLG="
                Command.CommandText &= Upd_List(Count).PACKAGE_FLG

                Command.CommandText &= ",PURCHASE_PRICE="
                Command.CommandText &= Upd_List(Count).PURCHASE_PRICE
                Command.CommandText &= ",IMMUNITY_PRICE="
                Command.CommandText &= Upd_List(Count).IMMUNITY_PRICE
                Command.CommandText &= ",REPAIR_PRICE="
                Command.CommandText &= Upd_List(Count).REPAIR_PRICE

                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Upd_List(Count).ID

                'UPDATE実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業データ取得
    ' <引数>
    ' C_Code : 企業コード
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Get_Mcustomer_Modify(ByRef C_Code As String, _
                                       ByRef SearchResult() As M_Customer_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0
        Dim Wheresql As String = Nothing
        Dim Itemsql As String = Nothing

        Dim checkflg As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_CUSTOMER.ID,M_CUSTOMER.C_CODE,M_CUSTOMER.C_NAME,M_CUSTOMER.SHEET_TYPE,M_CUSTOMER.D_NAME,M_CUSTOMER.D_ZIP,"
            Command.CommandText &= " M_CUSTOMER.D_ADDRESS,M_CUSTOMER.D_TEL,M_CUSTOMER.D_FAX,M_CUSTOMER.CUSTOMER_TYPE,M_CUSTOMER.DELIVERY_PRT_FLG,"
            Command.CommandText &= " M_CUSTOMER.CLAIM_CODE,M_CUSTOMER.DISCOUNT_RATE "
            Command.CommandText &= " FROM M_CUSTOMER "
            Command.CommandText &= " WHERE M_CUSTOMER.C_CODE='"
            Command.CommandText &= C_Code
            Command.CommandText &= "'"

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '企業ID
                SearchResult(Count).ID = SearchData("ID")

                '企業コード
                If IsDBNull(SearchData("C_CODE")) Then
                    SearchResult(Count).C_CODE = ""
                Else
                    SearchResult(Count).C_CODE = SearchData("C_CODE")
                End If

                '企業名
                If IsDBNull(SearchData("C_NAME")) Then
                    SearchResult(Count).C_NAME = ""
                Else
                    SearchResult(Count).C_NAME = SearchData("C_NAME")
                End If

                '伝票タイプ
                SearchResult(Count).SHEET_TYPE = SearchData("SHEET_TYPE")

                '納品先名
                If IsDBNull(SearchData("D_NAME")) Then
                    SearchResult(Count).D_NAME = ""
                Else
                    SearchResult(Count).D_NAME = SearchData("D_NAME")
                End If

                '納品先名郵便番号
                If IsDBNull(SearchData("D_ZIP")) Then
                    SearchResult(Count).D_ZIP = ""
                Else
                    SearchResult(Count).D_ZIP = SearchData("D_ZIP")
                End If
                '納品先名住所
                If IsDBNull(SearchData("D_ADDRESS")) Then
                    SearchResult(Count).D_ADDRESS = ""
                Else
                    SearchResult(Count).D_ADDRESS = SearchData("D_ADDRESS")
                End If
                '納品先名電話番号
                If IsDBNull(SearchData("D_TEL")) Then
                    SearchResult(Count).D_TEL = ""
                Else
                    SearchResult(Count).D_TEL = SearchData("D_TEL")
                End If
                '納品先名FAX
                If IsDBNull(SearchData("D_FAX")) Then
                    SearchResult(Count).D_FAX = ""
                Else
                    SearchResult(Count).D_FAX = SearchData("D_FAX")
                End If

                'カスタマータイプ
                SearchResult(Count).CUSTOMER_TYPE = SearchData("CUSTOMER_TYPE")

                '納品先別出荷リスト出力フラグ
                SearchResult(Count).DELIVERY_PRT_FLG = SearchData("DELIVERY_PRT_FLG")

                '請求先コード
                SearchResult(Count).CLAIM_CODE = SearchData("CLAIM_CODE")

                '掛け率
                SearchResult(Count).DISCOUNT_RATE = SearchData("DISCOUNT_RATE")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 企業マスタを修正する。
    ' <引数>
    ' Upd_List : 修正する企業情報が格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_Mcustomer(ByRef Upd_List() As M_Customer_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Upd_List.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE M_CUSTOMER SET C_CODE='"
                Command.CommandText &= Upd_List(Count).C_CODE
                Command.CommandText &= "',C_NAME='"
                Command.CommandText &= Upd_List(Count).C_NAME
                Command.CommandText &= "',SHEET_TYPE='"
                Command.CommandText &= Upd_List(Count).SHEET_TYPE
                Command.CommandText &= "',D_NAME='"
                Command.CommandText &= Upd_List(Count).D_NAME
                Command.CommandText &= "',D_ZIP='"
                Command.CommandText &= Upd_List(Count).D_ZIP
                Command.CommandText &= "',D_ADDRESS='"
                Command.CommandText &= Upd_List(Count).D_ADDRESS
                Command.CommandText &= "',D_TEL='"
                Command.CommandText &= Upd_List(Count).D_TEL
                Command.CommandText &= "',D_FAX='"
                Command.CommandText &= Upd_List(Count).D_FAX
                Command.CommandText &= "',DELIVERY_PRT_FLG="
                Command.CommandText &= Upd_List(Count).DELIVERY_PRT_FLG
                Command.CommandText &= ",CUSTOMER_TYPE='"
                Command.CommandText &= Upd_List(Count).CUSTOMER_TYPE
                Command.CommandText &= "',CLAIM_CODE='"
                Command.CommandText &= Upd_List(Count).CLAIM_CODE
                Command.CommandText &= "',DISCOUNT_RATE="
                Command.CommandText &= Upd_List(Count).DISCOUNT_RATE
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Upd_List(Count).ID

                'UPDATE実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 入荷出荷バランス　データ取得
    ' <引数>
    ' Date_From : 出庫確定日（From）
    ' Date_To : 出庫確定日（To)
    ' PL_ID : プロダクトライン名
    ' Place : 倉庫（1:八潮、2:東久留米）
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GET_IO_Balance(ByRef Date_From As String, _
                                   ByRef Date_To As String, _
                                   ByRef PL_ID As Integer, _
                                   ByRef Place As Integer, _
                                       ByRef SearchResult() As IO_Balance_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0

        Dim Wheresql As String = Nothing

        Wheresql = "AND OUT_TBL.PLACE_ID="
        Wheresql &= Place
        Wheresql &= " "

        If PL_ID <> 0 Then
            Wheresql &= "AND M_PLINE.ID="
            Wheresql &= PL_ID
        End If

        '出庫年月日From
        If Date_From <> "" Then
            Wheresql &= " AND OUT_TBL.FIX_DATE >= '"
            Wheresql &= Date_From
            Wheresql &= "'"
        End If

        '出庫年月日To
        If Date_To <> "" Then
            Wheresql &= " AND OUT_TBL.FIX_DATE <= '"
            Wheresql &= Date_To
            Wheresql &= "'"
        End If

        'If Wheresql <> "" Then
        '    Wheresql = " WHERE " & Wheresql
        'End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_ITEM.ID AS I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,M_PLINE.NAME AS PL_NAME,SUM(NUM) AS OUT_NUM,"
            Command.CommandText &= " ( SELECT SUM(PO_NUM) - SUM(FIX_NUM) - SUM(CANCEL_NUM) FROM P_ORDER WHERE P_ORDER.I_ID = M_ITEM.ID) AS PO_NUM,"
            Command.CommandText &= " ( SELECT SUM(NUM) FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE IN_DETAIL.I_ID = M_ITEM.ID AND STATUS='入荷予定' AND PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= ") AS IN_NUM,( SELECT SUM(NUM) FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE IN_DETAIL.I_ID = M_ITEM.ID AND STATUS='入荷済み' AND PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= ") AS IN_FIX_NUM, ( SELECT SUM(NUM) FROM STOCK WHERE STOCK.I_ID=M_ITEM.ID AND I_STATUS=1 AND PLACE_ID="
            Command.CommandText &= Place
            Command.CommandText &= " ) as STOCK_NUM "
            Command.CommandText &= " FROM OUT_TBL INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE=M_PLINE.ID"
            Command.CommandText &= " WHERE OUT_TBL.STATUS=4 "
            Command.CommandText &= Wheresql
            Command.CommandText &= " GROUP BY OUT_TBL.I_ID ORDER BY M_ITEM.I_CODE "

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '商品ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                '商品コード
                If IsDBNull(SearchData("I_CODE")) Then
                    SearchResult(Count).I_CODE = ""
                Else
                    SearchResult(Count).I_CODE = SearchData("I_CODE")
                End If
                '商品名
                If IsDBNull(SearchData("I_NAME")) Then
                    SearchResult(Count).I_NAME = ""
                Else
                    SearchResult(Count).I_NAME = SearchData("I_NAME")
                End If

                'JAN
                SearchResult(Count).JAN = SearchData("JAN")

                'OUT_NUM
                If IsDBNull(SearchData("OUT_NUM")) Then
                    SearchResult(Count).OUT_NUM = 0
                Else
                    SearchResult(Count).OUT_NUM = SearchData("OUT_NUM")
                End If
                'IN_NUM
                If IsDBNull(SearchData("IN_NUM")) Then
                    SearchResult(Count).IN_NUM = 0
                Else
                    SearchResult(Count).IN_NUM = SearchData("IN_NUM")
                End If

                'IN_FIX_NUM
                If IsDBNull(SearchData("IN_FIX_NUM")) Then
                    SearchResult(Count).IN_FIX_NUM = 0
                Else
                    SearchResult(Count).IN_FIX_NUM = SearchData("IN_FIX_NUM")
                End If

                'PO_NUM
                If IsDBNull(SearchData("PO_NUM")) Then
                    SearchResult(Count).PO_NUM = 0
                Else
                    SearchResult(Count).PO_NUM = SearchData("PO_NUM")
                End If

                'STOCK_NUM
                If IsDBNull(SearchData("STOCK_NUM")) Then
                    SearchResult(Count).STOCK_NUM = 0
                Else
                    SearchResult(Count).STOCK_NUM = SearchData("STOCK_NUM")
                End If

                'プロダクトライン名
                SearchResult(Count).PL_NAME = SearchData("PL_NAME")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品IDを元に発注データを取得
    ' <引数>
    ' I_ID :
    ' Date_To :
    ' PL_ID :
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetPoItemData(ByRef I_ID As Integer, _
                                       ByRef SearchResult() As IO_Graph_List, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0

        Dim Wheresql As String = Nothing


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT ID,PO_NUM,PO_DATE FROM P_ORDER WHERE I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " GROUP BY P_ORDER.ID"

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '発注ID
                SearchResult(Count).ID = SearchData("ID")

                'PO_NUM
                If IsDBNull(SearchData("PO_NUM")) Then
                    SearchResult(Count).PO_NUM = 0
                Else
                    SearchResult(Count).PO_NUM = SearchData("PO_NUM")
                End If

                'STOCK_NUM
                If IsDBNull(SearchData("PO_DATE")) Then
                    SearchResult(Count).PO_DATE = 0
                Else
                    SearchResult(Count).PO_DATE = SearchData("PO_DATE")
                End If

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' ラベルデータを取得
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' <戻り値>
    ' SearchResult : 検索結果
    ' DataCount : 件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetLabelPrtData(ByRef FILE_NAME As String, _
                                    ByRef SearchResult() As Label_Prt_List, _
                                    ByRef DataCount As Integer, _
                                    ByRef Result As String, _
                                    ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.SHEET_NO,rpad(M_ITEM.I_CODE,7,' ') AS I_CODE,OUT_TBL.NUM FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID AND M_ITEM.PL_CODE=99 "
            Command.CommandText &= " AND OUT_TBL.STATUS=2 AND OUT_TBL.I_STATUS=1 AND OUT_TBL.CATEGORY=1 AND OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            'Command.CommandText &= "' ORDER BY M_ITEM.I_CODE"
            Command.CommandText &= "' ORDER BY rpad(M_ITEM.I_CODE,7,' ')"

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '伝票番号
                SearchResult(Count).SHEET_NO = SearchData("SHEET_NO")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '数量
                SearchResult(Count).NUM = SearchData("NUM")

                Count += 1
            Loop

            DataCount = Count

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' グラフ用データ取得
    ' <引数>
    ' I_ID : 商品ID
    ' DateFrom : 日付での検索範囲のFromを格納
    ' Today : 本日の日付を格納
    ' DateTo : 日付での検索範囲のToを格納
    ' MonthFrom : 日付での検索範囲の年月のFromを格納
    ' MonthToday : 本日の日付を年月で格納
    ' <戻り値>
    ' Stock_Num : 在庫数を格納
    ' Start_Stock_Num : DateFrom時点での在庫数を格納
    ' GRAPH_Data : グラフを表示するためのデータを格納
    ' GRAPH_SummaryData : 在庫履歴欄に表示するサマリーデータを格納
    ' GRAPH_Data_Count : 過去データのデータ数を格納
    ' GRAPH_Future_Data : 未来分のグラフを作成するためのデータを格納
    ' GRAPH_Future_Data_Count : 未来分のデータ数を格納
    ' Label_Prt_List : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetGraphData(ByRef I_ID As Integer, _
                                 ByRef DateFrom As String, _
                                 ByRef Today As String, _
                                 ByRef DateTo As String, _
                                 ByRef MonthFrom As String, _
                                 ByRef MonthToday As String, _
                                 ByRef Stock_Num As Integer, _
                                 ByRef Start_Stock_Num As Integer, _
                                 ByRef PLACE As Integer, _
                                 ByRef GRAPH_Data() As GraphData, _
                                 ByRef GRAPH_SummaryData() As GraphSummaryData, _
                                 ByRef GRAPH_Data_Count As Integer, _
                                 ByRef GRAPH_Future_Data() As GraphData, _
                                 ByRef GRAPH_Future_Data_Count As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim SummaryData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim MonthCount As Integer = 1
        Dim Count As Integer = 0

        Dim Stock_Num_Data As MySqlDataReader
        Dim Stock_Log_Data As MySqlDataReader

        Dim TMP_Stock_Num As Integer
        Dim TMP_Stock_Order_Num As Integer

        Dim DataCount As Integer = 0


        ' 本日の日付を取得・設定
        Dim dtNow As DateTime = DateTime.Now.ToString("yyyy/MM/dd")

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            '現在の在庫数を取得
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT NUM FROM STOCK WHERE PLACE_ID= "
            Command.CommandText &= PLACE
            Command.CommandText &= " AND I_STATUS=1 AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= ";"

            Stock_Num_Data = Command.ExecuteReader(CommandBehavior.SingleRow)
            If Stock_Num_Data.Read Then
                ' レコードが取得できた時の処理
                Stock_Num = Stock_Num_Data("NUM")
            Else
                ' レコードが取得できなかった時の処理
                Stock_Num = 0

            End If
            'Close
            Stock_Num_Data.Close()


            'Fromの日付（今日 - 180日）時点の在庫数を取得する。
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT NUM,I_FLG,PACKAGE_NO FROM STOCK_LOG WHERE PLACE_ID="
            Command.CommandText &= PLACE
            Command.CommandText &= " AND I_STATUS=1 AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`U_DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(`U_DATE`, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' AND I_FLG in ('入荷','棚卸','在庫調整','返品出荷','良品変更','セットばらし','出荷済み');"

            'データ取得
            Stock_Log_Data = Command.ExecuteReader()
            DataCount = 0

            TMP_Stock_Num = Stock_Num
            Do While (Stock_Log_Data.Read)

                If Stock_Log_Data("I_FLG") = "入荷" Then
                    TMP_Stock_Num = TMP_Stock_Num - Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "セット組" Then
                    TMP_Stock_Num = TMP_Stock_Num + Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "棚卸" Then
                    TMP_Stock_Num = TMP_Stock_Num - Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "在庫調整" Then
                    TMP_Stock_Num = TMP_Stock_Num - Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "返品出荷" Then
                    TMP_Stock_Num = TMP_Stock_Num + Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "良品変更" Then
                    TMP_Stock_Num = TMP_Stock_Num + Stock_Log_Data("NUM")
                ElseIf Stock_Log_Data("I_FLG") = "セットばらし" Then
                    If IsDBNull(Stock_Log_Data("PACKAGE_NO")) Then
                        TMP_Stock_Num = TMP_Stock_Num - Stock_Log_Data("NUM")
                    End If
                ElseIf Stock_Log_Data("I_FLG") = "出荷済み" Then
                    TMP_Stock_Num = TMP_Stock_Num + Stock_Log_Data("NUM")
                End If
            Loop
            'Close
            Stock_Log_Data.Close()

            Start_Stock_Num = TMP_Stock_Num

            '過去データを取得する（入庫確定数、出庫確定数、在庫調整数、棚卸数、セット組数、セットばらし数）
            Command = Connection.CreateCommand
            Command.CommandText = "(SELECT DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d') AS DATE_DATA,SUM(FIX_NUM) AS NUM,'入庫確定' AS TYPE FROM IN_DETAIL "
            Command.CommandText &= " WHERE I_STATUS=1 AND STATUS='入荷済み' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(FIX_DATE, '%Y/%m/%d') AS DATE_DATA,SUM(FIX_NUM) AS NUM,'出庫確定' AS TYPE FROM OUT_TBL "
            Command.CommandText &= " WHERE STATUS=4 AND I_ID= "
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`FIX_DATE`, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m/%d') AS DATE_DATA,SUM(NUM) AS NUM,'在庫調整' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='在庫調整' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m/%d') AS DATE_DATA,SUM(NUM) AS NUM,'棚卸' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='棚卸' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m/%d') AS DATE_DATA,SUM(NUM) AS NUM,'セット組' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='セット組' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m/%d') AS DATE_DATA,SUM(NUM) AS NUM,'セットばらし' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='セットばらし' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m/%d') <='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m/%d'))"
            Command.CommandText &= "ORDER BY DATE_DATA DESC;"

            'データ取得
            SearchData = Command.ExecuteReader()
            DataCount = 0

            TMP_Stock_Num = Stock_Num
            Do While (SearchData.Read)
                ReDim Preserve GRAPH_Data(0 To Count)
                '日付
                GRAPH_Data(Count).DATE_DATA = SearchData("DATE_DATA")
                '数量
                GRAPH_Data(Count).NUM = SearchData("NUM")
                '種別
                GRAPH_Data(Count).TYPE = SearchData("TYPE")

                If SearchData("TYPE") = "入庫確定" Then
                    TMP_Stock_Num = TMP_Stock_Num - SearchData("NUM")
                ElseIf SearchData("TYPE") = "出庫確定" Then
                    TMP_Stock_Num = TMP_Stock_Num + SearchData("NUM")
                ElseIf SearchData("TYPE") = "棚卸" Then
                    TMP_Stock_Num = TMP_Stock_Num + SearchData("NUM")
                ElseIf SearchData("TYPE") = "在庫調整" Then
                    TMP_Stock_Num = TMP_Stock_Num - SearchData("NUM")
                ElseIf SearchData("TYPE") = "セット組" Then
                    TMP_Stock_Num = TMP_Stock_Num - SearchData("NUM")
                ElseIf SearchData("TYPE") = "セットばらし" Then
                    TMP_Stock_Num = TMP_Stock_Num - SearchData("NUM")
                End If

                '在庫数
                GRAPH_Data(Count).STOCK_NUM = TMP_Stock_Num
                Count += 1
            Loop

            GRAPH_Data_Count = Count

            SearchData.Close()

            'サマリーデータを取得
            Command = Connection.CreateCommand
            Command.CommandText = "(SELECT DATE_FORMAT(`FIX_DATE`, '%Y/%m') AS DATE_DATA,SUM(FIX_NUM) AS NUM,'入庫確定' AS TYPE FROM IN_DETAIL "
            Command.CommandText &= " WHERE I_STATUS=1 AND STATUS='入荷済み' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`FIX_DATE`, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(`FIX_DATE`, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`FIX_DATE`, '%Y/%m'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(FIX_DATE, '%Y/%m') AS DATE_DATA,SUM(FIX_NUM) AS NUM,'出庫確定' AS TYPE FROM OUT_TBL "
            Command.CommandText &= " WHERE STATUS=4 AND I_ID= "
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`FIX_DATE`, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(`FIX_DATE`, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`FIX_DATE`, '%Y/%m'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m') AS DATE_DATA,SUM(NUM) AS NUM,'在庫調整' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='在庫調整' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m') AS DATE_DATA,SUM(NUM) AS NUM,'棚卸' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='棚卸' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m') AS DATE_DATA,SUM(NUM) AS NUM,'セット組' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='セット組' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(U_DATE, '%Y/%m') AS DATE_DATA,SUM(NUM) AS NUM,'セットばらし' AS TYPE FROM STOCK_LOG "
            Command.CommandText &= " WHERE I_FLG='セットばらし' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(U_DATE, '%Y/%m') >='"
            Command.CommandText &= MonthFrom
            Command.CommandText &= "' AND DATE_FORMAT(U_DATE, '%Y/%m') <='"
            Command.CommandText &= MonthToday
            Command.CommandText &= "' GROUP BY DATE_FORMAT(U_DATE, '%Y/%m'))"
            Command.CommandText &= "ORDER BY DATE_DATA DESC;"

            'データ取得
            SummaryData = Command.ExecuteReader()
            DataCount = 0

            TMP_Stock_Num = Stock_Num
            Count = 0
            'Do While (SummaryData.Read)
            '    ReDim Preserve GRAPH_SummaryData(0 To Count)
            '    '日付
            '    GRAPH_SummaryData(Count).DATE_DATA = SummaryData("DATE_DATA")
            '    '数量
            '    GRAPH_SummaryData(Count).NUM = SummaryData("NUM")
            '    '種別
            '    GRAPH_SummaryData(Count).TYPE = SummaryData("TYPE")

            '    If SummaryData("TYPE") = "入庫確定" Then
            '        TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
            '    ElseIf SummaryData("TYPE") = "出庫確定" Then
            '        TMP_Stock_Num = TMP_Stock_Num + SummaryData("NUM")
            '    ElseIf SummaryData("TYPE") = "棚卸" Then
            '        TMP_Stock_Num = TMP_Stock_Num + SummaryData("NUM")
            '    ElseIf SummaryData("TYPE") = "在庫調整" Then
            '        TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
            '    ElseIf SummaryData("TYPE") = "セット組" Then
            '        TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
            '    ElseIf SummaryData("TYPE") = "セットばらし" Then
            '        TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
            '    End If

            '    '在庫数
            '    GRAPH_SummaryData(Count).STOCK_NUM = TMP_Stock_Num
            '    Count += 1
            'Loop

            Do While (SummaryData.Read)
                For i = 0 To 35
                    If GRAPH_SummaryData(i).DATE_DATA = SummaryData("DATE_DATA") And _
                         GRAPH_SummaryData(i).TYPE = SummaryData("TYPE") Then
                        GRAPH_SummaryData(i).NUM = SummaryData("NUM")
                    Else
                        ' GRAPH_SummaryData(i).NUM = 0
                    End If

                Next

                ''日付
                'GRAPH_SummaryData(Count).DATE_DATA = SummaryData("DATE_DATA")
                ''数量
                'GRAPH_SummaryData(Count).NUM = SummaryData("NUM")
                ''種別
                'GRAPH_SummaryData(Count).TYPE = SummaryData("TYPE")

                'If SummaryData("TYPE") = "入庫確定" Then
                '    TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
                'ElseIf SummaryData("TYPE") = "出庫確定" Then
                '    TMP_Stock_Num = TMP_Stock_Num + SummaryData("NUM")
                'ElseIf SummaryData("TYPE") = "棚卸" Then
                '    TMP_Stock_Num = TMP_Stock_Num + SummaryData("NUM")
                'ElseIf SummaryData("TYPE") = "在庫調整" Then
                '    TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
                'ElseIf SummaryData("TYPE") = "セット組" Then
                '    TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
                'ElseIf SummaryData("TYPE") = "セットばらし" Then
                '    TMP_Stock_Num = TMP_Stock_Num - SummaryData("NUM")
                'End If

                '在庫数
                GRAPH_SummaryData(Count).STOCK_NUM = TMP_Stock_Num
                Count += 1
            Loop

            GRAPH_Data_Count = Count

            SummaryData.Close()

            '未来データを取得(入庫予定、出庫予定）
            '未来分は全データ取得
            '未来分のデータと共に、過去のデータで入庫予定、出庫予定となったままのデータも取得（過去半年分）
            Command = Connection.CreateCommand
            Command.CommandText = "(SELECT DATE_FORMAT(`DATE`, '%Y/%m/%d') AS DATE_DATA, SUM(NUM) AS NUM,'入庫予定' AS TYPE,'' AS ORDER_DATE FROM IN_DETAIL "
            Command.CommandText &= " WHERE I_STATUS=1 AND STATUS='入荷予定' AND I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`DATE`, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(DATE, '%Y/%m/%d') AS DATE_DATA,SUM(NUM) AS NUM ,'出庫予定' AS TYPE,'' AS ORDER_DATE FROM OUT_TBL "
            Command.CommandText &= " WHERE STATUS='出荷予定' AND I_ID= "
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= DateFrom
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`DATE`, '%Y/%m/%d'))"
            Command.CommandText &= "UNION "
            Command.CommandText &= "(SELECT DATE_FORMAT(`PO_DATE`, '%Y/%m/%d') AS DATE_DATA,SUM(PO_NUM) - SUM(FIX_NUM) - SUM(CANCEL_NUM) AS NUM, "
            Command.CommandText &= "'発注' AS TYPE,DATE_FORMAT(`ORDER_DATE`, '%Y/%m/%d') AS ORDER_DATE FROM P_ORDER WHERE  I_ID="
            Command.CommandText &= I_ID
            Command.CommandText &= " AND DATE_FORMAT(`PO_DATE`, '%Y/%m/%d') >='"
            Command.CommandText &= Today
            Command.CommandText &= "' GROUP BY DATE_FORMAT(`PO_DATE`, '%Y/%m/%d'))"
            Command.CommandText &= " ORDER BY DATE_DATA;"


            'データ取得
            SearchData = Command.ExecuteReader()
            Count = 0
            DataCount = 0
            TMP_Stock_Num = Stock_Num
            TMP_Stock_Order_Num = Stock_Num
            Do While (SearchData.Read)
                ReDim Preserve GRAPH_Future_Data(0 To Count)
                '日付
                GRAPH_Future_Data(Count).DATE_DATA = SearchData("DATE_DATA")
                '日付
                GRAPH_Future_Data(Count).ORDER_DATE = SearchData("ORDER_DATE")
                '数量
                GRAPH_Future_Data(Count).NUM = SearchData("NUM")
                '種別
                GRAPH_Future_Data(Count).TYPE = SearchData("TYPE")

                If SearchData("TYPE") = "入庫予定" Then

                    If Today <= SearchData("DATE_DATA") Then

                        TMP_Stock_Num = TMP_Stock_Num + SearchData("NUM")
                        TMP_Stock_Order_Num = TMP_Stock_Order_Num + SearchData("NUM")
                    End If
                ElseIf SearchData("TYPE") = "出庫予定" Then
                    If Today <= SearchData("DATE_DATA") Then
                        TMP_Stock_Num = TMP_Stock_Num - SearchData("NUM")
                        TMP_Stock_Order_Num = TMP_Stock_Order_Num - SearchData("NUM")
                    End If
                ElseIf SearchData("TYPE") = "発注" Then
                    TMP_Stock_Order_Num = TMP_Stock_Order_Num + SearchData("NUM")
                End If

                '在庫数
                GRAPH_Future_Data(Count).STOCK_NUM = TMP_Stock_Num
                '在庫数＋発注数
                GRAPH_Future_Data(Count).STOCK_ORDER_NUM = TMP_Stock_Order_Num
                Count += 1
            Loop

            GRAPH_Future_Data_Count = Count

            SearchData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 発注予測検索
    ' <引数>
    ' I_CODE : 商品コード
    ' Standard_Num : 基準値
    ' Standard_Day : 基準日
    ' PL_ID : プロダクトラインID
    ' PLACE_ID : 倉庫ID
    ' <戻り値>
    ' SearchResult : 結果格納用配列 
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GET_PO_Prediction(ByVal I_CODE As String, _
                                ByVal Standard_Num As String, _
                                ByVal Standard_Day As String, _
                                ByVal PL_ID As Integer, _
                                ByVal PLACE_ID As Integer, _
                                ByRef SearchResult() As PO_Prediction_List, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        '発注予測データの結果を格納
        Dim Search_ListData As MySqlDataReader

        Dim Connection As New MySqlConnection
        Dim Connection2 As New MySqlConnection

        Dim Command As MySqlCommand = Nothing
        Dim Command2 As MySqlCommand = Nothing

        Dim Count As Integer = 0
        Dim Theory_Data1 As Integer = 0
        Dim Theory_Data2 As Integer = 0

        '商品コード、プロダクトライン検索条件を格納
        Dim WhereSQL As String = Nothing

        '入庫予定日　検索条件を格納（基準日在庫取得用）
        Dim IN_SQL As String = Nothing
        '出庫予定日　検索条件を格納（基準日在庫取得用）
        Dim OUT_SQL As String = Nothing
        '入庫予定日　検索条件（基準日以降取得用）
        Dim IN_PLAN_SQL As String = Nothing
        '出庫予定日　検索条件（基準日以降取得用）
        Dim OUT_PLAN_SQL As String = Nothing

        '今日の日付を取得
        Dim Today As String = DateTime.Now.ToString("yyyy/MM/dd")

        '日付差を計算するためにDate型で今日の日付を取得
        'Dim Today2 As Date = DateTime.Now.ToString("yyyy/MM/dd")
        '半年前の日付を取得する
        'Dim PastDay As String = DateTime.Now.AddDays(-180).ToString("yyyy/MM/dd")

        '基準日と今日との日付差を計算する。
        'Dim DateDiff As Integer
        'If TheoryDay <> "" Then
        '    DateDiff = DateTime.ParseExact(TheoryDay, "yyyy/MM/dd", Nothing).Subtract(Today2).Days
        'Else
        '    DateDiff = 0
        'End If

        '商品コードが設定されていれば、検索条件に追加
        If I_CODE <> "" Then
            WhereSQL = "AND M_ITEM.I_CODE ='"
            WhereSQL &= I_CODE
            WhereSQL &= "' "
        End If

        '基準日が入力されていれば、検索条件に追加
        If Standard_Day <> "" Then
            IN_SQL &= " AND IN_DETAIL.`DATE`<='"
            IN_SQL &= Standard_Day
            IN_SQL &= "'"

            OUT_SQL = "AND OUT_TBL.`DATE`<='"
            OUT_SQL &= Standard_Day
            OUT_SQL &= "'"

            IN_PLAN_SQL = "AND IN_DETAIL.`DATE` > '"
            IN_PLAN_SQL &= Standard_Day
            IN_PLAN_SQL &= "'"

            OUT_PLAN_SQL = "AND OUT_TBL.`DATE` > '"
            OUT_PLAN_SQL &= Standard_Day
            OUT_PLAN_SQL &= "'"
        End If

        'プロダクトラインIDが0じゃなければ、検索条件に追加
        If PL_ID <> 0 Then
            WhereSQL &= " AND M_PLINE.ID = "
            WhereSQL &= PL_ID
            WhereSQL &= " "
        End If

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_ITEM.ID, M_ITEM.I_CODE, M_ITEM.I_NAME,M_ITEM.JAN,M_PLINE.NAME AS PL_NAME,STOCK.NUM AS STOCK_NUM,"
            '基準日以前の入庫予定数を取得
            Command.CommandText &= " COALESCE((SELECT SUM(IN_DETAIL.NUM) FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE "
            Command.CommandText &= " IN_HEADER.PLACE_ID="
            Command.CommandText &= PLACE_ID
            Command.CommandText &= "  AND IN_DETAIL.I_ID=M_ITEM.ID AND IN_DETAIL.STATUS='入荷予定' AND IN_DETAIL.I_STATUS='良品' "
            Command.CommandText &= IN_SQL
            Command.CommandText &= " ),0) AS IN_NUM,"
            '基準日以前の出庫予定数を取得
            Command.CommandText &= " COALESCE((SELECT SUM(OUT_TBL.NUM) FROM OUT_TBL WHERE OUT_TBL.I_ID= M_ITEM.ID AND OUT_TBL.PLACE_ID ="
            Command.CommandText &= PLACE_ID
            Command.CommandText &= " AND OUT_TBL.STATUS='出荷予定' AND OUT_TBL.I_STATUS='良品' "
            Command.CommandText &= OUT_SQL
            Command.CommandText &= " ),0) AS OUT_NUM, "
            '過去180日の入庫確定数を取得
            'Command.CommandText &= "ROUND( (SELECT (SUM(IN_DETAIL.FIX_NUM)/180) * "
            'Command.CommandText &= DateDiff
            'Command.CommandText &= "   FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE"
            'Command.CommandText &= " IN_HEADER.PLACE='八潮' AND IN_DETAIL.I_ID=M_ITEM.ID AND IN_DETAIL.STATUS='入荷済み' AND IN_DETAIL.I_STATUS='良品' AND "
            'Command.CommandText &= " IN_DETAIL.FIX_DATE <='"
            'Command.CommandText &= Today
            'Command.CommandText &= "' AND IN_DETAIL.FIX_DATE >='"
            'Command.CommandText &= PastDay
            'Command.CommandText &= "' "
            'Command.CommandText &= "),0) AS IN_PREDICTION_NUM  ,"

            '過去180日の出庫確定数を取得
            'Command.CommandText &= " ROUND((SELECT (SUM(OUT_TBL.FIX_NUM)/180) * "
            'Command.CommandText &= DateDiff
            'Command.CommandText &= "  FROM OUT_TBL WHERE "
            'Command.CommandText &= " OUT_TBL.PLACE='八潮' AND OUT_TBL.I_ID=M_ITEM.ID AND OUT_TBL.STATUS='出荷済み' AND OUT_TBL.I_STATUS='良品' AND "
            'Command.CommandText &= " OUT_TBL.FIX_DATE <='"
            'Command.CommandText &= Today
            'Command.CommandText &= "' AND OUT_TBL.FIX_DATE >='"
            'Command.CommandText &= PastDay
            'Command.CommandText &= "' "
            'Command.CommandText &= " ),0) AS OUT_PREDICTION_NUM ,"

            '商品マスタの基準値 + 入力した基準値
            Command.CommandText &= "(M_ITEM.STANDARD_NUM + "
            Command.CommandText &= Standard_Num
            Command.CommandText &= ") AS STANDARD_NUM,"
            '在庫数+入庫予定数-出庫予定数
            '発注残数を取得
            Command.CommandText &= " (SELECT SUM(P_ORDER.PO_NUM) - SUM(P_ORDER.FIX_NUM) - SUM(P_ORDER.CANCEL_NUM) AS PO_NUM FROM P_ORDER WHERE M_ITEM.ID=P_ORDER.I_ID ) AS PO_NUM,"
            '基準日以降の入庫予定数を取得
            Command.CommandText &= " (SELECT SUM(IN_DETAIL.NUM) FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE "
            Command.CommandText &= " IN_HEADER.PLACE_ID="
            Command.CommandText &= PLACE_ID
            Command.CommandText &= " AND IN_DETAIL.I_ID=M_ITEM.ID AND IN_DETAIL.STATUS='入荷予定' AND IN_DETAIL.I_STATUS='良品' "
            Command.CommandText &= IN_PLAN_SQL
            Command.CommandText &= " ) AS IN_PLAN_NUM,"
            '基準日以降の出庫予定数を取得
            Command.CommandText &= " (SELECT SUM(OUT_TBL.NUM) FROM OUT_TBL WHERE OUT_TBL.I_ID= M_ITEM.ID AND OUT_TBL.PLACE_ID="
            Command.CommandText &= PLACE_ID
            Command.CommandText &= " AND OUT_TBL.STATUS='出荷予定' AND OUT_TBL.I_STATUS='良品' "
            Command.CommandText &= OUT_PLAN_SQL
            Command.CommandText &= " ) AS OUT_PLAN_NUM "
            Command.CommandText &= " FROM STOCK INNER JOIN M_ITEM ON  M_ITEM.ID = STOCK.I_ID "
            Command.CommandText &= " INNER JOIN M_PLINE ON M_ITEM.PL_CODE= M_PLINE.ID"
            Command.CommandText &= " WHERE STOCK.I_STATUS = '良品' AND STOCK.PLACE_ID = "
            Command.CommandText &= PLACE_ID & " "
            Command.CommandText &= WhereSQL
            Command.CommandText &= " AND (STOCK.NUM + COALESCE((SELECT SUM(IN_DETAIL.NUM) FROM IN_DETAIL INNER JOIN IN_HEADER ON IN_HEADER.ID=IN_DETAIL.ID WHERE  IN_HEADER.PLACE_ID ="
            Command.CommandText &= PLACE_ID
            Command.CommandText &= " AND IN_DETAIL.I_ID=M_ITEM.ID AND IN_DETAIL.STATUS='入荷予定' AND IN_DETAIL.I_STATUS='良品'  ),0)"

            Command.CommandText &= "- COALESCE((SELECT SUM(OUT_TBL.NUM) FROM OUT_TBL WHERE OUT_TBL.I_ID= M_ITEM.ID AND OUT_TBL.STATUS='出荷予定' AND OUT_TBL.I_STATUS='良品') ,0 )) <"
            Command.CommandText &= "(M_ITEM.STANDARD_NUM + "
            Command.CommandText &= Standard_Num
            Command.CommandText &= ") "
            Command.CommandText &= " GROUP BY M_ITEM.ID"
            Command.CommandText &= " ORDER BY M_ITEM.ID"

            'データ取得
            Search_ListData = Command.ExecuteReader

            Do While (Search_ListData.Read)
                ReDim Preserve SearchResult(0 To Count)
                SearchResult(Count).ID = Search_ListData("ID")
                SearchResult(Count).I_CODE = Search_ListData("I_CODE")
                SearchResult(Count).I_NAME = Search_ListData("I_NAME")
                SearchResult(Count).JAN = Search_ListData("JAN")
                SearchResult(Count).PL_NAME = Search_ListData("PL_NAME")

                If IsDBNull(Search_ListData("STOCK_NUM")) Then
                    SearchResult(Count).STOCK_NUM = 0
                Else
                    SearchResult(Count).STOCK_NUM = Search_ListData("STOCK_NUM")
                End If

                If IsDBNull(Search_ListData("IN_NUM")) Then
                    SearchResult(Count).IN_NUM = 0
                Else
                    SearchResult(Count).IN_NUM = Search_ListData("IN_NUM")
                End If

                If IsDBNull(Search_ListData("OUT_NUM")) Then
                    SearchResult(Count).OUT_NUM = 0
                Else
                    SearchResult(Count).OUT_NUM = Search_ListData("OUT_NUM")
                End If

                SearchResult(Count).STANDARD_NUM = Search_ListData("STANDARD_NUM")

                If IsDBNull(Search_ListData("PO_NUM")) Then
                    SearchResult(Count).PO_NUM = 0
                Else
                    SearchResult(Count).PO_NUM = Search_ListData("PO_NUM")
                End If

                If IsDBNull(Search_ListData("IN_PLAN_NUM")) Then
                    SearchResult(Count).IN_PLAN_NUM = 0
                Else
                    SearchResult(Count).IN_PLAN_NUM = Search_ListData("IN_PLAN_NUM")
                End If

                If IsDBNull(Search_ListData("OUT_PLAN_NUM")) Then
                    SearchResult(Count).OUT_PLAN_NUM = 0
                Else
                    SearchResult(Count).OUT_PLAN_NUM = Search_ListData("OUT_PLAN_NUM")
                End If

                'If IsDBNull(Search_ListData("IN_PREDICTION_NUM")) Then
                '    SearchResult(Count).IN_PREDICTION_NUM = 0
                'Else
                '    SearchResult(Count).IN_PREDICTION_NUM = Search_ListData("IN_PREDICTION_NUM")
                'End If

                'If IsDBNull(Search_ListData("OUT_PREDICTION_NUM")) Then
                '    SearchResult(Count).OUT_PREDICTION_NUM = 0
                'Else
                '    SearchResult(Count).OUT_PREDICTION_NUM = Search_ListData("OUT_PREDICTION_NUM")
                'End If

                Count += 1
            Loop

            '0件ならエラー
            If Count = 0 Then

                ErrorMessage = "データがみつかりません。"

                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品マスタの基準値を修正する。
    ' <引数>
    ' Standart_modify_Data : 修正する商品ID、基準値が格納された配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_StandardNum(ByRef Upd_List() As Standardnum_Import_List, _
                             ByRef Result As String, _
                             ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer


        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Upd_List.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE M_ITEM SET STANDARD_NUM="
                Command.CommandText &= Upd_List(Count).STANDARD_NUM
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Upd_List(Count).I_ID

                'UPDATE実行
                Command.ExecuteNonQuery()
            Next

            '全SQLの発行が完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' TMP_M_ITEMテーブルを作成する。
    ' <引数>
    ' 
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function CREATE_TMP_M_ITEM(ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean


        Dim TableCheck As MySqlDataReader
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SHOW TABLES FROM pfj like 'TMP_M_ITEM'"
            'Select実行
            TableCheck = Command.ExecuteReader(CommandBehavior.SingleRow)

            If TableCheck.Read Then
                TableCheck.Close()
                'コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " DROP TABLE TMP_M_ITEM;"
                '実行
                Command.ExecuteNonQuery()
            Else
                TableCheck.Close()
            End If


            'テーブル用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " CREATE TABLE TMP_M_ITEM("
            Command.CommandText &= " I_CODE			VARCHAR(24) NOT NULL		COMMENT '商品コード',"
            Command.CommandText &= " STANDARD_NUM			INTEGER NOT NULL		COMMENT '基準値',"
            Command.CommandText &= " INDEX INDEX_1 (I_CODE)"
            Command.CommandText &= " ) TYPE = InnoDB, ENGINE = InnoDB, ROW_FORMAT = COMPACT, DEFAULT CHARSET=utf8, CHARACTER SET UTF8;"
            '実行
            Command.ExecuteNonQuery()

            'コミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品マスタ基準値データ（エクセルデータ）をTMP_M_ITEMに登録。
    ' <引数>
    ' ExcelData : 基準値データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_TMP_M_ITEM(ByRef ExcelData() As Standardnum_Import_List, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean


        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To ExcelData.Length - 1

                'TMP_M_ITEMテーブル用コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = " INSERT INTO TMP_M_ITEM(I_CODE,STANDARD_NUM)VALUES('"
                Command.CommandText &= ExcelData(Count).I_CODE
                Command.CommandText &= "',"
                Command.CommandText &= ExcelData(Count).STANDARD_NUM
                Command.CommandText &= ");"

                'TMP_M_ITEMテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next

            '全てのデータをTMP_M_ITEMテーブルにINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 商品基準値データ取込後のTMPテーブルからデータ取得
    ' <引数>
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GET_TMP_M_ITEM(ByRef SearchResult() As Standardnum_Import_List, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean
        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_ITEM.ID AS I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,TMP_M_ITEM.STANDARD_NUM,M_PLINE.NAME AS PL_NAME "
            Command.CommandText &= " FROM (TMP_M_ITEM) LEFT JOIN M_ITEM ON TMP_M_ITEM.I_CODE = M_ITEM.I_CODE "
            Command.CommandText &= " LEFT JOIN M_PLINE ON M_ITEM.PL_CODE = M_PLINE.ID; "

            'データ取得
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)

                SearchResult(Count).I_ID = SearchData("I_ID")
                If SearchData("I_ID") = 0 Then

                End If

                If IsDBNull(SearchData("I_CODE")) Then
                    SearchResult(Count).I_CODE = 0
                Else
                    SearchResult(Count).I_CODE = SearchData("I_CODE")
                End If
                If IsDBNull(SearchData("I_NAME")) Then
                    SearchResult(Count).I_NAME = ""
                Else
                    SearchResult(Count).I_NAME = SearchData("I_NAME")
                End If

                If IsDBNull(SearchData("JAN")) Then
                    SearchResult(Count).JAN = 0
                Else
                    SearchResult(Count).JAN = SearchData("JAN")
                End If

                SearchResult(Count).STANDARD_NUM = SearchData("STANDARD_NUM")

                SearchResult(Count).PL_NAME = SearchData("PL_NAME")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' TMP_M_ITEMテーブルを削除する。
    ' <引数>
    ' 
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function DROP_TMP_M_ITEM(ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " DROP TABLE TMP_M_ITEM;"
            '実行
            Command.ExecuteNonQuery()

            'コミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 納品書リストのあて先一覧を取得する
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' <戻り値>
    ' Delivery_List_Data : チェックリストデータ格納
    ' Total : 総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDliveryData(ByVal FILE_NAME As String, _
                                        ByRef Delivery_List_Data() As Delivery_HeaderList, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean
        Dim DeliveryGroupListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing

        Dim DataCount As Integer = 0

        Dim CID_Check_List() As CheckDeliveryID_List = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,OUT_TBL.SHEET_NO,OUT_TBL.DATE AS OUT_DATE,OUT_TBL.FIX_DATE,OUT_TBL.ORDER_NO,"
            Command.CommandText &= " M_CUSTOMER.C_NAME,M_CUSTOMER.D_NAME,OUT_TBL.COMMENT1,OUT_TBL.COMMENT2,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME,M_CUSTOMER.claim_CODE "
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' GROUP BY OUT_TBL.C_ID,OUT_TBL.SHEET_NO;"

            'データ取得
            DeliveryGroupListData = Command.ExecuteReader()

            Do While (DeliveryGroupListData.Read)
                ReDim Preserve Delivery_List_Data(0 To DataCount)
                Delivery_List_Data(DataCount).C_ID = DeliveryGroupListData("C_ID")
                Delivery_List_Data(DataCount).SHEET_NO = DeliveryGroupListData("SHEET_NO")

                If IsDBNull(DeliveryGroupListData("FIX_DATE")) Then
                    Delivery_List_Data(DataCount).FIX_DATE = ""
                Else
                    Delivery_List_Data(DataCount).FIX_DATE = DeliveryGroupListData("FIX_DATE")
                End If

                Delivery_List_Data(DataCount).OUT_DATE = DeliveryGroupListData("OUT_DATE")

                Delivery_List_Data(DataCount).C_NAME = DeliveryGroupListData("C_NAME")
                Delivery_List_Data(DataCount).D_NAME = DeliveryGroupListData("D_NAME")

                If IsDBNull(DeliveryGroupListData("COMMENT1")) Then
                    Delivery_List_Data(DataCount).COMMENT1 = ""
                Else
                    Delivery_List_Data(DataCount).COMMENT1 = DeliveryGroupListData("COMMENT1")
                End If

                If IsDBNull(DeliveryGroupListData("COMMENT2")) Then
                    Delivery_List_Data(DataCount).COMMENT2 = ""
                Else
                    Delivery_List_Data(DataCount).COMMENT2 = DeliveryGroupListData("COMMENT2")
                End If

                If IsDBNull(DeliveryGroupListData("C_CODE")) Then
                    Delivery_List_Data(DataCount).C_CODE = ""
                Else
                    Delivery_List_Data(DataCount).C_CODE = DeliveryGroupListData("C_CODE")
                End If

                Delivery_List_Data(DataCount).ORDER_NO = DeliveryGroupListData("ORDER_NO")

                DataCount += 1
            Loop

            '値を取得できたのでClose
            DeliveryGroupListData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 納品書リストの各詳細一覧を取得する
    ' <引数>
    ' FILE_NAME : 出荷指示ファイル名
    ' C_ID : 納品先ID
    ' SHEET_NO : 伝票番号

    ' <戻り値>
    ' Check_List_Data : チェックリストデータ格納
    ' Total : 総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetDliveryDetailData(ByVal FILE_NAME As String, _
                                         ByVal C_ID As Integer, _
                                         ByVal SHEET_NO As String, _
                                        ByRef DeliveryDetailList() As Delivery_DetailList, _
                                        ByRef TotalAmount As Integer, _
                                        ByRef Result As String, _
                                        ByRef ErrorMessage As String) As Boolean

        Dim DeliveryListData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim i As Integer = 0

        Dim OrderNo_Loop As Integer = 0
        Dim DataCount As Integer = 0
        Dim DeliveryDataCount As Integer = 0
        Dim PageCheck As Integer = 0
        Dim DataTotalCount As Integer = 0
        Dim Delivery_DetailList() As Delivery_DetailList = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()

            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,M_CUSTOMER.C_CODE,"
            Command.CommandText &= "OUT_TBL.SHEET_NO,M_CUSTOMER.D_NAME,OUT_TBL.FILE_NAME,"
            Command.CommandText &= "M_ITEM.I_CODE,M_ITEM.I_NAME,M_ITEM.JAN,OUT_TBL.NUM,OUT_TBL.UNIT_COST,M_ITEM.PRICE"
            Command.CommandText &= " FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID"
            Command.CommandText &= " INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID"
            Command.CommandText &= " WHERE OUT_TBL.FILE_NAME='"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "' AND OUT_TBL.C_ID="
            Command.CommandText &= C_ID
            Command.CommandText &= " AND OUT_TBL.SHEET_NO='"
            Command.CommandText &= SHEET_NO
            Command.CommandText &= "';"

            'SQL実行
            DeliveryListData = Command.ExecuteReader()

            Do While (DeliveryListData.Read)
                ReDim Preserve DeliveryDetailList(0 To i)
                '商品コード
                DeliveryDetailList(i).I_CODE = DeliveryListData("I_CODE")
                '商品名
                DeliveryDetailList(i).I_NAME = DeliveryListData("I_NAME")
                'JAN
                DeliveryDetailList(i).JAN = DeliveryListData("JAN")
                '数量
                DeliveryDetailList(i).NUM = DeliveryListData("NUM")
                '納入単価
                DeliveryDetailList(i).UNIT_PRICE = DeliveryListData("UNIT_COST")
                '参考上代
                DeliveryDetailList(i).REFERENCE_PRICE = DeliveryListData("PRICE")

                '帳票用に合計金額を算出
                TotalAmount += DeliveryDetailList(i).NUM * DeliveryDetailList(i).UNIT_PRICE
                i += 1

            Loop
            '値を取得できたのでClose
            DeliveryListData.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' ロット番号を登録・更新
    ' <引数>
    ' LotData : 基準値データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Out_LotRegist(ByVal LotData() As Lot_List, _
                                  ByRef Result As String, _
                                  ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim OUTData As MySqlDataReader
        Dim OUTCount As Integer = 0
        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'OUT_IDを使用し、すでに登録されているかチェック。
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM OUT_LOT_MANAGEMENT WHERE OUT_ID="
            Command.CommandText &= LotData(0).OUT_ID
            Command.CommandText &= ";"

            'データ取得
            OUTData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If OUTData.Read Then
                ' レコードが取得できた時の処理
                OUTCount = OUTData("COUNT")
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "ロット情報取得中にエラーが発生しました。"
                Exit Function
            End If

            OUTData.Close()

            'OUTCountに件数があればデータが存在しているので更新
            If OUTCount <> 0 Then
                For Count = 0 To LotData.Length - 1

                    'OUT_LOT_MANAGEMENTテーブル用コマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " UPDATE OUT_LOT_MANAGEMENT SET LOT_NUMBER='"
                    Command.CommandText &= LotData(Count).LOT_NUMBER
                    Command.CommandText &= "',WARRANTY_CARD_NUMBER='"
                    Command.CommandText &= LotData(Count).WARRANTY_CARD_NUMBER
                    Command.CommandText &= "',U_DATE=current_timestamp WHERE OUT_ID="
                    Command.CommandText &= LotData(Count).OUT_ID
                    Command.CommandText &= " AND NO="
                    Command.CommandText &= LotData(Count).NO
                    Command.CommandText &= ";"

                    'OUT_LOT_MANAGEMENTテーブルへデータ登録
                    Command.ExecuteNonQuery()
                Next
            Else
                For Count = 0 To LotData.Length - 1

                    'OUT_LOT_MANAGEMENTテーブル用コマンド作成
                    Command = Connection.CreateCommand
                    Command.CommandText = " INSERT INTO OUT_LOT_MANAGEMENT(OUT_ID,NO,LOT_NUMBER,WARRANTY_CARD_NUMBER,R_USER,U_DATE)VALUES("
                    Command.CommandText &= LotData(Count).OUT_ID
                    Command.CommandText &= ","
                    Command.CommandText &= LotData(Count).NO
                    Command.CommandText &= ",'"
                    Command.CommandText &= LotData(Count).LOT_NUMBER
                    Command.CommandText &= "','"
                    Command.CommandText &= LotData(Count).WARRANTY_CARD_NUMBER
                    Command.CommandText &= "',"
                    Command.CommandText &= R_User
                    Command.CommandText &= ",current_timestamp);"

                    'OUT_LOT_MANAGEMENTテーブルへデータ登録
                    Command.ExecuteNonQuery()
                Next
            End If


            '全てのデータをOUT_LOT_MANAGEMENTテーブルにINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' ロット情報を取得する
    ' <引数>
    ' Out_Lot_Data : OUT_TBL.ID
    ' <戻り値>
    ' Out_Lot_Data : 取得したロットデータ
    ' Count : 件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetLOTList(ByVal OUT_ID As Integer, _
                               ByRef Out_Lot_Data() As Lot_List, _
                               ByRef Count As Integer, _
                               ByRef Result As String, _
                               ByRef ErrorMessage As String) As Boolean

        Dim Lot_Data As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim i As Integer = 0

        Dim LOT_List() As Delivery_DetailList = Nothing
        Count = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()

            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_ID,NO,LOT_NUMBER,WARRANTY_CARD_NUMBER FROM OUT_LOT_MANAGEMENT WHERE OUT_ID="
            Command.CommandText &= OUT_ID
            Command.CommandText &= " ORDER BY NO;"

            'SQL実行
            Lot_Data = Command.ExecuteReader()

            Do While (Lot_Data.Read)
                ReDim Preserve Out_Lot_Data(0 To i)
                'OUT_ID
                Out_Lot_Data(i).OUT_ID = Lot_Data("OUT_ID")
                'NO
                Out_Lot_Data(i).NO = Lot_Data("NO")
                'ロット番号
                Out_Lot_Data(i).LOT_NUMBER = Lot_Data("LOT_NUMBER")
                '保証書番号
                Out_Lot_Data(i).WARRANTY_CARD_NUMBER = Lot_Data("WARRANTY_CARD_NUMBER")

                i += 1
                Count += 1
            Loop
            '値を取得できたのでClose
            Lot_Data.Close()

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷予定データを登録する
    ' <引数>
    ' C_Id : 出荷先ＩＤ
    ' Customer_Order_No : 客先発注No
    ' S_Date : 出荷予定日
    ' PLACE_ID : 出荷倉庫
    ' Status : 区分
    ' Comment1 : コメント１
    ' Comment2 : コメント２
    ' S_Status: 出荷ステータス（出荷予定ボタンからの登録は１固定）
    ' Dt : DataGridViewに登録された商品データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Ins_S_Yotei(ByVal C_Id As Integer, _
                            ByVal Customer_Order_No As String, _
                            ByVal OUT_Date As String, _
                            ByVal PLACE_ID As String, _
                            ByVal I_Status As String, _
                            ByVal Comment1 As String, _
                            ByVal Comment2 As String, _
                            ByVal Status As String, _
                            ByVal Dt() As S_Yotei_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            For Count = 0 To Dt.Length - 1


                '出荷予定テーブル用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO OUT_SHIPPING_PLAN(C_ID,CUSTOMER_ORDER_NO,OUT_DATE,PLACE_ID,I_STATUS,"
                Command.CommandText &= "COMMENT1,COMMENT2,STATUS,I_ID,D_UNIT_PRICE,NUM,PLAN_NUM,FIX_NUM,S_STATUS,U_DATE,R_USER)VALUES("
                Command.CommandText &= C_Id
                Command.CommandText &= ",'"
                Command.CommandText &= Customer_Order_No
                Command.CommandText &= "','"
                Command.CommandText &= OUT_Date
                Command.CommandText &= "',"
                Command.CommandText &= PLACE_ID
                Command.CommandText &= ",'"
                Command.CommandText &= I_Status
                Command.CommandText &= "','"
                Command.CommandText &= Comment1
                Command.CommandText &= "','"
                Command.CommandText &= Comment2
                Command.CommandText &= "',"
                Command.CommandText &= Status
                Command.CommandText &= ","
                Command.CommandText &= Dt(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= Dt(Count).UNIT_PRICE
                Command.CommandText &= ","
                Command.CommandText &= Dt(Count).NUM
                Command.CommandText &= ","
                Command.CommandText &= Dt(Count).NUM
                Command.CommandText &= ",0,1,current_timestamp,"
                Command.CommandText &= R_User
                Command.CommandText &= ");"

                '出荷予定テーブルへデータ登録
                Command.ExecuteNonQuery()
            Next

            '出荷予定テーブルに全てINSERTが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示予定データ検索
    ' <引数>
    ' I_CODE : 商品コード
    ' PLACE_ID : 出荷倉庫
    ' I_STATUS : 区分
    ' STATUS : ステータス
    ' S_STATUS : 出荷ステータス
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOutShipping_Plan_Search(ByVal I_CODE As String, _
                                 ByVal PLACE_ID As Integer, _
                                 ByVal I_STATUS As String, _
                                 ByVal STATUS As String, _
                                 ByVal S_STATUS As Integer, _
                                 ByRef SearchResult() As OutShipping_Search_List, _
                                 ByRef Data_Total As Integer, _
                                 ByRef Data_Num_Total As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = ""
            '検索条件よりWhereの作成

            '区分
            WhereSql = "WHERE OUT_SHIPPING_PLAN.I_STATUS ='"
            WhereSql &= I_STATUS
            WhereSql &= "' "

            '出荷ステータス
            WhereSql &= "AND OUT_SHIPPING_PLAN.S_STATUS ="
            WhereSql &= S_STATUS
            WhereSql &= " "

            'ステータス
            WhereSql &= "AND OUT_SHIPPING_PLAN.STATUS ='"
            WhereSql &= STATUS
            WhereSql &= "' "

            '商品コード
            If I_CODE <> "" Then
                WhereSql = "AND M_ITEM.I_CODE ='"
                WhereSql &= I_CODE
                WhereSql &= "' "
            End If

            '倉庫ID
            If WhereSql <> "" Then
                WhereSql &= "AND OUT_SHIPPING_PLAN.PLACE_ID ="
                WhereSql &= PLACE_ID
                WhereSql &= " "
            End If

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " GROUP BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.STATUS"
            'Select実行
            SearchData = Command.ExecuteReader()
            Do While (SearchData.Read)
                Data_Total += 1
            Loop

            If Data_Total = 0 Then
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If
            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select SUM(OUT_SHIPPING_PLAN.NUM) AS TOTAL "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= WhereSql
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成（明細単位Ver）
            'Command = Connection.CreateCommand
            'Command.CommandText = "select OUT_SHIPPING_SCHEDULE.ID,OUT_SHIPPING_SCHEDULE.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.C_NAME,"
            'Command.CommandText &= "OUT_SHIPPING_SCHEDULE.CUSTOMER_ORDER_NO,OUT_SHIPPING_SCHEDULE.OUT_DATE,OUT_SHIPPING_SCHEDULE.PLACE_ID,"
            'Command.CommandText &= "M_PLACE.NAME AS P_NAME,OUT_SHIPPING_SCHEDULE.I_STATUS,OUT_SHIPPING_SCHEDULE.STATUS,OUT_SHIPPING_SCHEDULE.COMMENT1,"
            'Command.CommandText &= "OUT_SHIPPING_SCHEDULE.COMMENT2,OUT_SHIPPING_SCHEDULE.I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,"
            'Command.CommandText &= "OUT_SHIPPING_SCHEDULE.D_UNIT_PRICE,OUT_SHIPPING_SCHEDULE.NUM,OUT_SHIPPING_SCHEDULE.FIX_NUM,"
            'Command.CommandText &= "STOCK.NUM AS STOCK_NUM,OUT_SHIPPING_SCHEDULE.S_STATUS "
            'Command.CommandText &= "FROM (OUT_SHIPPING_SCHEDULE) INNER JOIN M_ITEM ON OUT_SHIPPING_SCHEDULE.I_ID=M_ITEM.ID "
            'Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_SCHEDULE.C_ID=M_CUSTOMER.ID "
            'Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_SCHEDULE.PLACE_ID=M_PLACE.ID "
            'Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_SCHEDULE.I_ID=STOCK.I_ID AND OUT_SHIPPING_SCHEDULE.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_SCHEDULE.PLACE_ID = STOCK.PLACE_ID "
            'Command.CommandText &= WhereSql
            'Command.CommandText &= " ORDER BY OUT_SHIPPING_SCHEDULE.ID"

            'データ取得用コマンド作成（商品ＩＤサマリーVer）
            Command = Connection.CreateCommand
            Command.CommandText = "select SUM(OUT_SHIPPING_PLAN.NUM) AS NUM,SUM(OUT_SHIPPING_PLAN.PLAN_NUM) AS PLAN_NUM,OUT_SHIPPING_PLAN.ID,"
            Command.CommandText &= "(SELECT IFNULL( SUM( NUM ) , 0 ) FROM OUT_TBL WHERE OUT_TBL.I_ID = M_ITEM.ID AND OUT_TBL.STATUS='出荷予定' AND OUT_TBL.I_STATUS=OUT_SHIPPING_PLAN.I_STATUS) AS OUT_NUM, "
            Command.CommandText &= "OUT_SHIPPING_PLAN.PLACE_ID,M_PLACE.NAME AS P_NAME,OUT_SHIPPING_PLAN.I_STATUS,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,OUT_SHIPPING_PLAN.STATUS,"
            Command.CommandText &= "STOCK.NUM AS STOCK_NUM,OUT_SHIPPING_PLAN.S_STATUS "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " GROUP BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.STATUS"
            'Select実行
            SearchData = Command.ExecuteReader()


            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)
                'OUT_SHIPPING_SCHEDULE.ID
                SearchResult(Count).ID = SearchData("ID")
                '出荷倉庫ID
                SearchResult(Count).P_ID = SearchData("PLACE_ID")
                '出荷倉庫名
                SearchResult(Count).P_NAME = SearchData("P_NAME")
                '区分
                SearchResult(Count).I_STATUS = SearchData("I_STATUS")
                '商品ＩＤ
                SearchResult(Count).I_ID = SearchData("I_ID")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '出荷希望数
                SearchResult(Count).NUM = SearchData("NUM")
                '出荷指示予定数量
                SearchResult(Count).PLAN_NUM = SearchData("PLAN_NUM")
                '在庫数
                SearchResult(Count).STOCK_NUM = SearchData("STOCK_NUM")
                '出荷予定数
                SearchResult(Count).OUT_NUM = SearchData("OUT_NUM")
                'S_STATUS
                SearchResult(Count).S_STATUS = SearchData("S_STATUS")
                'STATUS
                SearchResult(Count).STATUS = SearchData("STATUS")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷予定データ検索
    ' <引数>
    ' S_STATUS : 出荷ステータス
    ' P_ID : 倉庫ID
    ' SearchConditions : 検索結果から取得した検索条件（商品ID、区分）
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOutShipping_Plan_DetailSearch(ByVal S_STATUS As Integer, _
                                 ByVal P_ID As Integer, _
                                 ByVal SearchConditions() As OutShipping_Search_List, _
                                 ByRef SearchResult() As OutShipping_Search_List, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Dim Data_Total As Integer = 0

        Dim I_ID As String = Nothing
        Dim I_STATUS As String = Nothing
        Dim STATUS As String = Nothing

        '検索結果をループして商品ＩＤ、ステータス区分を格納
        For Count = 0 To SearchConditions.Length - 1
            If I_ID = "" Then
                I_ID = SearchConditions(Count).I_ID
                I_STATUS = "'" & SearchConditions(Count).I_STATUS & "'"
                STATUS = "'" & SearchConditions(Count).STATUS & "'"
            Else
                I_ID = I_ID & "," & SearchConditions(Count).I_ID
                I_STATUS = I_STATUS & ",'" & SearchConditions(Count).I_STATUS & "'"
                STATUS = STATUS & ",'" & SearchConditions(Count).STATUS & "'"
            End If
        Next

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS="
            Command.CommandText &= S_STATUS
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("COUNT")) Then
                    Data_Total = 0
                Else
                    Data_Total = SearchData("COUNT")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成（明細単位Ver）
            Command = Connection.CreateCommand
            Command.CommandText = "select OUT_SHIPPING_PLAN.ID,OUT_SHIPPING_PLAN.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.C_NAME,M_CUSTOMER.D_NAME,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.OUT_DATE,OUT_SHIPPING_PLAN.PLACE_ID,"
            Command.CommandText &= "M_PLACE.NAME AS P_NAME,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS,OUT_SHIPPING_PLAN.COMMENT1,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.COMMENT2,OUT_SHIPPING_PLAN.I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.D_UNIT_PRICE,OUT_SHIPPING_PLAN.NUM,OUT_SHIPPING_PLAN.PLAN_NUM,OUT_SHIPPING_PLAN.FIX_NUM,"
            Command.CommandText &= "STOCK.NUM AS STOCK_NUM,OUT_SHIPPING_PLAN.S_STATUS "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS="
            Command.CommandText &= S_STATUS
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            Command.CommandText &= " ORDER BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS"

            'Select実行
            SearchData = Command.ExecuteReader()

            ReDim Preserve SearchResult(0 To Data_Total - 1)
            Count = 0
            Do While (SearchData.Read)

                'OUT_SHIPPING_SCHEDULE.ID
                SearchResult(Count).ID = SearchData("ID")
                '出荷先ID
                SearchResult(Count).C_ID = SearchData("C_ID")
                '出荷先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '出荷先名
                SearchResult(Count).C_NAME = SearchData("C_NAME")
                '出荷先名
                SearchResult(Count).D_NAME = SearchData("D_NAME")

                '客先発注No
                SearchResult(Count).CUSTOMER_ORDER_NO = SearchData("CUSTOMER_ORDER_NO")
                '出荷予定日
                SearchResult(Count).OUT_DATE = SearchData("OUT_DATE")
                '倉庫ID
                SearchResult(Count).P_ID = SearchData("PLACE_ID")
                '倉庫名
                SearchResult(Count).P_NAME = SearchData("P_NAME")
                '区分
                SearchResult(Count).I_STATUS = SearchData("I_STATUS")
                'ステータス
                SearchResult(Count).STATUS = SearchData("STATUS")
                'コメント１
                SearchResult(Count).COMMENT1 = SearchData("COMMENT1")
                'コメント２
                SearchResult(Count).COMMENT2 = SearchData("COMMENT2")
                '商品ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '納品単価
                SearchResult(Count).D_UNIT_PRICE = SearchData("D_UNIT_PRICE")
                '出荷希望数
                SearchResult(Count).NUM = SearchData("NUM")
                '出荷指示予定数
                SearchResult(Count).PLAN_NUM = SearchData("PLAN_NUM")
                '出荷指示済数
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '現在庫数
                SearchResult(Count).STOCK_NUM = SearchData("STOCK_NUM")
                '出荷ステータス
                SearchResult(Count).S_STATUS = SearchData("S_STATUS")
                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷予定データの出荷希望数を修正する
    ' <引数>
    ' Dt : DataGridViewに登録された商品データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Out_Shipping_Update(ByVal Dt() As OutShipping_Search_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Dim UpdaeString As String = Nothing
        Dim sql As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction
            'データ件数分ループ
            For Count = 0 To Dt.Length - 1
                sql = Nothing
                'もし、予定数と出荷済み数が同じなら予定数が０なら、もう出荷しないことになるのでステータス変更。
                If Dt(Count).PLAN_NUM = 0 And Dt(Count).NUM = Dt(Count).FIX_NUM Then
                    sql = ", S_STATUS='出荷指示登録済' "
                End If

                '出荷予定テーブル修正用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_SHIPPING_PLAN SET NUM="
                Command.CommandText &= Dt(Count).NUM
                Command.CommandText &= ",PLAN_NUM="
                Command.CommandText &= Dt(Count).PLAN_NUM
                Command.CommandText &= sql
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Dt(Count).ID
                Command.CommandText &= ";"

                '出荷予定テーブルへUPDATE実行
                Command.ExecuteNonQuery()
            Next

            '出荷予定テーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷指示データを登録する
    ' <引数>
    ' P_ID:倉庫ID
    ' RegistData : 登録データを格納（配列）
    ' OUT_DATE : 出荷予定日
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function InsOutShipping_Plan(ByVal P_ID As Integer, _
                                        ByVal RegistData() As OutShipping_Search_List, _
                                        ByVal OUT_DATE As String, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim SearchData As MySqlDataReader

        Dim ForInsertData As MySqlDataReader
        Dim I_ID As String = Nothing
        Dim I_STATUS As String = Nothing
        Dim STATUS As String = Nothing
        Dim ID As String = Nothing

        Dim SearchResult() As OutShipping_Search_List
        Dim ForInsertDataResult() As OutShipping_Search_List
        Dim Data_Total As Integer
        Dim INS_Data_Total As Integer
        Dim Update_FIX_NUM As Integer
        Dim Update_PLAN_NUM As Integer
        Dim Update_S_Status As Boolean = False '(Trueのとき、アップデート）

        '出荷指示ファイルを作成
        Dim FILE_NAME As String = Nothing
        '出荷指示ファイルの後ろに追加する本日の出荷指示回数を算出時に使用
        Dim FILE_NAME_No As String = Nothing
        Dim dtNow As DateTime
        dtNow = DateTime.Now
        Dim UNIQUE_KEY As String = dtNow.ToString("yyyyMMddhhmmss")

        '伝票番号の自動採番に使用。
        Dim Check_ORDER_NO As String = Nothing
        Dim SHEET_NO_Count As Integer
        Dim SHEET_NO_SERIAL_NO As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            '明細データを取得する。
            '検索結果をループして商品ＩＤ、ステータス区分、後で使用するIDを格納
            For Count = 0 To RegistData.Length - 1
                If I_ID = "" Then
                    I_ID = RegistData(Count).I_ID
                    I_STATUS = "'" & RegistData(Count).I_STATUS & "'"
                    STATUS = "'" & RegistData(Count).STATUS & "'"
                    ID = RegistData(Count).ID
                Else
                    I_ID = I_ID & "," & RegistData(Count).I_ID
                    I_STATUS = I_STATUS & ",'" & RegistData(Count).I_STATUS & "'"
                    STATUS = STATUS & ",'" & RegistData(Count).STATUS & "'"
                    ID = ID & "," & RegistData(Count).ID
                End If
            Next

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS='出荷予定'"
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            'Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLAN_NUM <> 0"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("COUNT")) Then
                    Data_Total = 0
                Else
                    Data_Total = SearchData("COUNT")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成（明細単位Ver）
            Command = Connection.CreateCommand
            Command.CommandText = "select OUT_SHIPPING_PLAN.ID,OUT_SHIPPING_PLAN.C_ID,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.OUT_DATE,OUT_SHIPPING_PLAN.PLACE_ID,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS,OUT_SHIPPING_PLAN.COMMENT1,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.COMMENT2,OUT_SHIPPING_PLAN.I_ID,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.D_UNIT_PRICE,OUT_SHIPPING_PLAN.NUM,OUT_SHIPPING_PLAN.PLAN_NUM,OUT_SHIPPING_PLAN.FIX_NUM,"
            Command.CommandText &= "STOCK.NUM AS STOCK_NUM,OUT_SHIPPING_PLAN.S_STATUS "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS='出荷予定'"
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            'Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLAN_NUM <> 0 ORDER BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS"
            Command.CommandText &= " ORDER BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS"

            'Select実行
            SearchData = Command.ExecuteReader()

            ReDim Preserve SearchResult(0 To Data_Total - 1)
            Count = 0
            Do While (SearchData.Read)
                'If SearchData("PLAN_NUM") <> 0 Then

                'OUT_SHIPPING_SCHEDULE.ID
                SearchResult(Count).ID = SearchData("ID")
                '出荷先ID
                SearchResult(Count).C_ID = SearchData("C_ID")
                '客先発注No
                SearchResult(Count).CUSTOMER_ORDER_NO = SearchData("CUSTOMER_ORDER_NO")
                '出荷予定日
                SearchResult(Count).OUT_DATE = SearchData("OUT_DATE")
                '倉庫ID
                SearchResult(Count).P_ID = SearchData("PLACE_ID")
                '区分
                SearchResult(Count).I_STATUS = SearchData("I_STATUS")
                'ステータス
                SearchResult(Count).STATUS = SearchData("STATUS")
                'コメント１
                SearchResult(Count).COMMENT1 = SearchData("COMMENT1")
                'コメント２
                SearchResult(Count).COMMENT2 = SearchData("COMMENT2")
                '商品ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                '納品単価
                SearchResult(Count).D_UNIT_PRICE = SearchData("D_UNIT_PRICE")
                '出荷希望数
                SearchResult(Count).NUM = SearchData("NUM")
                '出荷指示予定数
                SearchResult(Count).PLAN_NUM = SearchData("PLAN_NUM")
                '出荷指示済数
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '現在庫数
                SearchResult(Count).STOCK_NUM = SearchData("STOCK_NUM")
                '出荷ステータス
                SearchResult(Count).S_STATUS = SearchData("S_STATUS")
                Count += 1
                'End If
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "出荷指示を行うデータがみつかりません。"
                Result = False
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'OUT_PRT、OUT_TBLに登録すべきデータ数を取得する
            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS='出荷予定'"
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLAN_NUM <> 0 ORDER BY M_CUSTOMER.C_CODE,OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.I_ID"

            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("COUNT")) Then
                    INS_Data_Total = 0
                Else
                    INS_Data_Total = SearchData("COUNT")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()


            'データ取得用コマンド作成（OUT_PRT，OUT_TBLに登録するためにデータをオーダー番号、商品IDでソートして取得する）
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_SHIPPING_PLAN.ID,OUT_SHIPPING_PLAN.C_ID,M_CUSTOMER.C_CODE,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.OUT_DATE,OUT_SHIPPING_PLAN.PLACE_ID,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS,OUT_SHIPPING_PLAN.COMMENT1,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.COMMENT2,OUT_SHIPPING_PLAN.I_ID,M_ITEM.I_CODE,M_ITEM.PRICE,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.D_UNIT_PRICE,OUT_SHIPPING_PLAN.NUM,OUT_SHIPPING_PLAN.PLAN_NUM,OUT_SHIPPING_PLAN.FIX_NUM,"
            Command.CommandText &= "STOCK.NUM AS STOCK_NUM,OUT_SHIPPING_PLAN.S_STATUS "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN STOCK ON OUT_SHIPPING_PLAN.I_ID=STOCK.I_ID AND OUT_SHIPPING_PLAN.I_STATUS = STOCK.I_STATUS AND OUT_SHIPPING_PLAN.PLACE_ID = STOCK.PLACE_ID "
            Command.CommandText &= "WHERE OUT_SHIPPING_PLAN.I_ID in ("
            Command.CommandText &= I_ID
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.I_STATUS in ("
            Command.CommandText &= I_STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.STATUS in ("
            Command.CommandText &= STATUS
            Command.CommandText &= ") AND OUT_SHIPPING_PLAN.S_STATUS='出荷予定'"
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLACE_ID="
            Command.CommandText &= P_ID
            'Command.CommandText &= " ORDER BY OUT_SHIPPING_PLAN.I_ID,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS"
            Command.CommandText &= " AND OUT_SHIPPING_PLAN.PLAN_NUM <> 0 ORDER BY M_CUSTOMER.C_CODE,OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.I_ID"
            'Command.CommandText &= " ORDER BY M_CUSTOMER.C_CODE,OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.I_ID"

            'Select実行
            ForInsertData = Command.ExecuteReader()

            ReDim Preserve ForInsertDataResult(0 To INS_Data_Total - 1)
            Count = 0
            Do While (ForInsertData.Read)

                'OUT_SHIPPING_SCHEDULE.ID
                ForInsertDataResult(Count).ID = ForInsertData("ID")
                '出荷先ID
                ForInsertDataResult(Count).C_ID = ForInsertData("C_ID")
                '出荷先コード
                ForInsertDataResult(Count).C_CODE = ForInsertData("C_CODE")
                '客先発注No
                ForInsertDataResult(Count).CUSTOMER_ORDER_NO = ForInsertData("CUSTOMER_ORDER_NO")
                '出荷予定日
                ForInsertDataResult(Count).OUT_DATE = ForInsertData("OUT_DATE")
                '倉庫ID
                ForInsertDataResult(Count).P_ID = ForInsertData("PLACE_ID")
                '区分
                ForInsertDataResult(Count).I_STATUS = ForInsertData("I_STATUS")
                'ステータス
                ForInsertDataResult(Count).STATUS = ForInsertData("STATUS")
                'コメント１
                ForInsertDataResult(Count).COMMENT1 = ForInsertData("COMMENT1")
                'コメント２
                ForInsertDataResult(Count).COMMENT2 = ForInsertData("COMMENT2")
                '商品ID
                ForInsertDataResult(Count).I_ID = ForInsertData("I_ID")
                '商品コード
                ForInsertDataResult(Count).I_CODE = ForInsertData("I_CODE")
                '納品単価
                ForInsertDataResult(Count).D_UNIT_PRICE = ForInsertData("D_UNIT_PRICE")
                '出荷指示予定数
                ForInsertDataResult(Count).PLAN_NUM = ForInsertData("PLAN_NUM")
                '出荷ステータス
                ForInsertDataResult(Count).S_STATUS = ForInsertData("S_STATUS")
                '商品マスタの出荷ステータス
                ForInsertDataResult(Count).UNIT_PRICE = ForInsertData("PRICE")

                Count += 1
            Loop

            'Close
            ForInsertData.Close()

            '明細データを使い、OUT_SHIPPING_PLANをアップデート
            '出荷指示予定数+出荷指示済み数を出荷指示済み数に、
            '出荷希望数 - (出荷指示予定数+出荷指示済み数)が0以上なら結果を出荷指示予定数に
            '上記で0なら、ステータスを出荷指示登録済みに変更
            For Count = 0 To SearchResult.Length - 1
                If SearchResult(Count).PLAN_NUM <> 0 Then


                    Update_FIX_NUM = 0
                    Update_PLAN_NUM = 0
                    '出荷指示予定数+出荷指示済み数を出荷指示済み数にするため、数値を求める
                    Update_FIX_NUM = SearchResult(Count).PLAN_NUM + SearchResult(Count).FIX_NUM

                    '通常出荷の場合、出荷希望数 - 出荷指示予定数 - 出荷指示済数が0以上なら結果を出荷指示予定数に格納にするため、数値を求める
                    If SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM > 0 And SearchResult(Count).STATUS = "通常出荷" Then
                        Update_PLAN_NUM = SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM
                    ElseIf SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM = 0 And SearchResult(Count).STATUS = "通常出荷" Then
                        '通常出荷の場合、出荷希望数 - 出荷指示予定数 - 出荷指示済数が0ならPLAN_NUMに0をいれ、ステータスを出荷指示済み数にする
                        Update_PLAN_NUM = 0
                        Update_S_Status = True
                    ElseIf SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM < 0 And SearchResult(Count).STATUS = "伝票出力のみ" Then
                        '伝票出力のみの場合、出荷希望数 - 出荷指示予定数 - 出荷指示済数が0以下なら結果を出荷指示予定数に格納にするため、数値を求める
                        Update_PLAN_NUM = SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM

                    ElseIf SearchResult(Count).NUM - SearchResult(Count).PLAN_NUM - SearchResult(Count).FIX_NUM = 0 And SearchResult(Count).STATUS = "伝票出力のみ" Then
                        '通常出荷の場合、出荷希望数 - 出荷指示予定数 - 出荷指示済数が0ならPLAN_NUMに0をいれ、ステータスを出荷指示済み数にする
                        Update_PLAN_NUM = 0
                        Update_S_Status = True
                    End If
                Else
                    'PLAN_NUMが0の場合、OUT_SHIPPING_PLANのPLAN_NUMに対して再計算（受注数 - 出荷指示済み数）を行い
                    'PLAN_NUMに受注残数を修正する。
                    Update_FIX_NUM = SearchResult(Count).FIX_NUM
                    Update_PLAN_NUM = SearchResult(Count).NUM - SearchResult(Count).FIX_NUM
                    Update_S_Status = False

                End If
                '出荷予定テーブル用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_SHIPPING_PLAN SET FIX_NUM="
                Command.CommandText &= Update_FIX_NUM
                Command.CommandText &= ", PLAN_NUM="
                Command.CommandText &= Update_PLAN_NUM
                If Update_S_Status = True Then
                    Command.CommandText &= ", S_STATUS='出荷指示登録済'"
                End If
                Command.CommandText &= " WHERE OUT_SHIPPING_PLAN.ID="
                Command.CommandText &= SearchResult(Count).ID

                '出荷予定テーブルへデータUpdate
                Command.ExecuteNonQuery()

            Next

            '事前に取得したForInsertDataResultを使い、OUT_PRT、OUT_TBLにデータを登録する。
            '出荷指示ファイル名を作成するため本日の出荷回数を求める。

            '出荷指示ファイル名の作成を行う。
            '出荷指示ファイル名は N + YYMMDD + 本日の出荷指示回数（2桁）なので
            '本日の出荷指示回数（2桁）をOUT_TBLから本日の出荷指示回数を取得
            FILE_NAME = "N" & dtNow.ToString("yyMMdd")

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select MAX(FILE_NAME) AS FILE_NAME "
            Command.CommandText &= "FROM OUT_TBL "
            Command.CommandText &= "WHERE OUT_TBL.FILE_NAME LIKE '"
            Command.CommandText &= FILE_NAME
            Command.CommandText &= "%'"

            Dim Int_FILE_NAME_No As Integer
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("FILE_NAME")) Then
                    FILE_NAME_No = "01"
                Else
                    FILE_NAME_No = SearchData("FILE_NAME").Remove(0, 7)
                    Int_FILE_NAME_No = Integer.Parse(FILE_NAME_No) + 1
                    FILE_NAME_No = Int_FILE_NAME_No.ToString("D2")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "本日の出荷指示回数が取得できませんでした。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            '出荷指示ファイル名は N + YYMMDD + 本日の出荷指示回数（2桁）
            FILE_NAME = FILE_NAME & FILE_NAME_No

            SHEET_NO_Count = 1

            'OUT_PRTに登録する
            For Count = 0 To ForInsertDataResult.Length - 1
                '伝票番号の後ろの連番を作成する。
                '伝票番号のルール：YYMMDD + 出荷指示の回数（2桁） + 4桁の連番
                '納品先コードが同じでもオーダー番号が違えば別の伝票番号をふる。
                '1件目は固定で0001を入れ、比較用のオーダー番号を格納
                If Count = 0 Then
                    SHEET_NO_SERIAL_NO = "0001"
                    Check_ORDER_NO = ForInsertDataResult(Count).CUSTOMER_ORDER_NO
                Else
                    '2件目以降は1つ前のオーダー番号と今の行のオーダー番号を比較し、同じであれば
                    'SHEET_NO_SERIAL_NOは同じものを引き続き使用し、
                    '違う場合はカウントアップして格納する。
                    If Check_ORDER_NO = ForInsertDataResult(Count).CUSTOMER_ORDER_NO Then
                        Check_ORDER_NO = ForInsertDataResult(Count).CUSTOMER_ORDER_NO
                    Else
                        SHEET_NO_Count += 1
                        SHEET_NO_SERIAL_NO = SHEET_NO_Count.ToString("D4")
                        Check_ORDER_NO = ForInsertDataResult(Count).CUSTOMER_ORDER_NO
                    End If
                End If

                'OUT_PRT登録テーブル用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO OUT_PRT(UNIQUE_KEY,FILE_NAME,SHEET_NO,ORDER_NO,I_CODE,C_CODE,I_ID,C_ID,UNIT_COST,"
                Command.CommandText &= "NUM,TOTAL_AMOUNT,COMMENT1,COMMENT2,PLACE_ID,O_DATE,PRT_STATUS,R_STATUS,U_DATE,U_USER)VALUES('"
                Command.CommandText &= UNIQUE_KEY
                Command.CommandText &= "','"
                Command.CommandText &= FILE_NAME
                Command.CommandText &= "','"
                Command.CommandText &= FILE_NAME & SHEET_NO_SERIAL_NO
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).CUSTOMER_ORDER_NO
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).I_CODE
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).C_CODE
                Command.CommandText &= "',"
                Command.CommandText &= ForInsertDataResult(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).C_ID
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).D_UNIT_PRICE
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).PLAN_NUM
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).PLAN_NUM * ForInsertDataResult(Count).D_UNIT_PRICE
                Command.CommandText &= ",'"
                Command.CommandText &= ForInsertDataResult(Count).COMMENT1
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).COMMENT2
                Command.CommandText &= "',"
                Command.CommandText &= ForInsertDataResult(Count).P_ID
                Command.CommandText &= ",'"
                Command.CommandText &= OUT_DATE
                Command.CommandText &= "','未印刷','登録済み',current_timestamp,"
                Command.CommandText &= R_User
                Command.CommandText &= ");"

                'OUT_PRTテーブルへデータ登録
                Command.ExecuteNonQuery()

                'OUT_TBL登録テーブル用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "INSERT INTO OUT_TBL(FILE_NAME,SHEET_NO,ORDER_NO,I_ID,C_ID,UNIT_PRICE,UNIT_COST,NUM,FIX_NUM,"
                Command.CommandText &= "COMMENT1,COMMENT2,DATE,STATUS,CATEGORY,I_STATUS,PLACE_ID,FIX_DATE,PRT_DATE,U_DATE,U_USER)VALUES('"
                Command.CommandText &= FILE_NAME
                Command.CommandText &= "','"
                Command.CommandText &= FILE_NAME & SHEET_NO_SERIAL_NO
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).CUSTOMER_ORDER_NO
                Command.CommandText &= "',"
                Command.CommandText &= ForInsertDataResult(Count).I_ID
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).C_ID
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).UNIT_PRICE
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).D_UNIT_PRICE
                Command.CommandText &= ","
                Command.CommandText &= ForInsertDataResult(Count).PLAN_NUM
                Command.CommandText &= ",0,'"
                Command.CommandText &= ForInsertDataResult(Count).COMMENT1
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).COMMENT2
                Command.CommandText &= "','"
                Command.CommandText &= OUT_DATE
                Command.CommandText &= "','"
                Command.CommandText &= ForInsertDataResult(Count).S_STATUS
                Command.CommandText &= "','通常出荷','"
                Command.CommandText &= ForInsertDataResult(Count).I_STATUS
                Command.CommandText &= "',"
                Command.CommandText &= ForInsertDataResult(Count).P_ID
                Command.CommandText &= ",NULL,'0000/00/00',current_timestamp,"
                Command.CommandText &= R_User
                Command.CommandText &= ");"

                'OUT_TBLテーブルへデータ登録
                Command.ExecuteNonQuery()
            Next


            '出荷予定テーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' オーダー番号がすでに使用されているかチェックする
    ' <引数>
    ' OrderNo : オーダー番号
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function OrderNo_Check(ByRef OrderNo As String, _
                                ByRef Result As String, _
                                ByRef ErrorMessage As String) As Boolean
        Dim CustmerData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer = 0
        Dim ResultCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()
            'コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = " SELECT COUNT(*) AS COUNT FROM OUT_SHIPPING_PLAN WHERE CUSTOMER_ORDER_NO='"
            Command.CommandText &= OrderNo
            Command.CommandText &= "';"

            'データ取得
            CustmerData = Command.ExecuteReader(CommandBehavior.SingleRow)
            If CustmerData.Read Then
                ' レコードが取得できた時の処理
                ResultCount = CustmerData("COUNT")

                If ResultCount >= 1 Then
                    ErrorMessage = "入力されたオーダー番号は既に登録済みです。"
                    Result = False
                    Exit Function
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "エラーが発生しました。"
                Result = False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 受注データ検索
    ' <引数>
    ' I_CODE : 商品コード
    ' PLACE_ID : 出荷倉庫
    ' I_STATUS : 区分
    ' STATUS : ステータス
    ' S_STATUS : 出荷ステータス
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Total : 検索結果レコード数
    ' Data_Num_Total : 検索結果の数量総数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOrder_Search(ByVal I_CODE As String, _
                                 ByVal C_CODE As String, _
                                 ByVal ORDER_NO As String, _
                                 ByVal OUT_DATE_FROM As String, _
                                 ByVal OUT_DATE_TO As String, _
                                 ByVal REGIST_DATE_FROM As String, _
                                 ByVal REGIST_DATE_TO As String, _
                                 ByVal I_STATUS As String, _
                                 ByVal STATUS As String, _
                                 ByVal PLACE_ID As Integer, _
                                 ByVal S_STATUS As Integer, _
                                 ByVal COMMENT As String, _
                                 ByRef SearchResult() As OutShipping_Search_List, _
                                 ByRef Data_Total As Integer, _
                                 ByRef Data_Num_Total As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = ""
            '検索条件よりWhereの作成

            '必須項目を検索条件に設定
            '区分
            WhereSql = "WHERE OUT_SHIPPING_PLAN.I_STATUS ='"
            WhereSql &= I_STATUS
            WhereSql &= "' "

            'ステータス
            WhereSql &= "AND OUT_SHIPPING_PLAN.STATUS ='"
            WhereSql &= STATUS
            WhereSql &= "' "

            '出荷ステータス
            WhereSql &= "AND OUT_SHIPPING_PLAN.S_STATUS ="
            WhereSql &= S_STATUS
            WhereSql &= " "

            '倉庫ID
            WhereSql &= "AND OUT_SHIPPING_PLAN.PLACE_ID ="
            WhereSql &= PLACE_ID
            WhereSql &= " "

            '商品コード
            If I_CODE <> "" Then
                WhereSql &= "AND M_ITEM.I_CODE ='"
                WhereSql &= I_CODE
                WhereSql &= "' "
            End If

            '出荷先コード
            If C_CODE <> "" Then
                WhereSql &= "AND M_CUSTOMER.C_CODE ='"
                WhereSql &= C_CODE
                WhereSql &= "' "
            End If

            'オーダー番号
            If ORDER_NO <> "" Then
                WhereSql &= "AND OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO ='"
                WhereSql &= ORDER_NO
                WhereSql &= "' "
            End If

            '出荷予定日From
            If OUT_DATE_FROM <> "" Then
                WhereSql &= " AND OUT_SHIPPING_PLAN.OUT_DATE >= '"
                WhereSql &= OUT_DATE_FROM
                WhereSql &= "'"
            End If

            '出荷予定日To
            If OUT_DATE_TO <> "" Then
                'Toのみ入力されている場合
                WhereSql &= " AND OUT_SHIPPING_PLAN.OUT_DATE <= '"
                WhereSql &= OUT_DATE_TO
                WhereSql &= "'"
            End If

            '登録日時From
            If REGIST_DATE_FROM <> "" Then
                'Fromのみ入力されている場合
                WhereSql &= " AND DATE_FORMAT(OUT_SHIPPING_PLAN.U_DATE, '%Y/%m/%d') >= '"
                WhereSql &= REGIST_DATE_FROM
                WhereSql &= "'"
            End If

            '登録日時TO
            If REGIST_DATE_TO <> "" Then
                'TOのみ入力されている場合
                WhereSql &= " AND DATE_FORMAT(OUT_SHIPPING_PLAN.U_DATE, '%Y/%m/%d') <= '"
                WhereSql &= REGIST_DATE_TO
                WhereSql &= "'"
            End If

            'コメント
            If COMMENT <> "" Then
                WhereSql &= "AND OUT_SHIPPING_PLAN.COMMENT1 LIKE '%"
                WhereSql &= COMMENT
                WhereSql &= "%' OR OUT_SHIPPING_PLAN.COMMENT2 LIKE '%"
                WhereSql &= COMMENT
                WhereSql &= "%' "
            End If

            'オープン
            Connection.Open()

            'データ件数を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select COUNT(*) AS COUNT "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY OUT_SHIPPING_PLAN.ID"
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("COUNT")) Then
                    Data_Total = 0
                Else
                    Data_Total = SearchData("COUNT")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ件数の数量合計を取得するコマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "select SUM(OUT_SHIPPING_PLAN.NUM) AS TOTAL "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= WhereSql
            'Select実行
            SearchData = Command.ExecuteReader(CommandBehavior.SingleRow)

            If SearchData.Read Then
                ' レコードが取得できた時の処理
                If IsDBNull(SearchData("TOTAL")) Then
                    Data_Num_Total = 0
                Else
                    Data_Num_Total = SearchData("TOTAL")
                End If
            Else
                ' レコードが取得できなかった時の処理
                ErrorMessage = "入力した検索条件に該当するデータがみつかりません。"
                Exit Function
            End If

            'COUNT値を取得できたのでClose
            SearchData.Close()

            'データ取得用コマンド作成（明細単位Ver）
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_SHIPPING_PLAN.ID,OUT_SHIPPING_PLAN.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.CUSTOMER_ORDER_NO,OUT_SHIPPING_PLAN.OUT_DATE,OUT_SHIPPING_PLAN.PLACE_ID,"
            Command.CommandText &= "M_PLACE.NAME AS P_NAME,OUT_SHIPPING_PLAN.I_STATUS,OUT_SHIPPING_PLAN.STATUS,OUT_SHIPPING_PLAN.COMMENT1,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.COMMENT2,OUT_SHIPPING_PLAN.I_ID,M_ITEM.I_CODE,M_ITEM.I_NAME,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.D_UNIT_PRICE,OUT_SHIPPING_PLAN.NUM,OUT_SHIPPING_PLAN.FIX_NUM,"
            Command.CommandText &= "OUT_SHIPPING_PLAN.PLAN_NUM,OUT_SHIPPING_PLAN.S_STATUS,DATE_FORMAT(OUT_SHIPPING_PLAN.U_DATE, '%Y/%m/%d') AS U_DATE "
            Command.CommandText &= "FROM (OUT_SHIPPING_PLAN) INNER JOIN M_ITEM ON OUT_SHIPPING_PLAN.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_SHIPPING_PLAN.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_SHIPPING_PLAN.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY OUT_SHIPPING_PLAN.ID"

            'Select実行
            SearchData = Command.ExecuteReader()


            ReDim Preserve SearchResult(0 To Data_Total - 1)

            Do While (SearchData.Read)
                'OUT_SHIPPING_PLAN.ID
                SearchResult(Count).ID = SearchData("ID")
                '出荷先ID
                SearchResult(Count).C_ID = SearchData("C_ID")
                '出荷先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '出荷先名
                SearchResult(Count).D_NAME = SearchData("D_NAME")
                'オーダー番号
                SearchResult(Count).CUSTOMER_ORDER_NO = SearchData("CUSTOMER_ORDER_NO")
                '出荷予定日
                SearchResult(Count).OUT_DATE = SearchData("OUT_DATE")
                '出荷倉庫ID
                SearchResult(Count).P_ID = SearchData("PLACE_ID")
                '出荷倉庫名
                SearchResult(Count).P_NAME = SearchData("P_NAME")
                '区分
                SearchResult(Count).I_STATUS = SearchData("I_STATUS")
                'ステータス
                SearchResult(Count).STATUS = SearchData("STATUS")
                'コメント１
                If IsDBNull(SearchData("COMMENT1")) Then
                    SearchResult(Count).COMMENT1 = ""
                Else
                    SearchResult(Count).COMMENT1 = SearchData("COMMENT1")
                End If

                'コメント２
                If IsDBNull(SearchData("COMMENT2")) Then
                    SearchResult(Count).COMMENT2 = ""
                Else
                    SearchResult(Count).COMMENT2 = SearchData("COMMENT2")
                End If
                '商品ＩＤ
                SearchResult(Count).I_ID = SearchData("I_ID")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '納品単価
                SearchResult(Count).D_UNIT_PRICE = SearchData("D_UNIT_PRICE")
                '出荷希望数
                SearchResult(Count).NUM = SearchData("NUM")
                '出荷指示予定数量
                SearchResult(Count).PLAN_NUM = SearchData("PLAN_NUM")
                '出荷指示済数量
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '出荷ステータス
                SearchResult(Count).S_STATUS = SearchData("S_STATUS")
                '登録日時
                SearchResult(Count).U_DATE = SearchData("U_DATE")

                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 受注予定データの出荷希望数を修正する
    ' <引数>
    ' Dt : DataGridViewに登録された商品データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Upd_OrderData(ByVal Dt() As OutShipping_Search_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer
        Dim sql As String = Nothing


        Dim UpdaeString As String = Nothing
        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction
            'データ件数分ループ
            For Count = 0 To Dt.Length - 1
                sql = Nothing
                If Dt(Count).PLAN_NUM = 0 Then
                    sql = ",S_STATUS='出荷指示登録済' "
                End If

                '出荷予定テーブル修正用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_SHIPPING_PLAN SET NUM="
                Command.CommandText &= Dt(Count).NUM
                Command.CommandText &= ", PLAN_NUM="
                Command.CommandText &= Dt(Count).PLAN_NUM
                Command.CommandText &= sql
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Dt(Count).ID
                Command.CommandText &= ";"

                '出荷予定テーブルへUPDATE実行
                Command.ExecuteNonQuery()
            Next

            '出荷予定テーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 受注予定データを削除する
    ' <引数>
    ' Dt : DataGridViewに登録された商品データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Del_OrderData(ByVal Dt() As OutShipping_Search_List, _
                            ByRef Result As String, _
                            ByRef ErrorMessage As String) As Boolean

        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Dim UpdaeString As String = Nothing
        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction
            'データ件数分ループ
            For Count = 0 To Dt.Length - 1

                '出荷予定テーブル修正用　コマンド作成
                Command = Connection.CreateCommand
                Command.CommandText = "DELETE FROM OUT_SHIPPING_PLAN WHERE ID="
                Command.CommandText &= Dt(Count).ID
                Command.CommandText &= ";"

                '出荷予定テーブルへDELETE実行
                Command.ExecuteNonQuery()
            Next

            '出荷予定テーブルに全てDELETEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' ロケーション変更を行う。
    '
    ' <引数>
    ' Location_Change_Data : 変更データ格納配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function Location_Change(ByRef Location_Change_Data() As Location_Change_List, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'データ件数ループ
            For Count = 0 To Location_Change_Data.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE STOCK SET LOCATION='"
                Command.CommandText &= Location_Change_Data(Count).NEW_LOCATION
                Command.CommandText &= "',U_DATE=Current_Timestamp WHERE ID="
                Command.CommandText &= Location_Change_Data(Count).STOCK_ID
                Command.CommandText &= ";"

                'UPDATE実行
                Command.ExecuteNonQuery()

            Next

            'STOCKテーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 出荷予定検索 - 出荷予定、出庫済み情報を取得する。
    ' <引数>
    ' CheckData : OUT_TBL.IDが格納されたデータ
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Data_Count : 件数
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetOutLOTSearch(ByVal CheckData() As Out_Search_List, _
                                 ByRef SearchResult() As Out_Search_List, _
                                 ByRef Data_Count As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing
        Dim DataCount As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = "WHERE OUT_TBL.ID in ("
            '検索条件よりWhereの作成

            For i = 0 To CheckData.Length - 1
                WhereSql &= CheckData(i).ID
                If i = CheckData.Length - 1 Then
                    WhereSql &= ") "
                Else
                    WhereSql &= ","
                End If
            Next

            'オープン
            Connection.Open()

            'データ取得用コマンド作成
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,OUT_TBL.C_ID,OUT_TBL.SHEET_NO,OUT_TBL.ORDER_NO,OUT_TBL.I_ID,"
            Command.CommandText &= "M_ITEM.I_CODE,M_ITEM.I_NAME,OUT_TBL.C_ID,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME AS C_NAME,"
            Command.CommandText &= "OUT_TBL.NUM,OUT_TBL.FIX_NUM,OUT_TBL.DATE,OUT_TBL.FIX_DATE,M_CUSTOMER.D_ZIP,M_CUSTOMER.D_ADDRESS,"
            Command.CommandText &= "OUT_TBL.FILE_NAME,OUT_TBL.STATUS,OUT_TBL.CATEGORY,OUT_TBL.I_STATUS,"
            Command.CommandText &= "OUT_TBL.UNIT_PRICE, OUT_TBL.UNIT_COST,DATE_FORMAT(OUT_TBL.PRT_DATE, '%Y/%m/%d') AS PRT_DATE,"
            Command.CommandText &= "OUT_TBL.COMMENT1, OUT_TBL.COMMENT2, OUT_TBL.REMARKS,OUT_TBL.PLACE_ID as P_ID,M_PLACE.NAME AS P_NAME,"
            Command.CommandText &= "OUT_LOT_MANAGEMENT.NO, OUT_LOT_MANAGEMENT.LOT_NUMBER, OUT_LOT_MANAGEMENT.WARRANTY_CARD_NUMBER "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_TBL.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= "INNER JOIN OUT_LOT_MANAGEMENT ON OUT_TBL.ID=OUT_LOT_MANAGEMENT.OUT_ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY OUT_TBL.ID,OUT_LOT_MANAGEMENT.NO"
            'Select実行
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To DataCount)
                'OUT.ID
                SearchResult(DataCount).ID = SearchData("ID")
                'I_ID
                SearchResult(DataCount).I_ID = SearchData("I_ID")
                '伝票番号
                SearchResult(DataCount).SHEET_NO = SearchData("SHEET_NO")
                'オーダー番号
                SearchResult(DataCount).ORDER_NO = SearchData("ORDER_NO")
                '商品コード
                SearchResult(DataCount).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(DataCount).I_NAME = SearchData("I_NAME")
                '納品先コード
                SearchResult(DataCount).C_CODE = SearchData("C_CODE")
                '納品先名
                SearchResult(DataCount).C_NAME = SearchData("C_NAME")
                '出荷予定数
                SearchResult(DataCount).NUM = SearchData("NUM")
                '出荷予定日
                SearchResult(DataCount).O_DATE = SearchData("DATE")
                '出庫数
                SearchResult(DataCount).FIX_NUM = SearchData("FIX_NUM")
                '出荷日
                If IsDBNull(SearchData("FIX_DATE")) Then
                    SearchResult(DataCount).FIX_DATE = ""
                Else
                    SearchResult(DataCount).FIX_DATE = SearchData("FIX_DATE")
                End If
                '出荷指示ファイル名
                SearchResult(DataCount).FILE_NAME = SearchData("FILE_NAME")
                'ステータス
                SearchResult(DataCount).STATUS = SearchData("STATUS")
                '種別
                SearchResult(DataCount).CATEGORY = SearchData("CATEGORY")
                '商品ステータス
                SearchResult(DataCount).DEFECT_TYPE = SearchData("I_STATUS")
                '納入単価
                SearchResult(DataCount).PRICE = SearchData("UNIT_PRICE")
                '印刷日
                SearchResult(DataCount).PRT_DATE = CStr(SearchData("PRT_DATE"))
                '売単価
                SearchResult(DataCount).COST = SearchData("UNIT_COST")
                'コメント１
                SearchResult(DataCount).COMMENT1 = SearchData("COMMENT1")
                'コメント２
                SearchResult(DataCount).COMMENT2 = SearchData("COMMENT2")
                '備考
                If IsDBNull(SearchData("REMARKS")) Then
                    SearchResult(DataCount).REMARKS = ""
                Else
                    SearchResult(DataCount).REMARKS = SearchData("REMARKS")
                End If

                '倉庫
                SearchResult(DataCount).PLACE = SearchData("P_NAME")
                '倉庫ID
                SearchResult(DataCount).P_ID = SearchData("P_ID")

                '郵便番号
                If IsDBNull(SearchData("D_ZIP")) Then
                    SearchResult(DataCount).D_ZIP = ""
                Else
                    SearchResult(DataCount).D_ZIP = SearchData("D_ZIP")
                End If
                '住所
                If IsDBNull(SearchData("D_ADDRESS")) Then
                    SearchResult(DataCount).D_ADDRESS = ""
                Else
                    SearchResult(DataCount).D_ADDRESS = SearchData("D_ADDRESS")
                End If

                'NO
                SearchResult(DataCount).NO = SearchData("NO")
                'ロット番号
                SearchResult(DataCount).LOT_NUMBER = SearchData("LOT_NUMBER")
                '保証書番号
                SearchResult(DataCount).WARRANTY_CARD_NUMBER = SearchData("WARRANTY_CARD_NUMBER")

                DataCount += 1
            Loop

            Data_Count = DataCount

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 請求書印刷データ検索
    ' <引数>
    ' DATE_FROM : 商品コード
    ' DATE_TO : 出荷倉庫
    ' PRT_STATUS1 : 通常出荷
    ' PRT_STATUS2 : 伝票のみ
    ' CLAIM_PRT_STATUS1 : 請求書出力済みデータ
    ' CLAIM_PRT_STATUS2 : 請求書未印刷データ
    ' <戻り値>
    ' SearchResult : 検索結果
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetClaimSeach(ByVal DATE_FROM As String, _
                                 ByVal DATE_TO As String, _
                                 ByVal CLAIM_CODE As String, _
                                 ByVal PRT_STATUS1 As String, _
                                 ByVal PRT_STATUS2 As String, _
                                 ByVal CLAIM_PRT_STATUS1 As String, _
                                 ByVal CLAIM_PRT_STATUS2 As String, _
                                 ByRef SearchResult() As Claim_List, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0
        Dim StatusWhere As String = Nothing

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = ""
            '検索条件よりWhereの作成

            '必須項目を検索条件に設定
            '出荷日FROM
            If DATE_FROM <> "" Then
                WhereSql &= " WHERE OUT_TBL.FIX_DATE >='"
                WhereSql &= DATE_FROM
                WhereSql &= "' "
            End If

            '出荷日TO
            If DATE_TO <> "" Then
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If

                WhereSql &= " OUT_TBL.FIX_DATE <='"
                WhereSql &= DATE_TO
                WhereSql &= "' "
            End If

            '請求先コード
            If CLAIM_CODE <> "" Then
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If
                WhereSql &= " M_CUSTOMER.CLAIM_CODE ='"
                WhereSql &= CLAIM_CODE
                WhereSql &= "' "
            End If


            '出力区分
            '種別の通常出荷、伝票のみ出力のどちらにもチェックが入っている
            If (PRT_STATUS1 = "True" And PRT_STATUS2 = "True") Or (PRT_STATUS1 = "False" And PRT_STATUS2 = "False") Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If
                WhereSql &= " OUT_TBL.STATUS = '出荷済み' OR OUT_TBL.STATUS = '伝票出力のみ' "
            ElseIf PRT_STATUS1 = "True" And PRT_STATUS2 = "False" Then
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If
                WhereSql &= "  OUT_TBL.STATUS = '出荷済み'"
            ElseIf PRT_STATUS1 = "False" And PRT_STATUS2 = "True" Then
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If
                WhereSql &= " OUT_TBL.STATUS = '伝票出力のみ'"
            End If

            '印刷済みデータの出力有無

            '種別の通常出荷、伝票のみ出力のどちらにもチェックが入っている
            If CLAIM_PRT_STATUS1 = "True" And CLAIM_PRT_STATUS2 = "True" Then
                '両方チェック、もしくは両方チェックなしは全件対象なので条件作成無し。
            ElseIf CLAIM_PRT_STATUS1 = "True" And CLAIM_PRT_STATUS2 = "False" Then
                'If WhereSql = "" Then
                '    WhereSql = " WHERE "
                'Else
                '    WhereSql &= " AND "
                'End If
                'WhereSql &= " OUT_TBL.CLAIM_PRT_DATE IS NULL "
            ElseIf CLAIM_PRT_STATUS1 = "False" And CLAIM_PRT_STATUS2 = "True" Then
                If WhereSql = "" Then
                    WhereSql = " WHERE "
                Else
                    WhereSql &= " AND "
                End If
                WhereSql &= " OUT_TBL.CLAIM_PRT_DATE IS NOT NULL"
            End If

            'オープン
            Connection.Open()

            'データ取得用コマンド作成（明細単位Ver）
            Command = Connection.CreateCommand
            Command.CommandText = "SELECT M_CUSTOMER.CLAIM_CODE,OUT_TBL.STATUS,OUT_TBL.SHEET_NO,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME,"
            Command.CommandText &= "M_ITEM.I_CODE,M_ITEM.I_NAME,OUT_TBL.FIX_NUM,OUT_TBL.UNIT_COST,OUT_TBL.CLAIM_NO,"
            Command.CommandText &= "DATE_FORMAT(OUT_TBL.CLAIM_PRT_DATE, '%Y/%m/%d') AS CLAIM_PRT_DATE,OUT_TBL.FILE_NAME,"
            Command.CommandText &= "OUT_TBL.ORDER_NO,DATE_FORMAT(OUT_TBL.FIX_DATE, '%Y/%m/%d') AS FIX_DATE,M_PLACE.NAME AS P_NAME,"
            Command.CommandText &= "OUT_TBL.COMMENT1,OUT_TBL.COMMENT2,OUT_TBL.ID,OUT_TBL.I_ID,OUT_TBL.C_ID "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_ITEM ON OUT_TBL.I_ID=M_ITEM.ID "
            Command.CommandText &= "INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID "
            Command.CommandText &= "INNER JOIN M_PLACE ON OUT_TBL.PLACE_ID=M_PLACE.ID "
            Command.CommandText &= WhereSql
            Command.CommandText &= " ORDER BY OUT_TBL.ID"

            'Select実行
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)
                '請求先コード
                SearchResult(Count).CLAIM_CODE = SearchData("CLAIM_CODE")
                'ステータス
                SearchResult(Count).STATUS = SearchData("STATUS")
                '伝票番号
                SearchResult(Count).SHEET_NO = SearchData("SHEET_NO")
                '納品先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '納品先名
                SearchResult(Count).D_NAME = SearchData("D_NAME")
                '商品コード
                SearchResult(Count).I_CODE = SearchData("I_CODE")
                '商品名
                SearchResult(Count).I_NAME = SearchData("I_NAME")
                '出荷数量
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '納品単価
                SearchResult(Count).UNIT_COST = SearchData("UNIT_COST")
                '出荷指示ファイル
                SearchResult(Count).FILE_NAME = SearchData("FILE_NAME")

                '請求書印刷日
                If IsDBNull(SearchData("CLAIM_PRT_DATE")) Then
                    SearchResult(Count).CLAIM_PRT_DATE = ""
                Else
                    SearchResult(Count).CLAIM_PRT_DATE = SearchData("CLAIM_PRT_DATE")
                End If
                'オーダー番号
                SearchResult(Count).ORDER_NO = SearchData("ORDER_NO")
                '出荷日
                If IsDBNull(SearchData("FIX_DATE")) Then
                    SearchResult(Count).FIX_DATE = ""
                Else
                    SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                End If

                '出荷倉庫
                SearchResult(Count).P_NAME = SearchData("P_NAME")
                'コメント１
                If IsDBNull(SearchData("COMMENT1")) Then
                    SearchResult(Count).COMMENT1 = ""
                Else
                    SearchResult(Count).COMMENT1 = SearchData("COMMENT1")
                End If
                'コメント２
                If IsDBNull(SearchData("COMMENT2")) Then
                    SearchResult(Count).COMMENT2 = ""
                Else
                    SearchResult(Count).COMMENT2 = SearchData("COMMENT2")
                End If
                'OUT_TBL.ID
                SearchResult(Count).ID = SearchData("ID")
                'OU_TBL.I_ID
                SearchResult(Count).I_ID = SearchData("I_ID")
                'OUT_TBL.C_ID
                SearchResult(Count).C_ID = SearchData("C_ID")
                '請求書番号
                SearchResult(Count).CLAIM_NO = SearchData("CLAIM_NO")
                Count += 1
            Loop

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 納品単価を更新する
    '
    ' <引数>
    ' Upd_Data : 変更データ格納配列
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function UpdOutTbl_UnitCost(ByRef Upd_Data() As Claim_List, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'データ件数ループ
            For Count = 0 To Upd_Data.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_TBL SET UNIT_COST="
                Command.CommandText &= Upd_Data(Count).UNIT_COST
                Command.CommandText &= " WHERE ID="
                Command.CommandText &= Upd_Data(Count).ID
                Command.CommandText &= ";"

                'UPDATE実行
                Command.ExecuteNonQuery()

            Next

            'OUT_TBLテーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 請求書印刷データ取得
    ' <引数>
    ' ClaimCode : 請求先コード
    ' IDList : DataGridでチェックをつけられたデータのOUT_TBL.IDの配列データ
    ' <戻り値>
    ' SearchResult : 検索結果
    ' ResultDataCount : 件数を返す
    ' Total : 合計金額を返す（各行の納入単価*数量のサマリー）
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetClaimPrtData(ByVal ClaimCode As String, _
                                  ByVal IDList() As Claim_List, _
                                 ByRef SearchResult() As Claim_List, _
                                 ByRef ResultDataCount As Integer, _
                                 ByRef Total As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            WhereSql = " WHERE M_CUSTOMER.CLAIM_CODE ='"
            WhereSql &= ClaimCode
            WhereSql &= "' "
            For i = 0 To IDList.Length - 1
                If i = 0 Then
                    WhereSql &= " AND OUT_TBL.ID in (" & IDList(i).ID
                Else
                    WhereSql &= "," & IDList(i).ID
                End If
            Next
            WhereSql &= ") "

            'オープン
            Connection.Open()

            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,M_CUSTOMER.CLAIM_CODE,M_CUSTOMER.C_NAME,OUT_TBL.FIX_DATE,OUT_TBL.SHEET_NO,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME, "
            Command.CommandText &= "OUT_TBL.FIX_NUM,OUT_TBL.UNIT_COST "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID  "
            Command.CommandText &= WhereSql
            Command.CommandText &= "ORDER BY OUT_TBL.FIX_DATE,OUT_TBL.SHEET_NO,OUT_TBL.ID "

            'Select実行
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)

                'OUT_TBL.ID
                SearchResult(Count).ID = SearchData("ID")
                '請求先コード
                SearchResult(Count).CLAIM_CODE = SearchData("CLAIM_CODE")
                '請求社名
                SearchResult(Count).C_NAME = SearchData("C_NAME")
                '発送日（出荷確定日）
                SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                '伝票番号
                SearchResult(Count).SHEET_NO = SearchData("SHEET_NO")
                '納品先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '納品先名
                SearchResult(Count).D_NAME = SearchData("D_NAME")
                '数量
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '納品単価
                SearchResult(Count).UNIT_COST = SearchData("UNIT_COST")

                Total += SearchData("FIX_NUM") * SearchData("UNIT_COST")

                Count += 1
            Loop

            ResultDataCount = Count

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "請求先データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 請求書印刷したデータに対して請求書印刷日、請求書番号を更新する
    '
    ' <引数>
    ' Claim_NO : 請求書番号
    ' Claim_DATE : 請求書印刷日（発行日）
    ' Upd_Data : 更新対象データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function UpdOutTbl_ClaimDate(ByRef Claim_NO As String, _
                                        ByRef Claim_DATE As String, _
                                        ByRef Upd_Data() As Claim_List, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'データ件数ループ
            For Count = 0 To Upd_Data.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_TBL SET CLAIM_PRT_DATE='"
                Command.CommandText &= Claim_DATE
                Command.CommandText &= "', CLAIM_NO='"
                Command.CommandText &= Claim_No
                Command.CommandText &= "' WHERE ID = "
                Command.CommandText &= Upd_Data(Count).ID
                Command.CommandText &= ";"

                'UPDATE実行
                Command.ExecuteNonQuery()

            Next

            'OUT_TBLテーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 請求書再印刷データ取得
    ' <引数>
    ' ClaimNO : 請求書番号
    ' <戻り値>
    ' SearchResult : 検索結果
    ' ResultDataCount : 件数を返す
    ' Total : 合計金額を返す（各行の納入単価*数量のサマリー）
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function GetReClaimPrtData(ByVal ClaimNo As String, _
                                 ByRef SearchResult() As Claim_List, _
                                 ByRef ResultDataCount As Integer, _
                                 ByRef Total As Integer, _
                                 ByRef Result As String, _
                                 ByRef ErrorMessage As String) As Boolean

        Dim SearchData As MySqlDataReader
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim WhereSql As String = Nothing
        Dim Count As Integer = 0

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring

            'オープン
            Connection.Open()

            Command = Connection.CreateCommand
            Command.CommandText = "SELECT OUT_TBL.ID,M_CUSTOMER.CLAIM_CODE,M_CUSTOMER.C_NAME,OUT_TBL.FIX_DATE,OUT_TBL.SHEET_NO,M_CUSTOMER.C_CODE,M_CUSTOMER.D_NAME, "
            Command.CommandText &= "OUT_TBL.FIX_NUM,OUT_TBL.UNIT_COST,DATE_FORMAT(OUT_TBL.CLAIM_PRT_DATE, '%Y/%m/%d') AS CLAIM_PRT_DATE,OUT_TBL.CLAIM_NO "
            Command.CommandText &= "FROM (OUT_TBL) INNER JOIN M_CUSTOMER ON OUT_TBL.C_ID=M_CUSTOMER.ID  "
            Command.CommandText &= "WHERE OUT_TBL.CLAIM_NO='"
            Command.CommandText &= ClaimNo
            Command.CommandText &= "' ORDER BY OUT_TBL.FIX_DATE,OUT_TBL.SHEET_NO,OUT_TBL.ID "

            'Select実行
            SearchData = Command.ExecuteReader()

            Do While (SearchData.Read)
                ReDim Preserve SearchResult(0 To Count)

                'OUT_TBL.ID
                SearchResult(Count).ID = SearchData("ID")
                '請求先コード
                SearchResult(Count).CLAIM_CODE = SearchData("CLAIM_CODE")
                '請求社名
                SearchResult(Count).C_NAME = SearchData("C_NAME")
                '発送日（出荷確定日）
                SearchResult(Count).FIX_DATE = SearchData("FIX_DATE")
                '伝票番号
                SearchResult(Count).SHEET_NO = SearchData("SHEET_NO")
                '納品先コード
                SearchResult(Count).C_CODE = SearchData("C_CODE")
                '納品先名
                SearchResult(Count).D_NAME = SearchData("D_NAME")
                '数量
                SearchResult(Count).FIX_NUM = SearchData("FIX_NUM")
                '納品単価
                SearchResult(Count).UNIT_COST = SearchData("UNIT_COST")
                '請求書印刷日
                SearchResult(Count).CLAIM_PRT_DATE = SearchData("CLAIM_PRT_DATE")
                '請求書番号
                SearchResult(Count).CLAIM_NO = SearchData("CLAIM_NO")

                Total += SearchData("FIX_NUM") * SearchData("UNIT_COST")

                Count += 1
            Loop

            ResultDataCount = Count

            'Countが0の場合、データが0件ということなのでメッセージを入れてReturn
            If Count = 0 Then
                ErrorMessage = "請求先データがみつかりません。"
                Result = False
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

    '***********************************************
    ' 問合せNo更新
    '
    ' <引数>
    ' Upd_Data : 更新対象データ
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function UpdOutTbl_INQUIRY_NO(ByRef Upd_Data() As Out_Search_List, _
                                   ByRef Result As String, _
                                   ByRef ErrorMessage As String) As Boolean
        Dim Tran As MySqlTransaction = Nothing
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand = Nothing
        Dim Count As Integer

        Try
            '接続文字列を設定
            Connection.ConnectionString = Constring
            'オープン
            Connection.Open()
            'begin
            Tran = Connection.BeginTransaction

            'データ件数ループ
            For Count = 0 To Upd_Data.Length - 1

                Command = Connection.CreateCommand
                Command.CommandText = "UPDATE OUT_TBL SET INQUIRY_NO='"
                Command.CommandText &= Upd_Data(Count).INQUIRY_NO
                Command.CommandText &= "' WHERE ID = "
                Command.CommandText &= Upd_Data(Count).ID
                Command.CommandText &= ";"

                'UPDATE実行
                Command.ExecuteNonQuery()

            Next

            'OUT_TBLテーブルに全てUPDATEが完了したらコミットを行う。
            Tran.Commit()
        Catch ex As Exception
            'エラーが発生した場合、ロールバックを行う。
            Tran.Rollback()
            ErrorMessage = ex.Message
            Result = False
        Finally
            If Connection IsNot Nothing Then
                Connection.Close()
                Connection.Dispose()
            End If
            If Command IsNot Nothing Then Command.Dispose()
        End Try
        Return Result
    End Function

End Module
