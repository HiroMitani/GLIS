Imports System.Text.RegularExpressions

Module CHKMOB
    '***********************************************
    ' 数値入力チェックを行う。
    ' <引数>
    ' ChkNum : チェックをするデータを格納
    ' ChkType : チェックするデータ型を指定（INTEGER:整数数値型、DOUBLE:浮動小数点型）
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' ZeroChk : 入力したテキストボックスで0の値を許可するか（TRUE:許可、FALSE:却下）
    ' MinusChk : 入力したテキストボックスでマイナスの値を許可するか(TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function NumChkVal(ByVal ChkNum As String, _
                              ByVal ChkType As String, _
                              ByVal NullChk As Boolean, _
                              ByVal ZeroChk As Boolean, _
                              ByVal MinusChk As Boolean, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        Dim CheckInteger As Integer
        Dim CheckDouble As Double

        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkNum = "" Then
                ErrorMessage = "数量を入力してください。"
                Return False
                Exit Function
            End If
        End If

        '渡されてきた値の入力チェックタイプを判別。
        If ChkType = "INTEGER" Then
            '渡されてきた値はINT型として正しいかチェック。
            If Int32.TryParse(ChkNum, CheckInteger) Then
            Else
                ErrorMessage = "半角整数の値しか入力できません。"
                Return False
                Exit Function
            End If

            '渡されてきた値は0がOKか。
            If ZeroChk = True Then
            Else
                If ChkNum = 0 Then
                    ErrorMessage = "1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '渡されてきた値はマイナス値がOKか。
            If MinusChk = True Then
            Else
                If ChkNum < 0 Then
                    ErrorMessage = "1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '小数点でもOKの場合のチェックはDOUBLE型へ
        ElseIf ChkType = "DOUBLE" Then
            '渡されてきた値はINT型として正しいかチェック。
            If Double.TryParse(ChkNum, CheckDouble) Then
            Else
                ErrorMessage = "半角数値の値しか入力できません。"
                Return False
                Exit Function
            End If

            '渡されてきた値は0がOKか。
            If ZeroChk = True Then
            Else
                If ChkNum = 0 Then
                    ErrorMessage = "数値は1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '渡されてきた値はマイナス値がOKか。
            If MinusChk = True Then
            Else
                If ChkNum < 0 Then
                    ErrorMessage = "数値は1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If
        End If
        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 日付入力チェックを行う。
    ' <引数>
    ' ChkDate : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' ChkDateAfter : 変換後のデータを格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function DateChkVal(ByVal ChkDate As String, _
                              ByVal NullChk As Boolean, _
                              ByRef ChkDateAfter As String, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        Dim CheckDate As Date

        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkDate = "" Then
                ErrorMessage = "日付をYYYY/MM/DD形式で入力してください。"
                Return False
                Exit Function
            End If
        End If

        '渡されてきた値はDATE型として正しいかチェック。
        If DateTime.TryParse(ChkDate, CheckDate) Then
            ChkDateAfter = CheckDate.ToString("yyyy/MM/dd")
        Else
            ErrorMessage = "入力された値は日付として正しくありません。"
            Return False
            Exit Function
        End If

        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 文字入力チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' SingleQuotationChk : シングルコーテーションを許可するか（TRUE:許可、FALSE:却下）　却下の場合は''に置き換え
    ' <戻り値>
    ' ChkStringAfter : 処理後の文字列を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function StringChkVal(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByVal SingleQuotationChk As Boolean, _
                              ByRef ChkStringAfter As String, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean


        Dim i As Integer

        ChkStringAfter = ChkString
        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkString = "" Then
                ErrorMessage = "文字を入力してください。"
                Return False
                Exit Function
            End If
        End If

        '渡されてきた値はシングルコーテーションを許可するか。許可しない場合は変換を行う。
        If SingleQuotationChk = True Then
        Else
            i = InStr(ChkString, "'")
            '文字列の中にシングルコーテーションがあった場合はシングルコーテーションを２つ''に変換する。
            If i > 0 Then
                ChkStringAfter = ChkString.Replace("'", "''")
            End If
        End If

        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 郵便番号チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function PostChkVal(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        'NULLがOKか。
        If NullChk = True Then
        Else
            If ChkString = "" Then
                ' 入力エラー
                ErrorMessage = "郵便番号が入力されていません。"
                Return False
            End If
        End If

        If Regex.IsMatch(ChkString, "^[0-9]{3}[\-]?[0-9]{4}$") Then
            Return True
        Else
            If ChkString = "" Then
                Return True

            Else
                ' 入力エラー
                ErrorMessage = "郵便番号の形式が正しくありません。"
                Return False
            End If

        End If
    End Function

    '***********************************************
    ' 郵便番号チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function PlaceChkVal(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        'NULLがOKか。
        If NullChk = True Then
        Else
            If ChkString = "" Then
                ' 入力エラー
                ErrorMessage = "郵便番号が入力されていません。"
                Return False
            End If
        End If

        If Regex.IsMatch(ChkString, "^[0-9]{3}[\-]?[0-9]{4}$") Then
            Return True
        Else
            If ChkString = "" Then
                Return True

            Else
                ' 入力エラー
                ErrorMessage = "郵便番号の形式が正しくありません。"
                Return False
            End If

        End If
    End Function
End Module
