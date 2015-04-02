
Public Class LoginForm1

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Dim loginid As String
        Dim password As String
        Dim displist() As DispID_List = Nothing

        'Dim Connection As New MySqlConnection
        'Dim Command As MySqlCommand = Nothing
        'Dim DataReader As MySqlDataReader = Nothing
        Dim errorMessage As String = Nothing
        Dim result As Boolean = True

        Dim loginidCheckAfter As String = Nothing
        Dim passwordCheckAfter As String = Nothing
        Dim StringChkResult As Boolean = True
        Dim StringErrorMessage As String = Nothing

        'ID入力チェック
        If Trim(UsernameTextBox.Text) = "" Then
            MsgBox("ログインIDを入力して下さい。")
            Exit Sub
        End If

        'パスワード入力チェック
        If Trim(PasswordTextBox.Text) = "" Then
            MsgBox("パスワードを入力して下さい。")
            Exit Sub
        End If

        'IDの文字列の妥当性チェック
        If Trim(UsernameTextBox.Text) <> "" Then
            '文字の妥当性チェックを行う。
            'NullOK、
            If StringChkVal(Trim(UsernameTextBox.Text), True, False, loginidCheckAfter, StringChkResult, StringErrorMessage) = False Then
                UsernameTextBox.BackColor = Color.Salmon
                MsgBox("ログインIDに不正な文字が入力されています。")
            Else
                'チェックに問題がなければ背景色を白に戻す。
                UsernameTextBox.BackColor = Color.White
            End If
        End If
        StringChkResult = True
        StringErrorMessage = Nothing

        'パスワードの文字列の妥当性チェック
        If Trim(PasswordTextBox.Text) <> "" Then
            '文字の妥当性チェックを行う。
            'NullOK、
            If StringChkVal(Trim(PasswordTextBox.Text), True, False, passwordCheckAfter, StringChkResult, StringErrorMessage) = False Then
                PasswordTextBox.BackColor = Color.Salmon
                MsgBox("ログインIDに不正な文字が入力されています。")
            Else
                'チェックに問題がなければ背景色を白に戻す。
                PasswordTextBox.BackColor = Color.White
            End If
        End If


        loginid = Trim(loginidCheckAfter)
        password = Trim(passwordCheckAfter)
        'ログインチェックFunction
        'result = LoginCheck(constring, loginid, password, result, errorMessage)
        result = LoginCheck(loginid, password, displist, result, errorMessage)
        If result = "true" Then
            topmenu.Show()
            Me.Hide()
        ElseIf result = "false" Then
            MsgBox(errorMessage)
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '画面タイトルの設定
        Me.Text = SYSTEM_NAME & " - ログイン"
        'フォームの表示位置設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(DisplayX, DisplayY)
        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

        Dim insProcess As Process() 'プロセスインスタンス
        Dim strProcessName As String  'プロセス名

        'フォームの最大化を非表示にする。
        Me.MinimizeBox = Not Me.MinimizeBox

        strProcessName = Process.GetCurrentProcess.ProcessName
        insProcess = Process.GetProcessesByName(strProcessName)

        If UBound(insProcess) > 0 Then
            MsgBox("すでに起動しています")
            End
        End If

        UsernameTextBox.Focus()


    End Sub

    Private Sub LoginForm1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

End Class

