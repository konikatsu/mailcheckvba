Sub CheckMailMonthly()
    ' メールサーバーの設定 (例: xserver POP3 with SSL/TLS)
    Dim conf As CDO.Configuration
    Set conf = New CDO.Configuration

    conf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receiveserver") = "svxxxx.xserver.jp"
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receiveserverport") = 995 ' POP3 over SSL
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receiveserversecureauth") = 2 ' cdoSecureSocketLayer
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receivesmtpserver") = "smtp.xserver.jp" ' SMTPサーバー (送信時に必要)
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receivesmtpserverport") = 465 ' SMTPサーバーのポート番号 (SSL)
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receivesmtpauthenticate") = 1 ' cdoBasic
    conf.Fields("http://schemas.microsoft.com/cdo/configuration/receivesendusername") = "your_email@example.com"
    

    ' メールセッションの作成
    Dim sess As CDO.Session
    Set sess = New CDO.Session
    sess.Configuration = conf

    ' インボックスを取得
    Dim folder As CDO.Folder
    Set folder = sess.GetDefaultFolder(olFolderInbox)

    ' メールアイテムを取得
    Dim items As CDO.Items
    Set items = folder.Items

    ' 当日の日付を取得
    Dim today As Date
    today = Date

    ' 特定の件名で、かつ当日受信のメールを検索
    Dim item As Object
    For Each item In items
        ' ヘッダーから送信日時を取得 (例: Dateフィールド)
        Dim headers As CDO.Fields
        Set headers = item.Headers
        Dim sentOnString As String
        sentOnString = headers.Item("Date")

        ' 送信日時をDate型に変換 (日付形式に合わせて調整)
        Dim sentOn As Date
        ' 例: 年月日がYYYY-MM-DD形式の場合
        sentOn = CDate(Format(sentOnString, "yyyy-mm-dd"))

        If InStr(1, item.Subject, "特定の件名", vbTextCompare) > 0 And _
           item.ReceivedTime >= today And item.ReceivedTime < DateAdd("d", 1, today) And _
           sentOn >= today And sentOn < DateAdd("d", 1, today) Then
            ' メールが見つかった場合の処理
            MsgBox "件名: " & item.Subject & vbCrLf & _
                   "受信日時: " & item.ReceivedTime & vbCrLf & _
                   "送信日時: " & sentOn
            ' データベースに記録する処理などをここに記述
        End If

        Set headers = Nothing
    Next item

    ' オブジェクトの解放
    Set item = Nothing
    Set items = Nothing
    Set folder = Nothing
    Set sess = Nothing
    Set conf = Nothing
End Sub