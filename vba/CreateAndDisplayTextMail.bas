' CreateAndDisplayTextMail "taro@example.com", , , "件名", "本文"
' CreateAndDisplayTextMail "a@b.com; c@d.com", , , "宛先は区切り可", ""
'========================================
' 新規メール作成関数（返り値なし）
'----------------------------------------
' 引数:
'   toAddr  - 宛先（カンマ・セミコロン区切り可）
'   ccAddr  - CC (省略可)
'   bccAddr - BCC (省略可)
'   subjTxt - タイトル
'   bodyTxt - 本文
' 注:
'   Outlook が利用できない環境では何もせず終了する
'========================================
Sub CreateAndDisplayTextMail(toAddr As String, _
                             Optional ccAddr As String = "", _
                             Optional bccAddr As String = "", _
                             Optional subjTxt As String = "", _
                             Optional bodyTxt As String = "")

    Dim olApp As Object
    Dim mail As Object

    ' Outlook アプリ生成
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Err.Clear
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    ' Outlook が起動できない場合は Object 変数が Nothing のままなので終了する
    If olApp Is Nothing Then Exit Sub

    ' 新規メール作成
    Set mail = olApp.CreateItem(0) ' 0 = olMailItem

    ' プロパティ設定
    With mail
        .To = toAddr
        .CC = ccAddr
        .BCC = bccAddr
        .BodyFormat = 1 ' 1 = olFormatPlain (テキスト形式)
        .Subject = subjTxt
        .Body = bodyTxt
        .Display  ' 作成したメールを表示
    End With
End Sub
