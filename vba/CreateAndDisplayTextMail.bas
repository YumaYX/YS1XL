'######### CreateAndDisplayTextMail
'========================================
' 新規メール作成関数（返り値なし）
'----------------------------------------
' 引数:
'   toAddr  - 宛先 (カンマ区切りでも可)
'   ccAddr  - CC (省略可)
'   bccAddr - BCC (省略可)
'   subjTxt - タイトル
'   bodyTxt - 本文
'========================================
Sub CreateAndDisplayTextMail(toAddr As String, _
                             Optional ccAddr As String = "", _
                             Optional bccAddr As String = "", _
                             Optional subjTxt As String = "", _
                             Optional bodyTxt As String = "")
    On Error Resume Next

    ' Outlook アプリ生成
    Dim olApp As Object: Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")

    On Error GoTo 0

    ' 新規メール作成
    Dim mail As Object: Set mail = olApp.CreateItem(0) ' 0 = olMailItem
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

