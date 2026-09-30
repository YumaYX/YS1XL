' Call SaveAttachments("アーカイブ", "C:\Temp", "キーワード")
'
' 第1引数：メールフォルダ（ルートからのパス）
' 第2引数：保存先フォルダ
' 第3引数：添付ファイル名のキーワード
' 注：
'   キーワードが空文字の場合は InStr が常に 1 となり
'   全添付ファイルを対象にしてしまうため何も保存せず終了する
'   Outlook が利用できない環境、または保存先フォルダ /
'   指定メールフォルダが存在しない場合は何もせず終了する
Public Sub SaveAttachments( _
    ByVal mailFolderPath As String, _
    ByVal targetFolderPath As String, _
    ByVal keyword As String)

    Dim olApp As Object
    Dim olNamespace As Object
    Dim olRoot As Object
    Dim olFolder As Object
    Dim olItems As Object
    Dim olMail As Object
    Dim olAttachment As Object
    Dim fso As Object
    Dim folders() As String
    Dim k As Long
    Dim i As Long
    Dim j As Long
    Dim savePath As String

    ' キーワード未指定時は何もしない
    If Len(keyword) = 0 Then Exit Sub

    ' 保存先フォルダの末尾に "\" があるときは連結時の重複を防ぐため取り除く
    Do While Right(targetFolderPath, 1) = "\"
        targetFolderPath = Left(targetFolderPath, Len(targetFolderPath) - 1)
    Loop
    If Len(targetFolderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(targetFolderPath) Then Exit Sub

    ' Outlook取得（既存優先）
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Err.Clear
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If olApp Is Nothing Then Exit Sub

    Set olNamespace = olApp.GetNamespace("MAPI")

    ' ルート取得
    Set olRoot = olNamespace.GetDefaultFolder(6).Parent

    ' フォルダ解決
    Set olFolder = olRoot

    folders = Split(mailFolderPath, "\")

    For k = 0 To UBound(folders)
        If folders(k) <> "" Then
            ' 存在しないフォルダを指定した場合はエラーになるため打ち切る
            On Error Resume Next
            Set olFolder = olFolder.Folders(folders(k))
            On Error GoTo 0
            If olFolder Is Nothing Then Exit Sub
        End If
    Next k

    ' メール処理（Items を変数に固定し、各アクセスで再取得しない）
    Set olItems = olFolder.Items

    For i = 1 To olItems.Count

        If TypeName(olItems(i)) = "MailItem" Then

            Set olMail = olItems(i)

            If olMail.Attachments.Count > 0 Then

                For j = 1 To olMail.Attachments.Count

                    Set olAttachment = olMail.Attachments(j)

                    If InStr(LCase(olAttachment.FileName), LCase(keyword)) > 0 Then

                        savePath = targetFolderPath & "\" & olAttachment.FileName

                        ' 同名の既存ファイルは上書きせずスキップする
                        If Not fso.FileExists(savePath) Then
                            olAttachment.SaveAsFile savePath
                        End If

                    End If

                Next j
            End If
        End If
    Next i
End Sub
