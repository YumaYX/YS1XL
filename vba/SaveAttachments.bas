'######### SaveAttachments
' Call SaveAttachments("アーカイブ", "C:\Temp", "キーワード")
'
' 第1引数：メールフォルダ（ルートからのパス）
' 第2引数：保存先フォルダ
' 第3引数：添付ファイル名のキーワード
Public Sub SaveAttachments( _
    ByVal mailFolderPath As String, _
    ByVal targetFolderPath As String, _
    ByVal keyword As String)

    ' Outlook取得（既存優先）
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If

    Dim olNamespace As Object
    Set olNamespace = olApp.GetNamespace("MAPI")

    ' ルート取得
    Dim olRoot As Object
    Set olRoot = olNamespace.GetDefaultFolder(6).Parent

    ' フォルダ解決
    Dim olFolder As Object
    Set olFolder = olRoot

    Dim folders() As String
    folders = Split(mailFolderPath, "\")

    Dim k As Long
    For k = 0 To UBound(folders)
        If folders(k) <> "" Then
            Set olFolder = olFolder.Folders(folders(k))
        End If
    Next k

    ' メール処理
    Dim i As Long
    For i = 1 To olFolder.Items.Count
        
        If TypeName(olFolder.Items(i)) = "MailItem" Then
            
            Dim olMail As Object
            Set olMail = olFolder.Items(i)

            If olMail.Attachments.Count > 0 Then
                
                Dim j As Long
                For j = 1 To olMail.Attachments.Count
                    
                    Dim olAttachment As Object
                    Set olAttachment = olMail.Attachments(j)

                    If InStr(LCase(olAttachment.FileName), LCase(keyword)) > 0 Then
                        
                        Dim savePath As String
                        savePath = targetFolderPath & "\" & olAttachment.FileName

                        If Dir(savePath) = "" Then
                            olAttachment.SaveAsFile savePath
                        End If

                    End If

                Next j
            End If
        End If
    Next i
End Sub


