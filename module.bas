
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

'######### ExportColumnToFile
'========================================
' 指定列の値をテキストファイルに書き出す
' ws       : 対象ワークシート
' colNum   : 書き出す列番号（省略時1列目）
' filePath : 保存先フルパス
' delimiter: 値をつなぐ区切り（省略時改行）
'========================================
Sub ExportColumnToFile(ws As Worksheet, _
                       filePath As String, _
                       Optional colNum As Long = 1, _
                       Optional delimiter As String = vbCrLf)
    Dim content As String: content = GetColumnValuesAsString(ws, colNum, delimiter)
    ' ファイル書き出し
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
End Sub

'######### GetColumnByHeader
'========================================
' 見出し名から列番号を探す
' ws      : 対象ワークシート
' header  : 探したい見出し名
' rowNum  : 見出しがある行番号（通常1行目）
' 戻り値  : 列番号（見つからなければ0）
'========================================
Function GetColumnByHeader(ws As Worksheet, header As String, Optional rowNum As Long = 1) As Long
    Dim lastCol As Long: lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long: For c = 1 To lastCol
        GetColumnByHeader = c
        If ws.Cells(rowNum, c).Value = header Then Exit Function
    Next c
    GetColumnByHeader = 0 ' 見つからなければ0
End Function

'######### GetColumnValuesAsString
'========================================
' 指定列の値を文字列で返す
' ws        : 対象ワークシート
' colNum    : 取得する列番号（省略時1列目）
' delimiter : 値をつなぐ区切り（省略時 vbCrLf で改行）
' 戻り値    : 列の値をつなげた文字列
'========================================
Function GetColumnValuesAsString(ws As Worksheet, _
                                 Optional colNum As Long = 1, _
                                 Optional delimiter As String = vbCrLf) As String
    ' 最終行を取得
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    Dim result As String: result = ""

    Dim r As Long: For r = 1 To lastRow
        result = result & ws.Cells(r, colNum).Value & delimiter
    Next r
    GetColumnValuesAsString = result
End Function

'######### GetTimestamp
Function GetTimestamp() As String
    ' yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す
    GetTimestamp = Format(Now, "yyyy-mm-dd-HH-MM-ss")
End Function

'######### GetValueByID
'========================================
' IDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
'========================================
Function GetValueByID(ws As Worksheet, _
                             idHeader As String, _
                             idValue As Variant, _
                             targetHeader As String, _
                             Optional headerRow As Long = 1) As Variant
    GetValueByID = ""
    
    Dim idCol     As Long: idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    Dim targetCol As Long: targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function    

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row

    Dim r As Long: For r = headerRow + 1 To lastRow
        If Not IsError(ws.Cells(r, idCol).Value) Then ' エラー値回避
            If ws.Cells(r, idCol).Value = idValue Then
                GetValueByID = ws.Cells(r, targetCol).Value
                Exit Function
            End If
        End If
    Next r
End Function

'######### GetValueByID_Hash
'========================================
' ハッシュでIDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
'========================================
Function GetValueByID_Hash(ws As Worksheet, _
                           idHeader As String, _
                           idValue As Variant, _
                           targetHeader As String, _
                           Optional headerRow As Long = 1) As Variant
    GetValueByID_Hash = "" ' 見出しが見つからない
    
    ' ID列と取得列の列番号を取得
    Dim idCol     As Long: idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    Dim targetCol As Long: targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function    

    ' 最終行
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    Dim r As Long: For r = headerRow + 1 To lastRow
        If idValue = ws.Cells(r, idCol).Value Then
            GetValueByID_Hash = ws.Cells(r, targetCol).Value
            Exit Function
        End If
    Next r
End Function

'######### GlobCollection
Function GlobCollection(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fileName As String

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    fileName = Dir(folderPath & pattern)

    Do While fileName <> ""
        col.Add folderPath & fileName
        fileName = Dir()
    Loop

    Set GlobCollection = col
End Function

'######### Hello
Function Hello() As String
    hello = "hello"
End Function

'######### IPaddress
Function IsValidIPAddress(ByVal ip As String) As Boolean
    Dim parts() As String
    Dim i As Integer
    Dim num As Integer

    parts = Split(ip, ".")

    If UBound(parts) <> 3 Then Exit Function

    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function

        num = CInt(parts(i))
        If num < 0 Or num > 255 Then Exit Function

        ' 先頭ゼロ防止（例: 01）
        If parts(i) <> CStr(num) Then Exit Function
    Next i

    IsValidIPAddress = True
End Function

Function IsValidSubnetMask(ByVal mask As String) As Boolean
    Dim parts() As String
    Dim i As Integer
    Dim num As Integer
    Dim binaryStr As String

    parts = Split(mask, ".")
    If UBound(parts) <> 3 Then Exit Function

    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function

        num = CInt(parts(i))
        If num < 0 Or num > 255 Then Exit Function

        binaryStr = binaryStr & Right("00000000" & WorksheetFunction.Dec2Bin(num), 8)
    Next i

    ' 「1が続いた後に0が続く」パターンのみ許可
    If InStr(binaryStr, "01") > 0 Then Exit Function

    ' 全部1 or 全部0は除外（必要に応じて調整）
    If binaryStr = String(32, "1") Or binaryStr = String(32, "0") Then Exit Function

    IsValidSubnetMask = True
End Function

Function IsValidNetworkAddress(ByVal ip As String, ByVal mask As String) As Boolean
    Dim ipParts() As String
    Dim maskParts() As String
    Dim i As Integer

    If Not IsValidIPAddress(ip) Then Exit Function
    If Not IsValidSubnetMask(mask) Then Exit Function

    ipParts = Split(ip, ".")
    maskParts = Split(mask, ".")

    For i = 0 To 3
        If (CInt(ipParts(i)) And Not CInt(maskParts(i))) <> 0 Then
            Exit Function
        End If
    Next i

    IsValidNetworkAddress = True
End Function

Function CIDR2Mask(cidr As Integer) As String
    Dim i As Integer
    Dim mask(3) As Integer
    Dim bits As Integer

    bits = cidr

    For i = 0 To 3
        If bits >= 8 Then
            mask(i) = 255
            bits = bits - 8
        ElseIf bits > 0 Then
            mask(i) = 256 - 2 ^ (8 - bits)
            bits = 0
        Else
            mask(i) = 0
        End If
    Next i

    CIDR2Mask = mask(0) & "." & mask(1) & "." & mask(2) & "." & mask(3)
End Function

Function Mask2CIDR(mask As String) As Integer
    Dim parts() As String
    Dim i As Integer
    Dim val As Integer
    Dim cidr As Integer

    parts = Split(mask, ".")

    For i = 0 To 3
        val = CInt(parts(i))

        Do While val > 0
            cidr = cidr + (val And 1)
            val = val \ 2
        Loop
    Next i

    Mask2CIDR = cidr
End Function

'######### LastUsedRow
Function LastUsedRow(ws As Worksheet, Optional col As Long = 1) As Long
    With ws
        If Application.WorksheetFunction.CountA(.Columns(col)) = 0 Then
            LastUsedRow = 0
        Else
            LastUsedRow = .Cells(.Rows.Count, col).End(xlUp).Row
        End If
    End With
End Function

'######### OpenExcel
'# Dim wb As Workbook: Set wb = OpenExcel()

Function OpenExcel() As Workbook
    Dim filename As Variant
    filename = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xls*),*.xls*,CSVファイル (*.csv),*.csv")

    If filename = False Then
        Set OpenExcel = Nothing
        Exit Function
    End If

    Set OpenExcel = Workbooks.Open(filename)
End Function

'######### ReadUtf8Text
Function ReadUtf8Text(filePath As String) As String
    
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With
    
    Set stm = Nothing

End Function

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


'######### SearchDataLocation
Function StoreDataLocation(data As Variant, keyIndex As Long) As Object

    Dim dataLocation As Object
    Set dataLocation = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    Dim idx As Long

    For i = 1 To UBound(data, 1)
        key = CStr(data(i, keyIndex))
        dataLocation(key) = IIf(dataLocation.Exists(key), 1, i)
    Next i
    Set StoreDataLocation = dataLocation
End Function


Function SearchDataLocation(Optional csvFilePath As String = "sample.csv", _
                            Optional targetKey As String = "id") As Object

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dataArr As Variant

    Set wb = Workbooks.Open(csvFilePath)
    Set ws = wb.Sheets(1)

    dataArr = ws.UsedRange.Value

    Dim headersHash As Object
    Set headersHash = CreateObject("Scripting.Dictionary")

    Dim col As Long
    Dim lastCol As Long
    lastCol = UBound(dataArr, 2)

    For col = 1 To lastCol
        headersHash(CStr(dataArr(1, col))) = col
    Next col

    Dim keyColumnIndexNumber As Long
    keyColumnIndexNumber = headersHash(targetKey)

    Set SearchDataLocation = StoreDataLocation(dataArr, keyColumnIndexNumber)

    wb.Close False

End Function
