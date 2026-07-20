
'######### CountValues
'========================================
' 列内の値の出現回数を集計
'----------------------------------------
' 引数:
'   ws  - Worksheetオブジェクト
'   col - 列番号（省略時: A列）
'
' 戻り値:
'   Dictionary
'     Key  : セルの値
'     Item : 出現回数
'========================================

Function CountValues(ws As Worksheet, Optional col As Long = 1) As Object

    Dim lastRow As Long
    Dim c As Range
    Dim d As Object

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Set d = CreateObject("Scripting.Dictionary")

    For Each c In ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
        d(c.Value) = d(c.Value) + 1
    Next

    Set CountValues = d

End Function

# CountValues

- 目的: 指定したワークシートの指定列について、各値の出現回数を集計し、`Dictionary` として返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 必須。集計対象のワークシートオブジェクト。 |
| col | Long | オプション。集計対象の列番号。省略時は `1`（A列）。 |

# 出力（日本語）

- 型: `Object`（`Scripting.Dictionary`）
- 内容: 指定列の値ごとの出現回数を格納した `Dictionary` を返します。

| Key | Item |
|-----|------|
| セルの値 | その値の出現回数 |

# 使用例

```vb
Dim d As Object
Dim k As Variant

Set d = CountValues(Worksheets("Sheet1"))

For Each k In d.Keys
    If d(k) > 1 Then
        Debug.Print k
    End If
Next
```

上記の例では、`Sheet1` のA列で2回以上出現する値をイミディエイトウィンドウに出力します。

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


# CreateAndDisplayTextMail

- 目的: Outlookを使用して、指定された宛先、件名、本文を含む新しいメールを作成し、それを画面に表示します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| toAddr | String | 必須。メールの主要な宛先アドレス（カンマ区切りで複数指定可）。 |
| ccAddr | String | オプション。CCに入れるアドレス（省略可能）。 |
| bccAddr | String | オプション。BCCに入れるアドレス（省略可能）。 |
| subjTxt | String | オプション。メールの件名（タイトル）。 |
| bodyTxt | String | オプション。メールの本文。 |

# 出力（日本語）

- 型: なし (Void)
- 内容: 新しく作成されたメールアイテムがOutlook画面に表示されます。

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


# ExportColumnToFile

- 目的: 指定したワークシートの列データをテキストファイルとして書き出します。

# 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | データを取り出す対象のワークシートを指定します。 |
| filePath | String | テキストファイルとして保存する先のフルパスを指定します。 |
| colNum | Long | 書き出す列の番号を指定します（省略時は1列目）。 |
| delimiter | String | セルの値を繋ぐ区切り文字を指定します（省略時は改行）。 |

# 出力

- 型: Sub
- 内容: 指定されたファイルパスに、指定列のデータ内容をテキストファイルとして書き出します。

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


# GetColumnByHeader

- 目的: 指定されたワークシートと見出し名に基づき、対応する列番号を検索して返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 検索対象となるワークシート。 |
| header | String | 検索したい見出し名（文字列）。 |
| rowNum | Long | 見出しが配置されている行番号（省略可、デフォルトは1行目）。 |

# 出力（日本語）

- 型: Long
- 内容: 見出しが見つかった場合はその列番号を返し、見つからなかった場合は0を返します。

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


# GetColumnValuesAsString

- 目的: 指定したワークシートの特定の列の全データを、指定された区切り文字で連結した一つの文字列として返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 値を読み取りたい対象のワークシートを指定します。 |
| colNum | Long | 値を取得したい列番号を指定します。（省略時は1列目） |
| delimiter | String | 各セルの値をつなぐための区切り文字を指定します。（省略時は改行コード） |

# 出力（日本語）

- 型: String
- 内容: 指定された列の全ての値を、指定された区切り文字でつないだ単一の文字列。

'######### GetTimestamp
Function GetTimestamp() As String
    ' yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す
    GetTimestamp = Format(Now, "yyyy-mm-dd-HH-MM-ss")
End Function


# GetTimestamp

- 目的: 現在の日時を「yyyy-mm-dd-HH-MM-ss」形式の文字列として取得する。

# 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| なし | - | 引数なし |

# 出力

- 型: String
- 内容: 現在のシステム日時を示す文字列（例: 2023-10-27-10-30-45）。

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


# GetValueByID

- 目的: 指定したIDをキーとして、ワークシート内の対応する値（ターゲット列の値）を検索し、取得します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 処理対象とするワークシートを指定します。 |
| idHeader | String | IDが記載されている列の見出し名（ヘッダー名）を指定します。 |
| idValue | Variant | 検索したいIDの値（検索キー）を指定します。 |
| targetHeader | String | 取得したい値が記載されている列の見出し名（ヘッダー名）を指定します。 |
| headerRow | Long | 見出し行が記載されている行番号です。省略した場合（=1）は1行目として扱われます。 |

# 出力（日本語）

- 型: Variant
- 内容: 検索したIDと一致した行の、ターゲット列（指定した値の列）のセル値が返されます。該当するIDが見つからない場合や、指定されたヘッダー列が存在しない場合は空文字（""）が返されます。

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


# GetValueByID_Hash

- Purpose: Retrieves a specific value from a target column by searching for a matching ID in a designated ID column.

# Inputs

| Argument Name | Type | Description |
|---|---|---|
| ws | Worksheet | The target worksheet containing the data. |
| idHeader | String | The header name of the column containing the IDs. |
| idValue | Variant | The specific ID value that needs to be searched for. |
| targetHeader | String | The header name of the column from which the value should be retrieved. |
| headerRow | Long | (Optional) The row number where the headers are located (defaults to 1). |

# Output

- Type: Variant
- Content: The value found in the target column corresponding to the matching ID; returns an empty string if no match is found.

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


# GlobCollection

- 目的: Returns a collection containing the full paths of all files that match a specified pattern within a given directory.

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| folderPath | String | The full path of the directory to search within. |
| pattern | String | The file pattern (e.g., "*.txt" or "report*.csv") to match. |

# 出力（日本語）

- 型: Collection
- 内容: A collection object where each item is the complete file path of a file found matching the pattern.

'######### Hello
Function Hello() As String
    hello = "hello"
End Function


# Hello

- 目的: この関数は、固定の文字列 "hello" を返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| None   | N/A | 引数は不要です。 |

# 出力（日本語）

- 型: String
- 内容: "hello" という文字列。

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


# IsValidIPAddress

- 目的: 与えられた文字列が有効なIPv4アドレス形式であるかどうかを検証します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ip | String | 検証対象のIPアドレス文字列。 |

# 出力（日本語）

- 型: Boolean
- 内容: 有効なIPv4アドレスであればTrue、そうでなければFalseを返します。

# IsValidSubnetMask

- 目的: 与えられた文字列が、連続した1と0からなる有効なサブネットマスク形式であるかどうかを検証します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| mask | String | 検証対象のサブネットマスク文字列。 |

# 出力（日本語）

- 型: Boolean
- 内容: 有効なサブネットマスクであればTrue、そうでなければFalseを返します。

# IsValidNetworkAddress

- 目的: 指定されたIPアドレスとサブネットマスクの組み合わせが、論理的に正しいネットワークアドレスを構成しているかを検証します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ip | String | 検証対象のIPアドレス文字列。 |
| mask | String | サブネットマスク文字列。 |

# 出力（日本語）

- 型: Boolean
- 内容: 有効なネットワークアドレスであればTrue、そうでなければFalseを返します。

# CIDR2Mask

- 目的: CIDR表記の整数（例: 24）を、ドット区切りのサブネットマスク文字列に変換します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| cidr | Integer | サブネットのCIDRプレフィックス長（例: 8, 16, 24）。 |

# 出力（日本語）

- 型: String
- 内容: 対応するドット区切りのサブネットマスク文字列（例: "255.255.255.0"）を返します。

# Mask2CIDR

- 目的: ドット区切りのサブネットマスク文字列を、対応するCIDR表記の整数値に変換します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| mask | String | 変換対象のサブネットマスク文字列。 |

# 出力（日本語）

- 型: Integer
- 内容: 対応するCIDRプレフィックス長（例: 24）を返します。

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


# LastUsedRow

- 目的: 指定されたワークシートの指定列における、データが入力されている最後の行番号を返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 処理を行うワークシートを指定します。 |
| col | Long | 調査する列番号を指定します。指定しない場合（省略時）はA列（1列目）が使用されます。 |

# 出力（日本語）

- 型: Long
- 内容: データが入力されている最後の行番号を返します。指定された列全体が空の場合、0を返します。

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


# OpenExcel

- 目的: ファイル選択ダイアログを開き、ユーザーが選択したExcelブックを新しいWorkbookオブジェクトとして開いて返します。

# 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| N/A | N/A | 実行時にファイル選択ダイアログを表示し、ユーザーにファイルを選択してもらいます。 |

# 出力

- 型: Workbook
- 内容: 開かれたワークブックオブジェクト。ユーザーがファイル選択をキャンセルした場合、Nothingを返します。

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


# ReadUtf8Text

- 目的: 指定されたファイルパスから、UTF-8エンコーディングでテキストデータを読み込む。

# 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| filePath | String | 読み込むファイルの完全なパス。 |

# 出力

- 型: String
- 内容: ファイルの内容全体をUTF-8としてデコードした文字列。

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



# SaveAttachments

- 目的: 指定されたメールフォルダ内の添付ファイルのうち、指定したキーワードを含むファイルを検索し、指定の保存先に保存する。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| mailFolderPath | String | 処理対象のメールフォルダのパス（ルートからのパス）。 |
| targetFolderPath | String | 添付ファイルを保存する先のフォルダのパス。 |
| keyword | String | 添付ファイルのファイル名に含まれるべきキーワード。 |

# 出力（日本語）

- 型: Sub (戻り値なし)
- 内容: 処理が完了したことを示す（値を返さない）。

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

# SearchDataLocation

- 目的: 指定されたCSVファイルから、特定のキー（列）に基づいて一意なデータ値とそれが出現した最初の行番号を辞書として取得します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| csvFilePath | String | 処理対象のCSVファイルのフルパス。指定しない場合は "sample.csv" が使用されます。 |
| targetKey | String | 辞書に格納するキーとして使用する列のヘッダー名（例: "id"）。 |

# 出力（日本語）

- 型: Scripting.Dictionary (Object)
- 内容: 指定されたキー列に含まれる一意な値（Key）を、その値がデータ内で最初に現れた行番号（Value）に対応づけた辞書オブジェクト。
