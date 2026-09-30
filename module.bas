'######### AutoGui
' マウス操作用の Win32 API 宣言
'
' 使用例:
'   SetCursorPos 100, 100
'   mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'   mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'
' ユーザー定義型は利用箇所より前に定義しておく必要がある
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const MOUSEEVENTF_LEFTDOWN  As Long = &H2
Private Const MOUSEEVENTF_LEFTUP    As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

Private Declare PtrSafe Function SetCursorPos Lib "user32" ( _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Private Declare PtrSafe Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, _
    ByVal dx As Long, _
    ByVal dy As Long, _
    ByVal dwData As Long, _
    ByVal dwExtraInfo As LongPtr _
)

Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI _
) As Long
'######### ControlFile
' If Cp("C:\src\a.txt", "C:\dst\a.txt") Then Debug.Print "コピー成功"
' If Cp("C:\src\folder", "C:\dst\folder") Then Debug.Print "フォルダも自動判定"

'==========================================================
' FileSystemObject（ファイルシステム操作オブジェクト）
'==========================================================
Private Function GetFso() As Object

    Static fso As Object

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    Set GetFso = fso

End Function


'==========================================================
' Pathの種類を判定
'
' 引数:
'   targetPath - 判定対象のパス（ファイル/フォルダ）
'
' 戻り値
'   "FILE" : ファイル
'   "DIR"  : フォルダ
'   ""     : 存在しない
'==========================================================
Private Function PathType(ByVal targetPath As String) As String

    Dim fso As Object

    Set fso = GetFso()

    If fso.FileExists(targetPath) Then
        PathType = "FILE"

    ElseIf fso.FolderExists(targetPath) Then
        PathType = "DIR"

    Else
        PathType = ""

    End If

End Function


'==========================================================
' Cp（copy：コピー）
'
' ファイル / フォルダを自動判定
'
' 引数:
'   sourcePath      - コピー元のパス（ファイル/フォルダを自動判定）
'   destinationPath - コピー先のパス
'   overwrite(省略可)- 上書きの可否（省略時: False）
'
' 戻り値:
'   True  : 成功
'   False : 失敗
'==========================================================
Public Function Cp( _
    ByVal sourcePath As String, _
    ByVal destinationPath As String, _
    Optional ByVal overwrite As Boolean = False _
) As Boolean

    Dim fso As Object

    On Error GoTo ErrorHandler

    Set fso = GetFso()

    Select Case PathType(sourcePath)

        Case "FILE"

            fso.CopyFile _
                sourcePath, _
                destinationPath, _
                overwrite

        Case "DIR"

            fso.CopyFolder _
                sourcePath, _
                destinationPath, _
                overwrite

        Case Else

            Err.Raise _
                vbObjectError + 1000, _
                "Cp", _
                "コピー元が存在しません: " & sourcePath

    End Select

    Cp = True
    Exit Function

ErrorHandler:

    Cp = False

End Function


'==========================================================
' Mv（move：移動）
'
' ファイル / フォルダを自動判定
'
' 引数:
'   sourcePath      - 移動元のパス（ファイル/フォルダを自動判定）
'   destinationPath - 移動先のパス
'
' 戻り値:
'   True  : 成功
'   False : 失敗
'==========================================================
' If Mv("C:\src\a.txt", "C:\dst\a.txt") Then Debug.Print "移動成功"
Public Function Mv( _
    ByVal sourcePath As String, _
    ByVal destinationPath As String _
) As Boolean

    Dim fso As Object

    On Error GoTo ErrorHandler

    Set fso = GetFso()

    Select Case PathType(sourcePath)

        Case "FILE"

            fso.MoveFile _
                sourcePath, _
                destinationPath

        Case "DIR"

            fso.MoveFolder _
                sourcePath, _
                destinationPath

        Case Else

            Err.Raise _
                vbObjectError + 1001, _
                "Mv", _
                "移動元が存在しません: " & sourcePath

    End Select

    Mv = True
    Exit Function

ErrorHandler:

    Mv = False

End Function


'==========================================================
' Rm（remove：削除）
'
' ファイル / フォルダを自動判定
'
' 引数:
'   targetPath - 削除対象のパス（ファイル/フォルダを自動判定）
'
' 戻り値:
'   True  : 成功
'   False : 失敗
'==========================================================
' If Rm("C:\src\a.txt") Then Debug.Print "削除成功"
Public Function Rm( _
    ByVal targetPath As String _
) As Boolean

    Dim fso As Object

    On Error GoTo ErrorHandler

    Set fso = GetFso()

    Select Case PathType(targetPath)

        Case "FILE"

            fso.DeleteFile _
                targetPath, _
                True

        Case "DIR"

            fso.DeleteFolder _
                targetPath, _
                True

        Case Else

            Err.Raise _
                vbObjectError + 1002, _
                "Rm", _
                "削除対象が存在しません: " & targetPath

    End Select

    Rm = True
    Exit Function

ErrorHandler:

    Rm = False

End Function

'######### CountValues
' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim d As Object: Set d = CountValues(ws, 1)
' Debug.Print d("りんご")   ' "りんご" の出現回数
'========================================
' 列内の値の出現回数を集計
'----------------------------------------
' 引数:
'   ws  - Worksheetオブジェクト
'   col - 列番号（省略時: A列）
'
' 戻り値:
'   Dictionary
'     Key  : セルの値を文字列化したもの
'     Item : 出現回数
'   ※空列のときは空の Dictionary を返す
'   ※エラー値（#N/A など）は集計対象から除外する
' 依存:
'   Application.WorksheetFunction.CountA
'========================================

Function CountValues(ws As Worksheet, Optional col As Long = 1) As Object

    Dim lastRow As Long
    Dim c As Range
    Dim d As Object
    Dim key As String

    Set d = CreateObject("Scripting.Dictionary")

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        Set CountValues = d
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For Each c In ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
        ' エラー値は Dictionary のキーにできないため除外する
        If Not IsError(c.Value) Then
            key = CStr(c.Value)
            If d.Exists(key) Then
                d(key) = d(key) + 1
            Else
                d.Add key, 1
            End If
        End If
    Next

    Set CountValues = d

End Function
'######### CreateAndDisplayTextMail
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
'######### ExportColumn
' 列の値を取得・ファイル出力する統合モジュール
' 含まれる関数: GetColumnValuesAsString / ExportColumnToFile
'========================================
' GetColumnValuesAsString - 指定列の値を区切り文字で連結した文字列を返す
' ExportColumnToFile      - 指定列の値をテキストファイルに書き出す
'========================================

' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim s As String: s = GetColumnValuesAsString(ws, 1, " / ")
' Debug.Print s    ' "A1 / A2 / A3 ..." （改行区切りのときは vbCrLf）
'========================================
' 指定列の値を文字列で返す
' ws        : 対象ワークシート
' colNum    : 取得する列番号（省略時1列目）
' delimiter : 値をつなぐ区切り（省略時 vbCrLf で改行）
' 戻り値    : 列の値をつなげた文字列
'             ※末尾に区切りは付けない
'             ※空列のときは空文字
' 依存     : Application.WorksheetFunction.CountA
'========================================
Function GetColumnValuesAsString(ws As Worksheet, _
                                 Optional colNum As Long = 1, _
                                 Optional delimiter As String = vbCrLf) As String

    Dim lastRow As Long
    Dim result As String
    Dim r As Long
    Dim value As Variant

    GetColumnValuesAsString = ""

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(colNum)) = 0 Then Exit Function

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    ' 2行目以降は区切りを前に付けることで末尾の区切りを避ける
    For r = 1 To lastRow
        value = ws.Cells(r, colNum).Value
        ' エラー値は文字列連結できないため空文字として扱う
        If IsError(value) Then value = ""
        If r > 1 Then result = result & delimiter
        result = result & value
    Next r

    GetColumnValuesAsString = result
End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' ExportColumnToFile ws, "C:\Temp\out.txt", 1   ' A列を改行区切りで出力
'========================================
' 指定列の値をテキストファイルに書き出す
' ws       : 対象ワークシート
' filePath : 保存先フルパス
' colNum   : 書き出す列番号（省略時1列目）
' delimiter: 値をつなぐ区切り（省略時改行）
' 依存     : GetColumnValuesAsString（本モジュール内）
'========================================
Sub ExportColumnToFile(ws As Worksheet, _
                       filePath As String, _
                       Optional colNum As Long = 1, _
                       Optional delimiter As String = vbCrLf)

    Dim content As String
    Dim fNum As Integer
    Dim isOpen As Boolean

    content = GetColumnValuesAsString(ws, colNum, delimiter)

    ' ファイル書き出し
    fNum = FreeFile
    isOpen = False

    On Error GoTo CloseAndExit

    Open filePath For Output As #fNum
    isOpen = True
    ' 行末の空白改行を付けないため Print の行末を ; で止める
    Print #fNum, content;
    Close #fNum
    isOpen = False
    Exit Sub

CloseAndExit:
    ' エラー時もファイルハンドルを解放する
    On Error Resume Next
    If isOpen Then Close #fNum
    On Error GoTo 0
End Sub
'######### GetFiles
' 引数:
'   folderPath - 対象フォルダパス（末尾の \ はなくても自動補完）
'   pattern    - ファイル名のワイルドカードパターン（例: *.txt）
'
' 戻り値:
'   パターンに一致したファイルのフルパスを格納した Collection
'   （対象フォルダが存在しない場合は空の Collection）
'
' 使用例:
'   Dim col As Collection: Set col = GetFiles("C:\Temp", "*.txt")
'   For Each f In col: Debug.Print f: Next
Function GetFiles(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fso As Object
    Dim fileName As String

    ' 対象フォルダが存在しない場合は空の Collection を返す
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Set GetFiles = col
        Exit Function
    End If

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    fileName = Dir(folderPath & pattern)

    Do While fileName <> ""
        col.Add folderPath & fileName
        fileName = Dir()
    Loop

    Set GetFiles = col
End Function


' 引数:
'   folderPath - 対象フォルダパス（末尾の \ はなくても自動補完）
'   pattern    - ファイル名のワイルドカードパターン（例: *.txt）
'
' 戻り値:
'   サブフォルダを含めてパターンに一致したファイルのフルパスを格納した Collection
'   （対象フォルダが存在しない場合は空の Collection）
'
' 使用例:
'   Dim col As Collection: Set col = GetFilesRecursive("C:\Temp", "*.txt")
'   For Each f In col: Debug.Print f: Next
Function GetFilesRecursive(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        GetFilesRecursive_Add fso.GetFolder(folderPath), pattern, col
    End If

    Set GetFilesRecursive = col
End Function


' 引数:
'   folder  - 走査対象の FSO Folder オブジェクト
'   pattern - ファイル名のワイルドカードパターン（例: *.txt）
'   col     - 一致したファイルのフルパスを追加する Collection（ByRef）
'
' 戻り値: なし（col に直接追加される）
Private Sub GetFilesRecursive_Add( _
    ByVal folder As Object, _
    ByVal pattern As String, _
    ByRef col As Collection)

    Dim file As Object
    Dim subFolder As Object

    ' ファイル
    For Each file In folder.Files
        ' Like は既定（Option Compare Binary）で大文字小文字を区別するため
        ' ファイル名もパターンも小文字に揃えて比較する
        If LCase(file.Name) Like LCase(pattern) Then
            col.Add file.Path
        End If
    Next file

    ' サブフォルダ
    For Each subFolder In folder.SubFolders
        GetFilesRecursive_Add subFolder, pattern, col
    Next subFolder

End Sub
'######### GetTimestamp
' Dim ts As String: ts = GetTimestamp()
' Debug.Print ts   ' 2026-08-19-09-30-00
Function GetTimestamp() As String
    ' yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す
    GetTimestamp = Format(Now, "yyyy-mm-dd-HH-MM-ss")
End Function




'######### IPaddress
' Debug.Print IsValidIPAddress("192.168.0.1")       ' True
' Debug.Print IsValidIPAddress("192.168.0.999")     ' False
' 引数:
'   ip - 検証するIPアドレス
' 戻り値:
'   有効なら True / 無効なら False
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

' Debug.Print IsValidSubnetMask("255.255.255.0")    ' True
' Debug.Print IsValidSubnetMask("255.0.255.0")      ' False（1の連続でない）
' 引数:
'   mask - 検証するサブネットマスク
' 戻り値:
'   有効なら True / 無効なら False
'   全0（0.0.0.0 = /0）と全1（255.255.255.255 = /32）も有効として扱う
'   （CIDR2Mask(0) / CIDR2Mask(32) と整合させるため）
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

        binaryStr = binaryStr & Dec2Bin8(num)
    Next i

    ' 「1が続いた後に0が続く」パターンのみ許可する
    If InStr(binaryStr, "01") > 0 Then Exit Function

    IsValidSubnetMask = True
End Function

' 0〜255 を 8桁の 2進文字列に変換する
' WorksheetFunction.Dec2Bin に依存せず自身の計算のみで求める
Private Function Dec2Bin8(ByVal num As Integer) As String
    Dim result As String
    Dim i As Integer

    For i = 7 To 0 Step -1
        If (num And 2 ^ i) > 0 Then
            result = result & "1"
        Else
            result = result & "0"
        End If
    Next i

    Dec2Bin8 = result
End Function

' Debug.Print IsValidNetworkAddress("192.168.0.0", "255.255.255.0")   ' True
' Debug.Print IsValidNetworkAddress("192.168.0.1", "255.255.255.0")   ' False
' 引数:
'   ip   - 検証するIPアドレス
'   mask - サブネットマスク
' 戻り値:
'   ネットワークアドレスなら True / そうでなければ False
Function IsValidNetworkAddress(ByVal ip As String, ByVal mask As String) As Boolean
    Dim ipParts() As String
    Dim maskParts() As String
    Dim i As Integer

    If Not IsValidIPAddress(ip) Then Exit Function
    If Not IsValidSubnetMask(mask) Then Exit Function

    ipParts = Split(ip, ".")
    maskParts = Split(mask, ".")

    For i = 0 To 3
        ' マスクの 0 の部分（ホストビット）が 1 ならネットワークアドレスではない
        ' Not を使うと符号つきの負値になるため 255 - mask でホストビットを求める
        If (CInt(ipParts(i)) And (255 - CInt(maskParts(i)))) <> 0 Then
            Exit Function
        End If
    Next i

    IsValidNetworkAddress = True
End Function

' Debug.Print CIDR2Mask(24)   ' 255.255.255.0
' Debug.Print CIDR2Mask(8)    ' 255.0.0.0
' 引数:
'   cidr - プレフィックス長（0～32）
' 戻り値:
'   ドット区切りのサブネットマスク文字列
'   範囲外（0未満・32超）の場合は空文字
Function CIDR2Mask(ByVal cidr As Integer) As String
    Dim i As Integer
    Dim mask(3) As Integer
    Dim bits As Integer

    If cidr < 0 Or cidr > 32 Then Exit Function

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

' Debug.Print Mask2CIDR("255.255.255.0")   ' 24
' Debug.Print Mask2CIDR("255.0.0.0")       ' 8
' 引数:
'   mask - ドット区切りのサブネットマスク
' 戻り値:
'   プレフィックス長（0～32）
'   形式不正または 1 の連続でない場合は -1
Function Mask2CIDR(ByVal mask As String) As Integer
    Dim parts() As String
    Dim i As Integer
    Dim val As Integer
    Dim cidr As Integer

    ' 形式不正・1の連続でないマスクは -1 を返す
    If Not IsValidSubnetMask(mask) Then
        Mask2CIDR = -1
        Exit Function
    End If

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
'######### JaroDistance
' Jaro距離を返す。0 = 完全一致、1 = 全く異なる
Function JaroDistance(ByVal s1 As String, ByVal s2 As String) As Double

    Dim len1 As Long, len2 As Long
    Dim window As Long
    Dim i As Long, j As Long
    Dim lo As Long, hi As Long
    Dim matches1() As Boolean, matches2() As Boolean
    Dim matchCount As Long
    Dim k As Long
    Dim transposeCnt As Long
    Dim jaro As Double

    len1 = Len(s1)
    len2 = Len(s2)

    ' 両方空、または完全一致なら距離0
    If len1 = 0 And len2 = 0 Then Exit Function
    If s1 = s2 Then Exit Function

    ' マッチ探索窓の半幅
    window = Max(len1, len2) \ 2 - 1
    If window < 0 Then window = 0

    ReDim matches1(1 To len1)
    ReDim matches2(1 To len2)

    ' 探索窓内で一致する文字をマーク
    For i = 1 To len1
        lo = Max(1, i - window)
        hi = Min(i + window, len2)
        For j = lo To hi
            If Not matches2(j) And Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                matches1(i) = True
                matches2(j) = True
                matchCount = matchCount + 1
                Exit For
            End If
        Next j
    Next i

    ' 共通文字なし → 距離1
    If matchCount = 0 Then
        JaroDistance = 1
        Exit Function
    End If

    ' 転置数（一致文字の並び順の違い）を数える
    k = 1
    For i = 1 To len1
        If matches1(i) Then
            Do While Not matches2(k)
                k = k + 1
            Loop
            If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then transposeCnt = transposeCnt + 1
            k = k + 1
        End If
    Next i

    ' Jaro類似度 → 距離 = 1 - 類似度
    jaro = (matchCount / len1 + matchCount / len2 + (matchCount - (transposeCnt \ 2)) / matchCount) / 3
    JaroDistance = 1 - jaro

End Function

Private Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

'######### LastUsedRow
' 使用例:
'   Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
'   Dim n As Long: n = LastUsedRow(ws)      ' A列の最終使用行
'   Debug.Print n                           ' 例: 100
' 引数:
'   ws  - 対象ワークシート
'   col - 調べる列番号（省略時: A列）
' 戻り値:
'   最終使用行番号（データが無ければ 0）
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
' Dim wb As Workbook: Set wb = OpenExcel()
' If wb Is Nothing Then Exit Sub   ' キャンセル時は Nothing が返る

Function OpenExcel() As Workbook
    Dim filename As Variant

    OpenExcel = Nothing

    filename = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xls*),*.xls*,CSVファイル (*.csv),*.csv")

    ' キャンセル時は Boolean の False が返るため VarType で判定する
    If VarType(filename) = vbBoolean Then Exit Function
    If Len(filename) = 0 Then Exit Function

    Set OpenExcel = Workbooks.Open(filename)
End Function
'######### ReadUtf8Text
' 使用例:
'   Dim s As String: s = ReadUtf8Text("C:\Temp\file.txt")
'   Debug.Print s   ' UTF-8 のテキスト内容を文字列で取得
' 引数:
'   filePath - 読み込むファイルのフルパス
' 戻り値:
'   UTF-8 のテキスト内容を文字列で返す
'   ※ファイルが存在しない／読み取り失敗の場合は空文字
' 依存:
'   ADODB.Stream（ado moniker 経由の遅延バインド。参照設定は不要）
Function ReadUtf8Text(ByVal filePath As String) As String

    Dim stm As Object

    ReadUtf8Text = ""

    ' ファイルが存在しない場合は空文字を返す
    If Len(Dir(filePath)) = 0 Then Exit Function

    On Error GoTo CleanUp

    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With

CleanUp:
    ' エラー時もストリームを確実に解放する
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
    On Error GoTo 0

End Function
'######### SaveAttachments
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
'######### SearchSheet
' ワークシート内の検索・値取得 統合モジュール
' 含まれる関数: GetColumnByHeader / GetPosition / GetValueByID
'========================================
' GetColumnByHeader - 見出し名から列番号を探す
' GetPosition       - 文字列を検索し最初に一致したセル座標を返す
' GetValueByID      - ID列の値から対象列の値を取得する
'========================================

' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim col As Long: col = GetColumnByHeader(ws, "ID")
' Debug.Print col              ' 見出し"ID"の列番号（例: 2）
'========================================
' 見出し名から列番号を探す
' ws      : 対象ワークシート
' header  : 探したい見出し名
' rowNum  : 見出しがある行番号（通常1行目）
' 戻り値  : 列番号（見つからなければ0）
'========================================
Function GetColumnByHeader(ws As Worksheet, header As String, Optional rowNum As Long = 1) As Long
    Dim lastCol As Long
    Dim c As Long

    GetColumnByHeader = 0 ' 見つからなければ0

    ' 空行のときは End(xlToLeft) が 1 列目を返すため、そのまま 1 から走査する
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If ws.Cells(rowNum, c).Value = header Then
            GetColumnByHeader = c
            Exit Function
        End If
    Next c
End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim pos As Variant: pos = GetPosition(ws, "商品コード")
' Debug.Print pos(0) & "," & pos(1)   ' 例: "5,3"（C5）
'----------------------------------------
' 指定した文字列を検索し、最初に一致したセル座標を返す
' ws        : 対象ワークシート
' searchText: 検索する文字列（完全一致）
' 戻り値    : Array(Y, X) 見つからなければ Array(0, 0)
'             Y: 行番号 / X: 列番号
' 依存      : GetColumnByHeader（本モジュール内）
'             ※空シートのときは Find が Nothing を返すため Array(0, 0)
'----------------------------------------
Function GetPosition(ws As Worksheet, searchText As String) As Variant

    Dim r As Long
    Dim lastRow As Long
    Dim col As Long
    Dim found As Range

    GetPosition = Array(0, 0)

    ' 空シートのときは Find が Nothing を返すため .Row でエラーになる
    Set found = ws.Cells.Find("*", _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious)
    If found Is Nothing Then Exit Function

    lastRow = found.Row

    For r = 1 To lastRow

        'この行の最終列まで検索
        col = GetColumnByHeader(ws, searchText, r)

        If col > 0 Then
            GetPosition = Array(r, col)
            Exit Function
        End If

    Next r

End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim v As Variant: v = GetValueByID(ws, "ID", 123, "名前")
' Debug.Print v    ' ID=123 の行の"名前"列の値（見つからなければ ""）
'========================================
' IDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
'                IDは文字列化して比較するため、数値/文字列の型差は吸収される
' 依存         : GetColumnByHeader（本モジュール内）
'                Application.WorksheetFunction.CountA
'========================================
Function GetValueByID(ws As Worksheet, _
                             idHeader As String, _
                             idValue As Variant, _
                             targetHeader As String, _
                             Optional headerRow As Long = 1) As Variant

    Dim idCol     As Long
    Dim targetCol As Long
    Dim lastRow   As Long
    Dim r         As Long
    Dim cellValue As Variant
    Dim found     As Variant

    GetValueByID = ""

    idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(idCol)) = 0 Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        cellValue = ws.Cells(r, idCol).Value
        If Not IsError(cellValue) Then ' エラー値回避
            ' Variant 同士でも型の差で比較が False になるため文字列化して比較する
            If CStr(cellValue) = CStr(idValue) Then
                found = ws.Cells(r, targetCol).Value
                If Not IsError(found) Then GetValueByID = found
                Exit Function
            End If
        End If
    Next r
End Function
'######### SelectFolder
' 使用例:
'   Dim path As String: path = SelectFolder()
'   If path = "" Then Exit Sub     ' キャンセル時は空文字
' 戻り値:
'   選択したフォルダのフルパス（キャンセル時は ""）
Function SelectFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "フォルダを選択してください"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
    Set fd = Nothing
End Function

'######### TheHash
' 使用例:
'   Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
'   Dim h As Object: Set h = TheHash(ws, 1)
'   Debug.Print h("りんご")          ' 値"りんご"が最後に現れた行番号
' 引数:
'   ws       - 対象ワークシート
'   keyIndex - キーとして使う列番号
' 戻り値:
'   Scripting.Dictionary
'     Key  : セルの値（文字列化）
'     Item : その値が最後に現れた行番号
'   ※同じ値が複数ある場合は最後の行番号で上書きされる
'   ※空列のときは空の Dictionary を返す
'   ※エラー値（#N/A など）は除外される
' 依存:
'   Application.WorksheetFunction.CountA
Function TheHash(ws As Worksheet, keyIndex As Long) As Object
    Dim myHash As Object: Set myHash = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    Dim value As Variant
    Dim lastRow As Long

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(keyIndex)) = 0 Then
        Set TheHash = myHash
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, keyIndex).End(xlUp).Row

    For i = 1 To lastRow
        value = ws.Cells(i, keyIndex).Value
        ' エラー値は CStr できないため除外する
        If Not IsError(value) Then
            key = CStr(value)
            myHash(key) = i
        End If
    Next i

    Set TheHash = myHash
End Function
