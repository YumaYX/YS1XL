'######### Grep
' 指定列を走査し、キーワードに一致するセルの値をイミディエイトに出力する
'
' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Grep ws, "りんご"         ' A列から部分一致で検索して出力
' Grep ws, "A001", 2        ' B列から検索（省略時はA列）
' Grep ws, "山田", 1, 2     ' A列を完全一致で検索
'========================================
' キーワードでセルの値を検索し、一致した値を出力（grep風）
' ws       : 対象ワークシート
' keyword  : 検索キーワード
' col      : 対象列番号（省略時: 1 = A列）
' matchMode: 一致モード（省略時: 1）
'            1=部分一致 / 2=完全一致 / 3=前方一致 / 4=後方一致
'            （すべて大文字小文字を区別しない）
' 戻り値   : なし（一致した値を Debug.Print で出力）
'========================================
Sub Grep(ws As Worksheet, _
         keyword As String, _
         Optional col As Long = 1, _
         Optional matchMode As Long = 1)

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim r As Long
    Dim v As Variant
    Dim s As String
    Dim k As String: k = LCase$(keyword)

    For r = 1 To lastRow
        v = ws.Cells(r, col).Value
        If Not IsError(v) Then ' エラー値回避
            s = LCase$(CStr(v))
            Select Case matchMode
                Case 2
                    If s = k Then Debug.Print CStr(v)
                Case 3
                    If Left$(s, Len(k)) = k Then Debug.Print CStr(v)
                Case 4
                    If Right$(s, Len(k)) = k Then Debug.Print CStr(v)
                Case Else ' 1=部分一致
                    If InStr(s, k) > 0 Then Debug.Print CStr(v)
            End Select
        End If
    Next r

End Sub
