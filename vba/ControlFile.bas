' If Cp("C:\src\a.txt", "C:\dst\a.txt") Then Debug.Print "コピー成功"
' If Cp("C:\src\folder", "C:\dst\folder") Then Debug.Print "フォルダも自動判定"
' Option Explicit

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


