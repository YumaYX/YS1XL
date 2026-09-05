' マウス操作用の Win32 API 宣言
'
' 使用例:
'   SetCursorPos 100, 100
'   mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'   mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'
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

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const MOUSEEVENTF_LEFTDOWN  As Long = &H2
Private Const MOUSEEVENTF_LEFTUP    As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

