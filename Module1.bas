Attribute VB_Name = "Module1"
Public Const A_CENTER = &H300&
Public Const A_TOP = &H400&
Public Const A_TOPLEFT = &H500&
Public Const A_TOPRIGHT = &H600&
Public Const A_BOTTOM = &H800&
Public Const A_BOTTOMLEFT = &H900&
Public Const A_BOTTOMRIGHT = &HA00&
Public Const A_LEFT = &H100&
Public Const A_RIGHT = &H200&
Public Const GWL_STYLE& = (-16)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long)
Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)


