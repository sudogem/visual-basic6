Attribute VB_Name = "Global1"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_SYSMENU = &H80000

Public Enum AddEditState
  AddState = 1
  EditState = 2
End Enum

Public Sub SetCaptionLess(hwnd As Long)
  Dim style As Long
  style = GetWindowLong(hwnd, GWL_STYLE)
  style = style Xor WS_CAPTION
  'style = style Xor WS_SYSMENU
  Call SetWindowLong(hwnd, GWL_STYLE, style)
End Sub
