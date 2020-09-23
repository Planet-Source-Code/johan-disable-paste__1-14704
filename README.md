<div align="center">

## Disable Paste


</div>

### Description

Disable paste in a textbox.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Johan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/johan.md)
**Level**          |Advanced
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/johan-disable-paste__1-14704/archive/master.zip)

### API Declarations

```
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_PASTE = &H302
Type POINTAPI
 x As Long
 y As Long
End Type
Type Msg
 hwnd As Long
 message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type
```


### Source Code

```
Put the following code in a bas module.
MODULE:
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_PASTE = &H302
Type POINTAPI
 x As Long
 y As Long
End Type
Type Msg
 hwnd As Long
 message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type
Dim mlPrevProc As Long
Public Sub Hook(robjTextbox As TextBox)
 mlPrevProc = SetWindowLong(robjTextbox.hwnd, GWL_WNDPROC, AddressOf TextProc)
End Sub
Public Sub UnHook(robjTextbox As TextBox)
 SetWindowLong robjTextbox.hwnd, GWL_WNDPROC, PrevProc
End Sub
Public Function TextProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 If uMsg = WM_PASTE Then
  uMsg = 0
 End If
 TextProc = CallWindowProc(mlPrevProc, hwnd, uMsg, wParam, lParam)
End Function
Put the following code in a form.
Option Explicit
Private Sub Form_Load()
 Hook Text1
End Sub
Private Sub Form_Unload(Cancel As Integer)
 UnHook Text1
End Sub
```

