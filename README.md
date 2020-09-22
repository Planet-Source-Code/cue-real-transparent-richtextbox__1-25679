<div align="center">

## REAL Transparent RichTextBox


</div>

### Description

Create a real transparent RichTextBox with the standard Microsoft RichtextBox Control !!
 
### More Info
 
hWnd of the RichTextBox

Put the : result = SetWindowLong(txtLogFile.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT) Function into a Event like FormLoad e.g. !!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\-cue\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cue.md)
**Level**          |Advanced
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cue-real-transparent-richtextbox__1-25679/archive/master.zip)

### API Declarations

```
Const GWL_EXSTYLE = (-20)
 Const WS_EX_TRANSPARENT = &H20&
 Const WS_EX_LAYERED = &H80000
 Const LWA_ALPHA = &H2&
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
```


### Source Code

```
Dim result As Long
'//set Richtext Box Backgroundstyle to transparent
result = SetWindowLong(txtLogFile.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
```

