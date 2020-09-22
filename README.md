<div align="center">

## Titlebar Animation Module


</div>

### Description

Have you ever wanted to make use of the animations Windows uses when you minimize and maximize open windows? Now you can, and it's eaiser than you think! Use this simple module to make all of your forms open and close with animations. Works on any Win32 system. The animation is drawn with the caption of the opening window in it, and uses your system colors (and gradients, if your system supports them) to create the titlebar animations. Enjoy! I don't care about the votes.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Doctor Evil](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/doctor-evil.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/doctor-evil-titlebar-animation-module__1-22883/archive/master.zip)





### Source Code

```
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Const IDANI_CAPTION = &H3
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
'ShowWindow
'Opens your from with animation from an object to the window to show.
' From_Object_hWnd: the hWnd of the object to start the animation from. This is usually the button that is clicked on to open a form.
'ToWindow: The form to open.
'ShowModal: Show the the from as a modal form? (similar to the [Modal] parameter of Form.Show)
'OwnerOfNewWindow: The owner of a form. (similar to the [OwnerForm] parameter of Form.Show)
'CenterWindow: Center the window on the screen? This is important, as if you only set the StartUpPosition property of a form to CenterScreen, the animation will run before the form is centered and will look funny. The form will be centered over the owner.
Public Sub ShowWindow(From_Object_hWnd As Long, ToWindow As Form, Optional ShowModal As Integer = vbModeless, Optional OwnerOfNewWindow As Form, Optional CenterWindow As Boolean)
If ShowModal <> 0 And ShowModal <> 1 Then
Err.Raise 15448, "ShowWindowAnimation", "Animated Window Show: ShowModal must be a value of 0 or 1. Requested value was " & ShowModal & ". Window will not be opened."
Exit Sub
End If
On Error Resume Next
Load ToWindow
If CenterWindow Then
CenterChild OwnerOfNewWindow, ToWindow
End If
  Dim FromRect As RECT, ToRect As RECT
  GetWindowRect From_Object_hWnd, FromRect
  GetWindowRect ToWindow.hwnd, ToRect
  DrawAnimatedRects ToWindow.hwnd, IDANI_CAPTION, FromRect, ToRect
ToWindow.Show ShowModal, OwnerOfNewWindow
End Sub
'UnloadWindow
'Use this to make an animation from a window to an object when a window is closing. You could put this in the Form_Unload event:
' UnloadWindow Me, PreviousWindow.Command1.hWnd
Public Sub UnloadWindow(WindowToClose As Form, Close_To_Object_hWnd As Long)
On Error Resume Next
  Dim FromRect As RECT, ToRect As RECT
  GetWindowRect WindowToClose.hwnd, FromRect
  GetWindowRect Close_To_Object_hWnd, ToRect
  DrawAnimatedRects WindowToClose.hwnd, IDANI_CAPTION, FromRect, ToRect
Unload WindowToClose
End Sub
'Centers a child window over a parent window.
Public Sub CenterChild(Parent As Form, Child As Form)
  On Local Error Resume Next
  If Parent.WindowState = 1 Then
    Exit Sub
  Else
    Child.Left = (Parent.Left + (Parent.Width / 2)) - (Child.Width / 2)
    Child.Top = (Parent.Top + (Parent.Height / 2)) - (Child.Height / 2)
  End If
End Sub
'ShowWindowFromMouse
'Somewhat like ShowWindow, but instead of starting the animation from an object, it starts the animation from the position of the mouse on the screen. This is useful for menus.
Public Sub ShowWindowFromMouse(ToWindow As Form, Optional ShowModal As Integer = vbModeless, Optional OwnerOfNewWindow As Form)
If ShowModal <> 0 And ShowModal <> 1 Then
Err.Raise 15448, "ShowWindowAnimation", "Animated Window Show: ShowModal must be a value of 0 or 1. Requested value was " & ShowModal & ". Window will not be opened."
Exit Sub
End If
On Error Resume Next
Load ToWindow
  Dim FromRect As RECT, ToRect As RECT, Mouse As POINTAPI
  GetCursorPos Mouse
  FromRect.Top = Mouse.Y
  FromRect.Left = Mouse.X
  FromRect.Bottom = Mouse.Y + 32
  FromRect.Right = Mouse.X + 32
  GetWindowRect ToWindow.hwnd, ToRect
  DrawAnimatedRects ToWindow.hwnd, IDANI_CAPTION, FromRect, ToRect
ToWindow.Show ShowModal, OwnerOfNewWindow
End Sub
'Makes an animation from the hWnd of an object to the position of the mouse.
Public Sub MouseTohWnd(AnimateTo As Long)
On Error Resume Next
  Dim FromRect As RECT, ToRect As RECT, Mouse As POINTAPI
  GetCursorPos Mouse
  FromRect.Top = Mouse.Y
  FromRect.Left = Mouse.X
  FromRect.Bottom = Mouse.Y + 32
  FromRect.Right = Mouse.X + 32
  GetWindowRect AnimateTo, ToRect
  DrawAnimatedRects AnimateTo, IDANI_CAPTION, FromRect, ToRect
End Sub
```

