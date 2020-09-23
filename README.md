<div align="center">

## Change the system menu 4 ur form\!


</div>

### Description

Changes your forms system menu (visible when you right click on the forms titlebar or press [Alt] + [Space]) to what every you want, and does it really easily too!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MangaMorgan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mangamorgan.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mangamorgan-change-the-system-menu-4-ur-form__1-5866/archive/master.zip)





### Source Code

```
'~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~
' place a button on the form called "command1" and test
' run this project. Notice how BEFORE you click the button
' the forms system menu (press [Alt] + [Space]) is the
' normal on. Now press the button! It has changed! :)
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
' ^ APIs required 4 menu change!
Const MF_STRING = &H0&
' ^ CONSTANTs required 4 menu change!
Private Sub command1_click()
 Dim hMenu As Long, MenuItem As Long
 hMenu = GetSystemMenu(Me.hwnd, 0)
 MenuItem = GetMenuItemID(hMenu, 0)
 ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, "Restore my Bollocks"
 MenuItem = GetMenuItemID(hMenu, 1)
 ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, "Move u'r fat arse!"
 MenuItem = GetMenuItemID(hMenu, 6)
 ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, "Bugger off!"
End Sub
```

