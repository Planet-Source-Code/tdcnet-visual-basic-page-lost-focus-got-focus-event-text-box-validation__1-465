<div align="center">

## Lost Focus / Got Focus Event\-\-text box validation


</div>

### Description

These events are usually ignored or inconsistent amongst programs. For the users benefit, highlighting the current textbox, or tab control will aid in their navigation of your forms. But how to keep all these events consistent? Here is the answer. (Well our answer anyhow... until full-inheritance in VB 5.0)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TDCNET Visual Basic Page](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tdcnet-visual-basic-page.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tdcnet-visual-basic-page-lost-focus-got-focus-event-text-box-validation__1-465/archive/master.zip)





### Source Code

```
Add these two procedures to a Module. In each object GotFocus.LostFocus event, place a call to the respective procedure (the CALL qualifier is not neccesary, just the procedure name). This process can also be placed in a VB 4.0 Class.
Public Sub GotFocus()
Set gLastObjectFocus = Screen.ActiveControl
With gLastObjectFocus
If (TypeOf gLastObjectFocus Is TextBox) Or _
(TypeOf gLastObjectFocus Is ComboBox) Or _
(TypeOf gLastObjectFocus Is CSComboBox) Or _
(TypeOf gLastObjectFocus Is sidtEdit) _
Then
.BackColor = &HFF0000 'Dark Blue
ElseIf (TypeOf gLastObjectFocus Is SSTab) Then
.Font.Bold = True
.Font.Italic = True
.ShowFocusRect = True
ElseIf (TypeOf gLastObjectFocus Is CheckBox) Or _
(TypeOf gLastObjectFocus Is CSOptList) Or _
(TypeOf gLastObjectFocus Is OptionButton) Or _
(TypeOf gLastObjectFocus Is SSOption) Then
.ForeColor = &HFF0000 'Dark Blue
End If
End With
End Sub
Public Sub LostFocus()
With gLastObjectFocus
If (TypeOf gLastObjectFocus Is TextBox) Or _
(TypeOf gLastObjectFocus Is ComboBox) Or _
(TypeOf gLastObjectFocus Is CSComboBox) Or _
(TypeOf gLastObjectFocus Is sidtEdit) _
Then
.BackColor = &H00C0C0C0& 'Light Grey
ElseIf (TypeOf gLastObjectFocus Is SSTab) Then
.Font.Bold = False
.Font.Italic = False
.ShowFocusRect = False
ElseIf (TypeOf gLastObjectFocus Is CheckBox) Or _
(TypeOf gLastObjectFocus Is CSOptList) Or _
(TypeOf gLastObjectFocus Is OptionButton) Or _
(TypeOf gLastObjectFocus Is SSOption) Then
.ForeColor = &H0& 'Black
End If
End With
End Sub
```

