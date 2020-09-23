<div align="center">

## Make MSHFlexGrid Editable without help of any textbox or other control


</div>

### Description

Make MSHFlexGrid Editable without help of any textbox or other control
 
### More Info
 
'Take an MSHFlexgrid name it as msh1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kazi Khalid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kazi-khalid.md)
**Level**          |Intermediate
**User Rating**    |4.9 (113 globes from 23 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kazi-khalid-make-mshflexgrid-editable-without-help-of-any-textbox-or-other-control__1-50802/archive/master.zip)





### Source Code

```
'in the keypress event of msh1 write the following code.
Private Sub msh1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyReturn, vbKeyTab
      'move to next cell.
      With msh1
        If .Col + 1 <= .Cols - 1 Then
          .Col = .Col + 1
        Else
          If .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .Col = 0
          Else
            .Row = 1
            .Col = 0
          End If
        End If
      End With
    Case vbKeyBack
      With msh1
        'remove the last character, if any.
        If Len(.Text) Then
          .Text = Left(.Text, Len(.Text) - 1)
        End If
      End With
    Case Is < 32
    Case Else
      With msh1
        .Text = .Text & Chr(KeyAscii)
      End With
  End Select
End Sub
```

