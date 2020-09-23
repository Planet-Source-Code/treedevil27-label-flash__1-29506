<div align="center">

## Label Flash


</div>

### Description

Flash a label and its caption between starting forecolour and colour of your choice.
 
### More Info
 
Label by reference, number of flash cycles, flash colour.

The Timer object


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[treedevil27](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/treedevil27.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/treedevil27-label-flash__1-29506/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  LabelFlash Me.Label1, 5, vbBlack
End Sub
Private Sub Form_Load()
  Me.Label1.ForeColor = vbWhite
End Sub
Public Function LabelFlash(ByRef lblLabel As Label, _
              ByVal lngCycles As Integer, _
              ByVal lngOffColour As Long) As Integer
  Dim lngOnColour   As Long
  Dim lngStart    As Long
  Dim lngTick     As Long
  Dim lngX      As Long
  ' Get the starting colour
  lngOnColour = lblLabel.ForeColor
  ' Get the starting time rounded to seconds
  lngStart = Round(Timer, 0)
  DoEvents
  While Not Round(Timer, 0) > (lngStart + lngCycles)
    If Round(Timer) > lngTick Then 'only kick over if a second has passed
      DoEvents
      ' Swap the on and off colours
      lblLabel.ForeColor = IIf(lblLabel.ForeColor = lngOffColour, lngOnColour, lngOffColour)
      lngTick = Round(Timer, 0)
    End If
  Wend
  ' Go Back to original colours
  lblLabel.ForeColor = lngOnColour
End Function
```

