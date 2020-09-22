<div align="center">

## MsFlexGrid act like ListBox


</div>

### Description

When you click on a grid cell in msflex grid it will highlight the whole row and unhighlight the other rows.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brandon Lepley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brandon-lepley.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brandon-lepley-msflexgrid-act-like-listbox__1-29747/archive/master.zip)





### Source Code

```
'Author: Brandon Lepley
'Purpose: When the MsFlexGrid is clicked it will act like the list box control
Private Sub MsFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call SelectGridCell(MSFlexGrid1.Row)
End Sub
Private Sub SelectGridCell(GridRow As Integer)
  Dim i As Integer
  Dim j As Integer
  Dim NumRows As Integer
  Dim NumCols As Integer
  NumRows = MSFlexGrid1.Rows - 1 '.rows returns num of rows
  NumCols = MSFlexGrid1.Cols - 1 '.cols reutrns num of columns
  MSFlexGrid1.HighLight = flexHighlightNever 'since this sub takes
  'care of highlighting we tell it to never highlight so only 1 row
  'is selected at a time
  For i = 1 To NumRows
    If i <> GridRow Then
      MSFlexGrid1.Row = i
      For j = 1 To NumCols
        MSFlexGrid1.Col = j
        If MSFlexGrid1.CellBackColor = vbHighlight Then
          MSFlexGrid1.CellBackColor = vbWindowBackground
          MSFlexGrid1.CellForeColor = vbWindowText
        Else
          Exit For
        End If
      Next j
    End If
  Next i
  MSFlexGrid1.Row = GridRow 'set the row to the clicked row
  For i = 1 To NumCols 'setting the clicked row to highlighted
    MSFlexGrid1.Col = i
    MSFlexGrid1.CellBackColor = vbHighlight
    MSFlexGrid1.CellForeColor = vbHighlightText
  Next i
  'note when leaving this sub the msflexgrid.row will be the gridrow
  'and the column will be the last column. i.e. if there are 4 columns
  'then the msflexgrid1.col will be 4.
End Sub
```

