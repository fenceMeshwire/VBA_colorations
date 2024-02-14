Option Explicit

' ________________________________________________________________
Sub cross_yellow_cells_diagonal_up()

Dim lngCol As Long, lngColMax As Long
Dim lngRow As Long, lngRowMax As Long

Dim lngColorNumber As Long

Dim wkSheet As Worksheet

Set wkSheet = Sheet1

lngColorNumber = 65535

With wkSheet
  lngColMax = .UsedRange.Columns.Count
  lngRowMax = .UsedRange.Rows.Count
  
  For lngCol = 1 To lngColMax
    For lngRow = 2 To lngRowMax
      If .Cells(lngRow, lngCol).Interior.Color = lngColorNumber Then
        Call cross_yellow_cell(wkSheet, lngRow, lngCol)
      End If
    Next lngRow
  Next lngCol
  
End With

End Sub

' ________________________________________________________________
Function cross_yellow_cell_diagonal_up(ByRef wkSheet As Worksheet, ByVal lngRow As Long, ByVal lngCol As Long)

With wkSheet.Cells(lngRow, lngCol).Borders(xlDiagonalUp)
  .LineStyle = xlContinuous
  .ColorIndex = 0
  .TintAndShade = 0
  .Weight = xlThin
End With

End Function
