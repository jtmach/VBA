Public Function Unpivot()
  Dim wrkSheet As Worksheet
  Set wrkSheet = ActiveSheet

  Dim newWrkSheet As Worksheet
  Set newWrkSheet = Sheets.Add()
  
  Dim newRowIndex As Long
  newRowIndex = 1
  
  Dim dataColumn As Integer
  For dataColumn = 7 To wrkSheet.UsedRange.Columns.Count
    Dim rowIndex As Long
    For rowIndex = 2 To wrkSheet.UsedRange.Rows.Count
      Dim startDate As Date
      startDate = wrkSheet.Cells(1, dataColumn)
      
      newWrkSheet.Cells(newRowIndex, 1) = wrkSheet.Cells(rowIndex, 4)
      newWrkSheet.Cells(newRowIndex, 2) = startDate
      newWrkSheet.Cells(newRowIndex, 3) = startDate + 6
      newWrkSheet.Cells(newRowIndex, 4) = wrkSheet.Cells(rowIndex, dataColumn)
      newWrkSheet.Cells(newRowIndex, 5) = wrkSheet.Cells(rowIndex, 6)
      
      newRowIndex = newRowIndex + 1
    Next
  Next
End Function