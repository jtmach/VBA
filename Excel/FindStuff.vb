Public Function FindStuff(criteria As String, criteriaDate As Date, _
  lookupSheet As String, criteriaColumn As Integer, startDateColumn As Integer, _
  endDateColumn As Integer, resultsColumn As Integer)
  
  Dim foundStuff As Boolean: foundStuff = False
  Dim wrkSheet As Worksheet
  For Each wrkSheet In Worksheets
    If LCase(wrkSheet.Name) = LCase(lookupSheet) Then
      Dim rowIndex As Long
      For rowIndex = 1 To wrkSheet.UsedRange.Rows.Count
        If wrkSheet.Cells(rowIndex, criteriaColumn) = criteria Then
          If IsDate(wrkSheet.Cells(rowIndex, startDateColumn)) And IsDate(wrkSheet.Cells(rowIndex, endDateColumn)) Then
            If wrkSheet.Cells(rowIndex, startDateColumn) <= criteriaDate And wrkSheet.Cells(rowIndex, endDateColumn) >= criteriaDate Then
              FindStuff = wrkSheet.Cells(rowIndex, resultsColumn)
              foundStuff = True
              Exit For
            End If
          End If
        End If
      Next
    End If
  Next
  
  If Not foundStuff Then
    FindStuff = "N\A"
  End If
End Function