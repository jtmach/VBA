Public Function MarkChanges()
  Me.Columns.AutoFit

  ' Change colToCheck to the column number that will have changes in it that you want to underline based off of
  Dim colToCheck As Integer: colToCheck = 2
  Dim lastValue As String

  Dim maxRows As Long
  maxRows = Me.UsedRange.Rows.Count
  
  Dim currRow As Long
  ' Start at row 2 becasue we presume that the top row contains headers, change if necessary.
  For currRow = 2 To maxRows Step 1
    If lastValue <> Me.Cells(currRow, colToCheck) Then
      Dim rng As Range
      Set rng = Me.Rows(currRow)
      
      With rng.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
    End If
    lastValue = Me.Cells(currRow, colToCheck)
  Next
End Function