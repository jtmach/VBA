Option Explicit

Dim DymaxCodes As New Collection

Public Function GetDymaxCodes()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "Results" Then
      Dim rowIndex As Long
      For rowIndex = 1 To ws.UsedRange.Rows.Count
        If (Trim(ws.Cells(rowIndex, "F").Value) <> "") Then
          AddCode Trim(ws.Cells(rowIndex, "F").Value)
        End If
      Next rowIndex
    End If
  Next
    
  Dim codeIndex As Long
  For codeIndex = 1 To DymaxCodes.Count
    ActiveWorkbook.Worksheets(1).Cells(codeIndex, 1).Value = DymaxCodes.Item(codeIndex)
  Next codeIndex
End Function

Private Function AddCode(dymaxCode As String)
  Dim code As Variant
  For Each code In DymaxCodes
    If code = dymaxCode Then
      Exit Function
    End If
  Next
  
  DymaxCodes.Add (dymaxCode)
End Function

