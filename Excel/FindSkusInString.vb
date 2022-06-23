Option Explicit On

' Looks for a 7 or 8 digit number within cell text, if found return the first match.
Function FindSkus(searchCell)
  Dim RegEx As Object
  RegEx = CreateObject("VBScript.RegExp")
  RegEx.Global = True
  RegEx.Pattern = "(\d){7,8}"

  Dim foundSkus
  If RegEx.Test(searchCell) Then
    Dim matches As Object
    matches = RegEx.Execute(searchCell)
    Dim match As Object
    For Each match In matches
      foundSkus = foundSkus & match & ","
    Next
    foundSkus = Left(foundSkus, Len(foundSkus) - 1)
  End If

  FindSkus = foundSkus
End Function