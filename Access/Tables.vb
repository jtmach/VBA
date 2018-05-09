option explicit

Public Function ShowTableRows()
  Dim tblDef As TableDef
  For Each tblDef In CurrentDb.TableDefs
    If tblDef.Connect = "" Then
      Debug.Print tblDef.Name & " - " & tblDef.RecordCount
    End If
  Next
End Function