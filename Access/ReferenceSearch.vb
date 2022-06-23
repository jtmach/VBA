Option Compare Database
Option Explicit

Public Function Search()
  ClearTempQueries

  Dim qry As queryDef
  For Each qry In CurrentDb.QueryDefs
    If Not Left(qry.Name, 1) = "~" And Not Left(qry.Name, 2) = "x_" Then
      If Not SearchAll(qry.Name, True) Then
        Debug.Print "--Query (" & qry.Name & ") does not appear to be used."
        qry.Name = "x_" & qry.Name
      End If
    End If
  Next
  
  Dim tbl As TableDef
  For Each tbl In CurrentDb.TableDefs
    If Not UCase(Left(tbl.Name, 4)) = "MSYS" And Not UCase(Left(tbl.Name, 2)) = "x_" Then
      If Not SearchAll(tbl.Name, True) Then
        Debug.Print "--Table (" & tbl.Name & ") does not appear to be used."
        tbl.Name = "x_" & tbl.Name
      End If
    End If
  Next
End Function


Public Function SearchAll(searchText As String, Optional silent As Boolean = False) As Boolean
  Debug.Print "Searching all for: " & searchText
  
  Dim foundQueries As Boolean: foundQueries = SearchQueries(searchText, silent)
  Dim foundForms As Boolean: foundForms = SearchForms(searchText, silent)
  Dim foundModules As Boolean: foundModules = SearchModules(searchText, silent)
  Dim foundReports As Boolean: foundReports = SearchReports(searchText, silent)
  
  SearchAll = foundQueries Or foundForms Or foundModules Or foundReports
End Function

Public Function SearchQueries(searchText As String, Optional silent As Boolean = False) As Boolean
On Error GoTo Err
  Dim qry As queryDef
  For Each qry In CurrentDb.QueryDefs
    If Not Left(qry.Name, 2) = "x_" Then
      If InStr(qry.sql, searchText) > 0 Then
        If Not silent Then Debug.Print "Search text found in query: " & qry.Name
        SearchQueries = True
      End If
    End If
  Next
  
Exit Function
Err:
  If Err.Number = 3258 Then
    If Not silent Then Debug.Print "Error evaluating query (" & qry.Name & "): " & Err.Description
    Resume
  End If
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function SearchForms(searchText As String, Optional silent As Boolean = False) As Boolean
  Dim idx As Integer
  Dim frm As Form
  For idx = 0 To CurrentProject.AllForms.Count - 1
    If Not (CurrentProject.AllForms(idx).Name = "frmClean") Then
      If Not CurrentProject.AllForms(idx).IsLoaded Then
        DoCmd.OpenForm CurrentProject.AllForms(idx).Name, acDesign
      End If
      
      Set frm = Forms(CurrentProject.AllForms(idx).Name)
      If InStr(frm.RecordSource, searchText) > 0 Then
        If Not silent Then Debug.Print "Search text found in form: " & frm.Name
        SearchForms = True
      End If
      
      If (frm.HasModule) Then
        If frm.Module.Find(searchText, 0, 0, frm.Module.CountOfLines, 1000, , , True) Then
          If Not silent Then Debug.Print "Search text found in module: " & frm.Module.Name
          SearchForms = True
        End If
      End If
      
      Dim ctrl As Control
      For Each ctrl In frm.Controls
        If (SearchControl(ctrl, searchText)) Then
          If Not silent Then Debug.Print "Search text found in control: " & ctrl.Name & " on form: " & frm.Name
          SearchForms = True
        End If
      Next
      
      DoCmd.Close acForm, CurrentProject.AllForms(idx).Name, acSaveNo
    End If
  Next idx
End Function

Public Function SearchModules(searchText As String, Optional silent As Boolean = False) As Boolean
  Dim idx As Integer
  Dim mo As Module
  
  For Each mo In Modules
    DoCmd.Close acModule, mo.Name, acSaveNo
  Next
  
  For idx = 0 To CurrentProject.AllModules.Count - 1
    If Not CurrentProject.AllModules(idx).IsLoaded Then
      DoCmd.OpenModule CurrentProject.AllModules(idx).Name
    End If
    
    Set mo = Modules(0)
    If mo.Find(searchText, 0, 0, mo.CountOfLines, 1000, , , True) Then
      If Not silent Then Debug.Print "Search text found in module: " & mo.Name
      SearchModules = True
    End If
    
    DoCmd.Close acModule, mo.Name, acSavePrompt
  Next
End Function

Public Function SearchReports(searchText As String, Optional silent As Boolean = False) As Boolean
  Dim idx As Integer
  Dim rpt As Report

  For Each rpt In Reports
    DoCmd.Close acReport, rpt.Name, acSaveNo
    Debug.Print rpt.Name
  Next
  
  For idx = 0 To CurrentProject.AllReports.Count - 1
    If Not CurrentProject.AllReports(idx).IsLoaded Then
      DoCmd.OpenReport CurrentProject.AllReports(idx).Name, acViewDesign, , , acHidden
    End If
    
    Set rpt = Reports(0)
    If InStr(rpt.RecordSource, searchText) > 0 Then
      If Not silent Then Debug.Print "Search text found in report: " & rpt.Name
      SearchReports = True
    End If
    
    If (rpt.HasModule) Then
      If rpt.Module.Find(searchText, 0, 0, rpt.Module.CountOfLines, 1000, , , True) Then
        If Not silent Then Debug.Print "Search text found in module: " & rpt.Module.Name
        SearchReports = True
      End If
    End If
    
    Dim ctrl As Control
    For Each ctrl In rpt.Controls
      If (SearchControl(ctrl, searchText)) Then
        If Not silent Then Debug.Print "Search text found in control: " & ctrl.Name & " on report: " & rpt.Name
        SearchReports = True
      End If
    Next
    
    DoCmd.Close acReport, rpt.Name, acSavePrompt
  Next
End Function

Private Function ClearTempQueries(db As Database)
  Dim qryCount As Integer: qryCount = db.QueryDefs.Count
  Dim qryIndex As Integer
  
  For qryIndex = 1 To qryCount
    If Left(db.QueryDefs(qryIndex).Name, 4) = "~sq_" Then
      Debug.Print "Deleting: " & db.QueryDefs(qryIndex).Name
      db.QueryDefs.Delete (db.QueryDefs(qryIndex).Name)
      qryIndex = qryIndex - 1
      qryCount = qryCount - 1
    End If
  Next
End Function

Private Function SearchControl(ctrl As Control, searchText As String, Optional silent As Boolean = False) As Boolean
  Select Case TypeName(ctrl)
    Case "CheckBox":
      DoEvents
    Case "CommandButton":
      DoEvents
    Case "ComboBox":
      Dim cmbBox As ComboBox
      Set cmbBox = ctrl
      If InStr(cmbBox.ControlSource, searchText) > 0 Then
        SearchControl = True
      End If
    Case "Image":
      DoEvents
    Case "Label":
      DoEvents
    Case "Line":
      DoEvents
    Case "ListBox":
      Dim lstBox As listBox
      Set lstBox = ctrl
      If InStr(lstBox.ControlSource, searchText) > 0 Then
        SearchControl = True
      End If
    Case "ObjectFrame":
      DoEvents
    Case "OptionButton":
      DoEvents
    Case "OptionGroup":
      Dim optGrp As OptionGroup
      Set optGrp = ctrl
      If InStr(optGrp.ControlSource, searchText) > 0 Then
        SearchControl = True
      End If
    Case "Page":
      DoEvents
    Case "PageBreak":
      DoEvents
    Case "Rectangle":
      DoEvents
    Case "SubForm":
      DoEvents
    Case "TabControl":
      DoEvents
    Case "TextBox":
      Dim txtBox As TextBox
      Set txtBox = ctrl
      If InStr(txtBox.ControlSource, searchText) > 0 Then
        SearchControl = True
      End If
    Case "ToggleButton":
      DoEvents
    Case Else:
      If Not silent Then Debug.Print "Unknown Type: (" & TypeName(ctrl) & ")";
  End Select
End Function