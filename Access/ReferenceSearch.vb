Option Compare Database
Option Explicit

Public Function SearchAll(searchText As String) As Boolean
  SearchAll = SearchQueries(searchText)
  If Not SearchAll Then SearchAll = SearchForms(searchText)
  If Not SearchAll Then SearchAll = SearchModules(searchText)
End Function

Public Function SearchQueries(searchText As String) As Boolean
On Error GoTo err
  Dim qry As queryDef
  For Each qry In CurrentDb.QueryDefs
    If InStr(qry.sql, searchText) > 0 Then
      Debug.Print "Search text found in query: " & qry.Name
      SearchQueries = True
    End If
  Next
  
Exit Function
err:
  If err.Number = 3258 Then
    Debug.Print "Error evaluating query (" & qry.Name & "): " & err.Description
    Resume
  End If
  
  err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function

Public Function SearchForms(searchText As String) As Boolean
  Dim idx As Integer
  Dim frm As Form
  For idx = 0 To CurrentProject.AllForms.Count - 1
    If Not CurrentProject.AllForms(idx).IsLoaded Then
      DoCmd.OpenForm CurrentProject.AllForms(idx).Name, acDesign
    End If
    
    Set frm = Forms(0)
    If InStr(frm.RecordSource, searchText) > 0 Then
      Debug.Print "Search text found in form: " & frm.Name
      SearchForms = True
    End If
    
    Dim ctrl As Control
    For Each ctrl In frm.Controls
      If (SearchControl(ctrl, searchText)) Then
        Debug.Print "Search text found in control: " & ctrl.Name & " on form: " & frm.Name
        SearchForms = True
      End If
    Next
    
    DoCmd.Close acForm, CurrentProject.AllForms(idx).Name, acSaveNo
  Next idx
End Function

Public Function SearchModules(searchText As String) As Boolean
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
      Debug.Print "Search text found in module: " & mo.Name
      SearchModules = True
    End If
    
    DoCmd.Close acModule, mo.Name, acSavePrompt
  Next
End Function

Public Function SearchReports(searchText As String) As Boolean
  Dim idx As Integer
  Dim rpt As Report
  
  For Each rpt In Reports
    DoCmd.Close acReport, rpt.Name, acSaveNo
  Next
  
  For idx = 0 To CurrentProject.AllReports.Count - 1
    If Not CurrentProject.AllReports(idx).IsLoaded Then
      DoCmd.OpenReport CurrentProject.AllReports(idx).Name, acViewDesign, , , acHidden
    End If
    
    Set rpt = Reports(0)
    If InStr(rpt.RecordSource, searchText) > 0 Then
      Debug.Print "Search text found in report: " & rpt.Name
      SearchReports = True
    End If
    
    DoCmd.Close acReport, rpt.Name, acSavePrompt
  Next
End Function

Private Function SearchControl(ctrl As Control, searchText As String) As Boolean
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
    Case "ListBox":
      Dim lstBox As listBox
      Set lstBox = ctrl
      If InStr(lstBox.ControlSource, searchText) > 0 Then
        SearchControl = True
      End If
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
      Debug.Print "Unknown Type: (" & TypeName(ctrl) & ")";
  End Select
End Function
