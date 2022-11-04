Attribute VB_Name = "UpdateConnectionStrings"
Option Compare Database
Option Explicit

' Use this function to update a specific database file
Public Function UpdateDb()
  Dim dbPath As String
  dbPath = "C:\Temp\Test.mdb"

  Dim db As DAO.Database
  Set db = OpenDatabase(dbPath)
    
  UpdateConnectionStrings db, "SearchFor", "ReplaceWith"
  
  Debug.Print "Finished"
End Function

' Use this function to update all database files within a folder, and any subfolders of that folder
Public Function UpdateDbs()
  RunSearch "C:\Temp", "SearchFor", "ReplaceWith"

  Debug.Print "Finished"
End Function

'Requires a reference to Microsoft Scripting Runtime
Private Sub RunSearch(folderPath As String, pattern As String, newText As String)
On Error GoTo Err
  Dim FileSystem As New FileSystemObject
  Dim Folder As Folder
  Set Folder = FileSystem.GetFolder(folderPath)
  
  Dim File As File
  For Each File In Folder.Files
        ' This File.Type is specific for different OS versions, you may have better luck looking at the extension or something else
    If File.Type = "Microsoft Access Database" Then
      Debug.Print "Checking database: " & File.Path
      
      UpdateConnectionStrings OpenDatabase(File.Path), pattern, newText
    End If
  Next
  
  Dim childFolder As Folder
  For Each childFolder In Folder.SubFolders
    RunSearch childFolder.Path, pattern, newText
  Next
Exit Sub
Err:
  Debug.Print Err.Description
End Sub

Private Function UpdateConnectionStrings(db As DAO.Database, pattern As String, newText As String)
  Dim tblDef As TableDef
  For Each tblDef In db.TableDefs
    If InStr(LCase(tblDef.Connect), LCase(pattern)) > 0 Then
      Debug.Print "  Found match: " & tblDef.Name & " - " & tblDef.Connect
      tblDef.Connect = Replace(tblDef.Connect, pattern, newText)
      tblDef.RefreshLink
    End If
  Next
End Function
