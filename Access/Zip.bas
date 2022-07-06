Attribute VB_Name = "Zip"
Option Compare Database
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
  Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

public Sub TestZipFolder()
  'ZapFolder takes two arguments, the path to the destination file, and the folder that contains the items to zip
  ZipFolder "C:\Temp\Zip.zip", "C:\Temp\Zip"
 End Sub
 
Sub ZipFolder(zipFilePath As Variant, sourceFolder As Variant)
  If (Dir(sourceFolder, vbDirectory) = "") Then
    Err.Raise 1234, "ZipFolder", "Source folder (" & sourceFolder & ") does not exist, please check the path and try again."
  End If
  
  If (Dir(Left(zipFilePath, InStrRev(zipFilePath, "\") - 1), vbDirectory) = "") Then
    Err.Raise 1234, "ZipFolder", "Destination folder (" & Left(zipFilePath, InStrRev(zipFilePath, "\") - 1) & ") does not exist, please check the path and try again."
  End If
  
  '-------------------Create new empty Zip File-----------------
  If Len(Dir(zipFilePath)) > 0 Then Kill zipFilePath
  Open zipFilePath For Output As #1
  Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
  Close #1
  '=============================================================

  Dim childItemCount As Integer
  Dim shellApp As Object
  Set shellApp = CreateObject("Shell.Application")
  Dim fileSystem As Object
  Set fileSystem = CreateObject("Scripting.FileSystemObject")
  
  'Loop through the files in the selected folder
  Dim childFile
  For Each childFile In fileSystem.GetFolder(sourceFolder).Files
    shellApp.NameSpace(zipFilePath).CopyHere childFile.path
    childItemCount = childItemCount + 1
    
    Do Until shellApp.NameSpace(zipFilePath).items.Count = childItemCount
      'Wait until the item has been added
      Sleep 100
      DoEvents
    Loop
  Next
  
  Dim childFolder
  For Each childFolder In fileSystem.GetFolder(sourceFolder).SubFolders
    shellApp.NameSpace(zipFilePath).CopyHere childFolder.path
    childItemCount = childItemCount + 1
  
    Do Until shellApp.NameSpace(zipFilePath).items.Count = childItemCount
      'Wait until the item has been added
      Sleep 100
      DoEvents
    Loop
  Next
  Set shellApp = Nothing
  Set fileSystem = Nothing
End Sub
