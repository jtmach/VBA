Attribute VB_Name = "LockUnlock Functions"
Option Compare Database
Option Explicit

Function ChangeProperty(strDbName As String, strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
On Error GoTo Err
  'This will enable/disable the shift key when opening a database.
  'Holding the shift key will allow the user to bypass any startup functions
  'ChangeProperty "YourDbName", "AllowBypassKey", dbBoolean, True

  Dim mWsp As Workspace, dbs As Database, prp As Property
  Const conPropNotFoundError = 3270

  Set mWsp = DBEngine.Workspaces(0)
  Set dbs = mWsp.OpenDatabase(strDbName)
  dbs.Properties(strPropName) = varPropValue
  ChangeProperty = True

Exit Function
Err:
  If Err = conPropNotFoundError Then  ' Property not found.
    Set prp = dbs.CreateProperty(strPropName, varPropType, varPropValue)
    dbs.Properties.Append prp
    Resume Next
  Else ' Unknown error.
    ChangeProperty = False
  End If
End Function

Public Function Test()
  'This is an example of how you would call the ChangeProperty function
  'This will lock the database
  ChangeProperty "C:\Temp\Test.accdb", "AllowBypassKey", dbBoolean, False
  'This will unlock the database
  'ChangeProperty "C:\Temp\Test.accdb", "AllowBypassKey", dbBoolean, True
End Function
