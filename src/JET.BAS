Attribute VB_Name = "JET"
Option Explicit

'Global xog As Integer
'Global xle As Integer
'Global Const xTE = 4
Sub MyJet()
   
    'Dim MyDb As Database
    'Dim MyTable As Recordset
    
    'On Error GoTo JetError
    
    'Dim Db As Database, Rs As Recordset
    'Set Db = OpenDatabase("NEWS.MDB")
    'Set Rs = Db.OpenRecordset("Groups")
    
'    Set MyDb = DBEngine.Workspaces(0).Databases("NEWS.MDB")
'    Set MyTable = MyDb.OpenRecordset("News", DB_OPEN_TABLE)
'    MyTable.Index = "PrimaryKey"
'    MyTable.Seek "=", Me![SSN#]
'    If Not MyTable.NoMatch Then
'        MyTable.Edit
'        MyTable("Last Name") = Me![Last Name]
'        MyTable("First Name") = Me![First Name]
      '  temp1 = MyTable("Middle Name")
'        MyTable.Update
    'MyTable.Close
Exit Sub

JetError:
    Select Case Err
'        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
        Case Else
'           MsgBox "This is an ISP connection error.  Please click on Help Send Feedback and send the attached message through your regular e-mail program.  This will enable you to obtain technical support f
    End Select
    Exit Sub
End Sub
