Attribute VB_Name = "Module1"
Public acon As New ADODB.Connection

Public CON As Boolean

Public Function connect_db() As Boolean
On Error GoTo errtrap
With acon
    .ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & App.Path & "\EMPDB.mdb;UserId=Admin;password="
    .Open
    
    connect_db = True
    CON = True
    Exit Function
End With
errtrap:
  MsgBox Err.Description
  
End Function
