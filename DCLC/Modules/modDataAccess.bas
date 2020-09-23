Attribute VB_Name = "modDataAccess"
Option Explicit

'for searching/retrieving records using the SELECT SQL Statement.
'can also be used for adding, deleting & updating records.
Public Function GetRecordset(ByVal paramSql As String) As ADODB.Recordset
    On Error GoTo errhandler
    
    Set GetRecordset = New ADODB.Recordset
    GetRecordset.CursorLocation = adUseClient
    'GetRecordset.Open strSql, adoConnection, adOpenDynamic, adLockOptimistic
    GetRecordset.Open paramSql, adoConnection, adOpenStatic, adLockOptimistic, adCmdText
    Exit Function
errhandler:
    'Set GetRecordset = New ADODB.Recordset
    'Call MsgBox("GetRecordset Error.", vbOKOnly + vbCritical, "Error")
End Function

'for adding, deleting & updating using the INSERT, DELETE & UPDATE SQL Statements
Public Sub RunSql(strSql As String)
    Set adoCommand = New ADODB.Command
    adoConnection.BeginTrans
    adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strSql 'CleanUpSql(strSql)
    adoCommand.Execute
    adoConnection.CommitTrans
End Sub

'Increment primary key
Public Sub UpdatePK(ByVal srcField As String, ByVal strLastOperation As String)
    If LCase(srcField) = "studentno" And (LCase(strLastOperation) = "insert" Or LCase(strLastOperation) = "save") Then
        RunSql ("UPDATE dbo.LastNo SET " & srcField & " = " & srcField & " + 1; ")
        'Debug.Print "UPDATE dbo.LastNo SET " & srcField & " = " & srcField & " + 1;"
    ElseIf LCase(srcField) = "assessno" And (LCase(strLastOperation) = "insert" Or LCase(strLastOperation) = "save" Or LCase(strLastOperation) = "commit") Then
        RunSql ("UPDATE dbo.LastNo SET " & srcField & " = " & srcField & " + 1; ")
    ElseIf LCase(srcField) = "receiptno" And (LCase(strLastOperation) = "insert" Or LCase(strLastOperation) = "save") Then
        RunSql ("UPDATE dbo.LastNo SET " & srcField & " = " & srcField & " + 1; ")
    End If
End Sub



