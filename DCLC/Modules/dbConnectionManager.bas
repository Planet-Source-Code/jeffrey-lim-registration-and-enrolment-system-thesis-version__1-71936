Attribute VB_Name = "dbConnectionManager"
Option Explicit

'This procedure is to be called in every Form_Load event
Public Function ConnectToSqlServer(dbServer As String, dbUser As String, dbPassword As String) As Boolean
    On Error GoTo errhandler
    
    dbName = "Dclc"
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "Provider=MSDASQL;" & _
                       "DRIVER={SQL Server};" & _
                       "SERVER=" & dbServer & ";" & _
                       "trusted_connection=no;" & _
                       "user id=" & dbUser & ";" & _
                       "password=" & dbPassword & ";" & _
                       "database=" & dbName & ";"
    'adoConnection.Open "Provider=SQLOLEDB.1;User ID=" & dbUser & _
                                 ";Password=" & dbPassword & _
                                 ";Initial Catalog=" & dbName & _
                                 ";Data Source=" & dbServer & ";"
    ConnectToSqlServer = True   'return true for successful connection
    SaveSetting App.Title, "LOGIN", "SERVER", dbServer 'saves the value of the txtServer (server name)
    Exit Function
errhandler:
    ConnectToSqlServer = False  'return false for unsuccessful connection
    Call MsgBox("Connection Error.", vbOKOnly, "Database Connection Error")
    'MsgBox Err.Description, vbCritical, "Error12:" & Err.Number
End Function

Public Function ConnectToSqlServer3(dbServer As String, dbUser As String, dbPassword As String) As Boolean
    'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=Pc4;Initial Catalog=CSUMIS;Data Source=WINDOWS-7EC8A49
    On Error GoTo Erb
    With adoConnection
    If .State <> 0 Then .Close
        .ConnectionString = "Provider=MSDATASHAPE.1;data Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & dbUser & ";Password=" & dbPassword & ";Initial Catalog=CSUMIS;Data Source=" & dbServer
        .CursorLocation = adUseClient
        .Open
    End With
    ConnectToSqlServer3 = True
    'SetRs "Select * from names"
    Exit Function
Erb:
    MsgBox err.Description, vbCritical, "Error12:" & err.Number
    'adoConnection.Close
    ConnectToSqlServer3 = False
End Function

Public Sub CloseSqlServerConnection()
    adoConnection.Close
    Set adoConnection = Nothing
End Sub

