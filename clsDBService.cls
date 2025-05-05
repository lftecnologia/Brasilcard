VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mstrLastError           As String

'Public ConnectionString    As String
Public CursorLocation       As Integer
Public AccessOption         As Byte

Private mConn As ADODB.Connection

Private Sub Class_Initialize()
    CursorLocation = adUseClient
End Sub

Sub CloseConnection()
    On Error Resume Next
    If Not mConn Is Nothing Then
        mConn.Close
        Set mConn = Nothing
    End If
End Sub

Function ExecuteSqlWithRec(ByVal pSqlStatement As String, Optional pCursorLocation As Integer = -1, Optional Conexao As ADODB.Connection = Nothing) As ADODB.Recordset
   On Error GoTo TrataErro
   
   mstrLastError = ""
   
   'Call Util.WriteFile("C:\t.txt", pSqlStatement, True)
   
   If Conexao Is Nothing Then
   
      If pCursorLocation <> -1 Then
         CloseConnection
         Set mConn = CreateObject("ADODB.Connection")
         mConn.ConnectionString = GetConnection
         mConn.CursorLocation = pCursorLocation
         mConn.Open
      Else
         If mConn Is Nothing Then
            Set mConn = CreateObject("ADODB.Connection")
            mConn.ConnectionString = GetConnection
            mConn.CursorLocation = CursorLocation
            mConn.Open
         ElseIf mConn.State = adStateClosed Then
            mConn.ConnectionString = GetConnection
            mConn.CursorLocation = CursorLocation
            mConn.Open
         End If
      End If
      Set ExecuteSqlWithRec = mConn.Execute(pSqlStatement)
   Else
      Set ExecuteSqlWithRec = Conexao.Execute(pSqlStatement)
   End If
   
   Exit Function

TrataErro:
    'StoreErrorLog pSqlStatement, "DbService", "ExecuteSqlWithRec"
    'MsgBox pSqlStatement
    'MsgBox mstrLastError
    'Call Util.WriteFile("C:\t.txt", pSqlStatement, True)
    mstrLastError = Err.Description
    'If Config.DbType = PostgreSQL Then MsgBox mstrLastError
    Set ExecuteSqlWithRec = Nothing
    'Resume
    Call TrataErro(Err.Number, Conexao)
End Function

Sub ExecuteSql(ByVal pSqlStatement As String, Optional ByRef pRecordsAffected As Integer = -1, _
               Optional Conexao As ADODB.Connection = Nothing)
               
    On Error GoTo TrataErro
    Dim conn As ADODB.Connection
    mstrLastError = ""
    pRecordsAffected = -1
    'Call Util.WriteFile("c:\t.txt", pSqlStatement, True)
    If Conexao Is Nothing Then
       Set conn = CreateObject("ADODB.Connection")
       conn.ConnectionString = GetConnection()
       conn.CursorLocation = CursorLocation
       conn.Open

       conn.Execute pSqlStatement, pRecordsAffected
       conn.Close
       Set conn = Nothing
    Else
       Conexao.Execute pSqlStatement, pRecordsAffected
    End If
    
    Exit Sub

TrataErro:
    mstrLastError = Err.Description
    'Call Util.WriteFile("c:\teste.txt", pSqlStatement, True)
    'StoreErrorLog pSqlStatement, "DbService", "ExecuteSql"
    'MsgBox Err.Description
'    If Config.DbType = PostgreSQL Then MsgBox err.Description
    'Resume    'Set Conn = Nothing
End Sub

Function CreateInsertStatement(pTableName As String, pFields As String, ParamArray ColumnValues()) As String
    'On Error GoTo TrataErro
    Dim strSQL  As String
    Dim i       As Integer
    
    If Trim$(pFields) <> "" Then
        strSQL = "INSERT INTO " & pTableName & " (" & pFields & ") VALUES("
    Else
        strSQL = "INSERT INTO " & pTableName & " VALUES("
    End If
    
    For i = LBound(ColumnValues) To UBound(ColumnValues)
        If ColumnValues(i) = "@@" Then
            'Dim rsTemp As ADODB.Recordset
            'Dim lngID  As Long
            'Set rsTemp = Conn.Execute("SELECT MAX(id) AS MaxID FROM " & pTableName)
            'lngID = Val(rsTemp!MaxID) + 1
            'strSql = strSql & CStr(lngID) & ", "
            'rsTemp.Close
            'Set rsTemp = Nothing
        Else
            Debug.Print TypeName(ColumnValues(i))
            Select Case TypeName(ColumnValues(i))
            Case "Byte", "Integer", "Long"
                strSQL = strSQL & CStr(ColumnValues(i)) & ", "
            Case "Single", "Double", "Currency"
                strSQL = strSQL & Replace(CStr(ColumnValues(i)), ",", ".") & ", "
            Case "String"
                strSQL = strSQL & "'" & CStr(ColumnValues(i)) & "', "
            Case "Null"
                strSQL = strSQL & "Null" & ", "
            Case "Date"
                strSQL = strSQL & "'" & Format(ColumnValues(i), FormatoData) & "', "
            Case "Boolean"
                strSQL = strSQL & CStr(ColumnValues(i)) & ", "
            Case Else
                strSQL = strSQL & "'" & CStr(ColumnValues(i)) & "', "
            End Select
        End If
    Next
    CreateInsertStatement = Left(strSQL, Len(strSQL) - 2) & ")"
    Exit Function

TrataErro:
    CreateInsertStatement = ""
    'If Not (rsTemp Is Nothing) Then Set rsTemp = Nothing
    'Call TrataErro(Err.Number, Conexao)
End Function

Function CreateUpdateStatement(ByVal pTableName As String, ByVal pWhereStatement As String, ParamArray ColumnValues()) As String
    On Error GoTo TrataErro
    Dim strSQL As String
    Dim i As Byte
    Dim strTypeName As String
    Dim strFieldName As String

    strSQL = "UPDATE " & pTableName & " SET "
    For i = 0 To UBound(ColumnValues) Step 2
        strFieldName = ColumnValues(i)
        strTypeName = TypeName(ColumnValues(i + 1))

        strSQL = strSQL & strFieldName & " = "
        Select Case strTypeName
        Case "Byte", "Integer", "Long"
            strSQL = strSQL & CStr(ColumnValues(i + 1)) & ", "
        Case "Single", "Double"
            strSQL = strSQL & Replace(CStr(ColumnValues(i + 1)), ",", ".") & ", "
        Case "String"
            strSQL = strSQL & "'" & CStr(ColumnValues(i + 1)) & "', "
        Case "Null"
            strSQL = strSQL & "Null" & ", "
        Case "Date"
            strSQL = strSQL & ftDat(ColumnValues(i + 1)) & ", "
        Case "Boolean"
            strSQL = strSQL & CStr(ColumnValues(i + 1)) & ", "
        Case Else
            strSQL = strSQL & "'" & CStr(ColumnValues(i + 1)) & "', "
        End Select
    Next
    CreateUpdateStatement = Left(strSQL, Len(strSQL) - 2) & " WHERE " & pWhereStatement
    'Call Util.WriteFile("c:\t.txt", CreateUpdateStatement, True)
    Exit Function

TrataErro:
    CreateUpdateStatement = Err.Description
    'Call TrataErro(Err.Number, Conexao)
End Function

Function InsertSql(ByVal pTableName As String, ByVal pFields As String, ParamArray ColumnValues()) As Integer
    On Error GoTo TrataErro
    Dim strSQL As String
    Dim intRecordsAffected As Integer

    mstrLastError = ""
    
    intRecordsAffected = -1
    strSQL = CreateInsertStatement(pTableName, pFields, ColumnValues())
    ExecuteSql strSQL, intRecordsAffected
    InsertSql = intRecordsAffected
    Exit Function

TrataErro:
    mstrLastError = Err.Description
    'Call TrataErro(Err.Number, Conexao)
End Function

Function GetUniqueValue(ByVal pSqlStatement As String, Optional Conexao As ADODB.Connection = Nothing) As String
    On Error GoTo TrataErro
    Dim rs As ADODB.Recordset

    mstrLastError = ""
    'Call Util.WriteFile("c:\t.txt", pSqlStatement, True)
    Set rs = ExecuteSqlWithRec(pSqlStatement, -1, Conexao)
    If Not Error Then
        If Not (rs.BOF And rs.EOF) Then
            GetUniqueValue = NZ(rs.Fields(0).value)
        End If
        rs.Close
    Else
        GetUniqueValue = ""
    End If
    
    Set rs = Nothing
    Exit Function

TrataErro:
    Set rs = Nothing
    mstrLastError = Err.Description
    Call TrataErro(Err.Number, Conexao)
End Function

'Function ExecuteSqlWithTransaction(ByRef pErro As String, ByVal ParamArray ColumnValues() As Object) As Boolean
'    Dim conn As ADODB.Connection
'    Dim cmd As SqlCommand
'    Dim i As Integer

'    ExecuteSqlWithTransaction = False
'    Try
'        conn = New SqlConnection(strConnectionString)
'        conn.Open()
'        cmd = conn.CreateCommand()
'        cmd.Transaction() = conn.BeginTransaction()
'        For i = 0 To ColumnValues.Length - 1
'            cmd.CommandText = ColumnValues(i)
'            cmd.ExecuteNonQuery()
'        Next
'        cmd.Transaction.Commit()
'        conn.Close() : conn = Nothing : cmd = Nothing
'    Catch Erro As Exception
'        pErro = Erro.Message
'        cmd.Transaction.Rollback()
'        conn.Close() : conn = Nothing : cmd = Nothing
'    End Try
'    ExecuteSqlWithTransaction = True
'End Function

Sub ExecuteSqlWithTransaction(ColumnValues() As String)
    On Error GoTo TrataErro
    Dim conn As ADODB.Connection
    Dim i As Integer
    Dim bytCheckPoint As Byte
    Dim strTemp As String
    
    mstrLastError = ""
    
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = GetConnection
    conn.CursorLocation = CursorLocation
    conn.Open
    
    conn.BeginTrans
    bytCheckPoint = 1
    
    For i = 0 To UBound(ColumnValues)
        If ColumnValues(i) <> "" Then
            strTemp = ColumnValues(i)
            'Debug.Print strTemp
            conn.Execute ColumnValues(i)
        End If
    Next
    conn.CommitTrans
    
    conn.Close
    Set conn = Nothing
    Exit Sub

TrataErro:
    If bytCheckPoint = 1 Then conn.RollbackTrans
    mstrLastError = Err.Description
    'StoreErrorLog strTemp, "DbService", "ExecuteSqlWithTransaction"
    If conn.State = adStateClosed Then conn.Close
    Set conn = Nothing
    'Call TrataErro(Err.Number, Conexao)
    On Error GoTo 0
End Sub

Private Function NZ(value As Variant) As String
    If value & "" = "" Then
        NZ = ""
    Else
        NZ = CStr(value)
    End If
End Function

Private Function GetConnection() As String
    On Error GoTo TrataErro
    Dim strconn As String
    
    mstrLastError = ""
    
'   If Config.IntegratedSecurity = "SSPI" Then
'       strconn = "Provider=SQLOLEDB;Server=" & Config.Server & ";Database=" & Config.DatabaseName & ";Integrated Security=SSPI;Application Name=" & App.Title
'   Else
       strconn = "Provider=SQLOLEDB;Data Source=#Source;Initial Catalog=#Name;User ID=#User;Password=#Pwd"
       strconn = Replace(strconn, "#Source", ".\sqlexpress")
       strconn = Replace(strconn, "#Name", "Cartao")
       strconn = Replace(strconn, "#User", "sa")
       strconn = Replace(strconn, "#Pwd", "12")
'   End If
        
    GetConnection = strconn
    Exit Function

TrataErro:
    mstrLastError = Err.Description
 '   Call TrataErro(Err.Number, Conexao)
End Function

Public Function TestConnection(Optional vEndOnError As Boolean) As Boolean
    On Error GoTo TrataErro
    Dim strconn As String
    Dim oConn As ADODB.Connection
    
    mstrLastError = ""
    
        
'   If Config.IntegratedSecurity = "SSPI" Then
'       strconn = "Provider=SQLOLEDB;Server=" & Config.Server & ";Database=" & Config.DatabaseName & ";Integrated Security=SSPI;Application Name=" & App.Title
'   Else
       Dim Provider As String
       strconn = "Provider=SQLOLEDB;Data Source=#Source;Initial Catalog=#Name;User ID=#User;Password=#Pwd"
       strconn = Replace(strconn, "#Source", ".\sqlexpress")
       strconn = Replace(strconn, "#Name", "Cartao")
       strconn = Replace(strconn, "#User", "sa")
       strconn = Replace(strconn, "#Pwd", "12")
'   End If
    
    Set oConn = CreateObject("ADODB.Connection")
    oConn.Open strconn
    oConn.Close
    Set oConn = Nothing
    
    TestConnection = True
    Exit Function

TrataErro:
    mstrLastError = Err.Description
    'If vEndOnError Then Call TrataErro(Err.Number, Conexao)
End Function

Public Property Get GetLastError() As String
    GetLastError = mstrLastError
End Property

Public Property Get Error() As Boolean
    Error = Trim$(mstrLastError <> "")
End Property

Private Sub TrataErro(erro As String, ByRef Conexao As ADODB.Connection)
   '#ERRODEREDE
   If erro = "-2147467259" Then
      MsgBox "Erro de rede, o sistema será fechado!", vbCritical, "Perda de Conexão de Rede"
      End
   End If
End Sub


