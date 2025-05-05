VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oConexao As ADODB.Connection
Private mstrLastError As String

'Opção para retornar a conexão do Banco Local ou Remoto (0 ou 1)
Private mintAccessOption As Integer

Public Property Get Conexao(Optional DbOption As Integer = 0) As ADODB.Connection
    Dim strconn As String
    
    mintAccessOption = DbOption
    
    Set oConexao = CreateObject("ADODB.Connection")
    
    SetConnection
    
    If Trim$(mstrLastError) = "" Then
        Set Conexao = oConexao
    End If
End Property

Private Sub SetConnection()
    On Error GoTo TrataErro
    Dim strconn As String
    
    mstrLastError = ""
    
   'If Config.IntegratedSecurity = "SSPI" Then
   '    oConexao.Open "Provider=SQLOLEDB;Server=" & Config.Server & ";Database=" & Config.DatabaseName & ";Integrated Security=SSPI;Application Name=" & App.Title
   'Else
       strconn = "Provider=SQLOLEDB;Data Source=#Source;Initial Catalog=#Name;User ID=#User;Password=#Pwd"
       strconn = Replace(strconn, "#Source", ".\sqlexpress")
       strconn = Replace(strconn, "#Name", "Cartao")
       strconn = Replace(strconn, "#User", "sa")
       strconn = Replace(strconn, "#Pwd", "12")
       oConexao.Open strconn
  ' End If
      
   oConexao.CursorLocation = adUseClient
   
   Exit Sub

TrataErro:
    mstrLastError = Err.Description
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    oConexao.Close
    Set oConexao = Nothing
End Sub

Public Property Get GetLastError() As String
    GetLastError = mstrLastError
End Property

Public Function GetStrCnn() As String
    'Dim obj As New DataLinks
    'GetStrCnn = obj.PromptNew
End Function

Public Property Get AccessOption() As Integer
    AccessOption = mintAccessOption
End Property




