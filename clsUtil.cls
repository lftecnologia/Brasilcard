VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub WriteFile(ByVal vFile As String, ByVal vContent As String, Optional vOpenFile As Boolean, Optional AddLog As Boolean = False)
   ' Grava um arquivo texto com opção para visualizar pelo notepad
   
   Dim FileNumber As Integer
   On Error GoTo Trata
   
   FileNumber = FreeFile
   If AddLog And Dir("c:\fusion_log.log") <> "" Then
      Open vFile For Append As #FileNumber
      Print #FileNumber, vContent
   Else
      Open vFile For Output As #FileNumber
      Print #FileNumber, vContent
   End If
   Close #FileNumber
 
   If vOpenFile = True Then
      Shell "Notepad " & vFile, vbMaximizedFocus
   End If
   Exit Sub
   
Trata:
   trataErros Err.Number
   Close FileNumber
   'Resume
End Sub

Public Sub logSistema(ByVal Usuario As String, ByVal Tabela As String, ByVal ID As String, ByVal Empresa As String, ByVal Exception As String)
    On Error GoTo TrataErro

    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim novoId As Long
    Dim dataAtual As String
    Dim horaAtual As String
    Dim objCon As New cConecta

    ' Instancia conexão local
    Set cnn = objCon.Conexao(0) ' 0 = banco local

    ' Obter novo ID
    Set rs = cnn.Execute("SELECT ISNULL(MAX(ID), 0) + 1 AS NovoID FROM logsis")
    novoId = rs!novoId
    rs.Close
    Set rs = Nothing

    ' Preparar data e hora
    dataAtual = Format(Date, "yyyy-MM-dd")
    horaAtual = Format(Time, "HH:mm")

    ' Montar INSERT com tratamento de aspas
    strSQL = "INSERT INTO logsis (ID, LogData, LogHora, LogUsuario, LogTabela, LogTabelaID, Empresa, Exception) VALUES (" & _
             novoId & ", '" & dataAtual & "', '" & horaAtual & "', '" & _
             Replace(Usuario, "'", "''") & "', '" & Replace(Tabela, "'", "''") & "', '" & _
             Replace(ID, "'", "''") & "', '" & Replace(Empresa, "'", "''") & "', '" & _
             Replace(Exception, "'", "''") & "')"

    ' Executar
    cnn.Execute strSQL

    ' Encerrar conexão
    cnn.Close
    Set cnn = Nothing
    Exit Sub

TrataErro:
    MsgBox "Erro ao registrar log: " & Err.Description, vbCritical
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then cnn.Close
        Set cnn = Nothing
    End If
End Sub


Public Function getFieldCSV(ByVal vLin As String, ByVal vDelimiter As String, ByVal n As Integer) As String
   ' Retorna o campo n de uma linha de arquivo csv
   
   Dim vPalavra As String
   Dim vDelCont As Integer
   Dim i As Integer
   
   vDelCont = 0
   vPalavra = ""
   
   For i = 0 To Len(vLin)
      If vDelCont = n Then
         vPalavra = vPalavra & Mid(vLin, i + 1, 1)
      End If
      If Mid(vLin, i + 1, 1) = vDelimiter Then vDelCont = vDelCont + 1
      If vDelCont > n Then Exit For
   Next i
   'getFieldCSV = IIf(vPalavra = "", "", Replace(vPalavra, ";", ""))
   getFieldCSV = IIf(vPalavra = "", "", Replace(vPalavra, vDelimiter, ""))
End Function

Function SetaCombo(combo As ComboBox, ByVal vCod As String) As Integer
   Dim i As Integer
   
   SetaCombo = 0
   For i = 0 To combo.ListCount - 1
      If Mid(combo.List(i), 1, Len(vCod)) = vCod Then
         SetaCombo = i
         Exit For
      End If
   Next i
   combo.ListIndex = SetaCombo
End Function

Public Sub NumeraGrid(ByRef grdPad As DataGrid)
   Dim i As Integer
   For i = 1 To grdPad.Columns.Count
      If i < 10 Then
         grdPad.Columns(i - 1).Caption = "(" & i & ")" & StrConv(grdPad.Columns(i - 1).Caption, vbProperCase)
      Else
         grdPad.Columns(i - 1).Caption = "(" & Chr(55 + i) & ")" & StrConv(grdPad.Columns(i - 1).Caption, vbProperCase)
      End If
   Next
End Sub


