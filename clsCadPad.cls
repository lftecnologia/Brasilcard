VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCadPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type campos
   nome As String
   Tipo As Integer
End Type

Private sTabela As String     ' Nome da tabela
Private sChave As String      ' Chave da tabela
Private sCampoSeq As String   ' nome do campo Cod Sequencial
Private nCodigo As Double     ' codigo informado
Private sStringAux As String  ' string auxiliar para concatencação de inclusao, alteracao
Private bEmpresa As Boolean   ' true=se for por empresa, false=não
Private sCondicao As String

Private f As Form

Private oConexao As cConecta
Private Conexao As ADODB.Connection
Private rsTabela As ADODB.Recordset


Public Util As cUtil
Public DbService As cDBService

Public Property Get Tabela() As String
   Tabela = sTabela
End Property

Public Property Let Tabela(ByVal PMT As String)
   sTabela = PMT
End Property

Public Property Get Chave() As String
   Chave = sChave
End Property

Public Property Let Chave(ByVal PMT As String)
   sChave = PMT
   If sCampoSeq = "" Then sCampoSeq = sChave
End Property

Public Property Get CampoSeq() As String
   CampoSeq = sCampoSeq
End Property

Public Property Let CampoSeq(ByVal PMT As String)
   sCampoSeq = PMT
End Property

Public Property Get codigo() As Double
   codigo = nCodigo
End Property

Public Property Let codigo(ByVal PMT As Double)
   nCodigo = PMT
End Property

Public Property Get stringAux() As String
   stringAux = sStringAux
End Property

Public Property Let stringAux(ByVal PMT As String)
   'sintaxe
   'campo1=valor, campo2=valor, campo3=valor
   sStringAux = PMT
End Property

Public Property Let formulario(ByVal PMT As Form)
   If Not (PMT Is Nothing) Then
      Set f = PMT
   End If
End Property

Public Property Let Empresa(ByVal PMT As Boolean)
   bEmpresa = PMT
End Property

Public Property Let condicao(ByVal PMT As String)
   sCondicao = PMT
End Property

Public Function Grava(Optional Conexao1 As ADODB.Connection = Nothing, Optional ByRef ID As Long) As Boolean
   Dim sCampos As String, sValores As String
   Dim sCampoAdi As String, sValorAdi As String
   Dim sAltera As String, vFieldReq As String
   
   On Error GoTo erro
   
   Set DbService = New cDBService
   
   sAltera = retCampos(sCampos, sValores, vFieldReq)
   
   If vFieldReq <> "" Then
      rsTabela.Close
      Exit Function
   End If
   
   If f.Tag = "NEW" Then
      formatSQLAdi sCampoAdi, sValorAdi
      
      If sCampos <> "" Then
         sCampos = Mid(sCampos, 1, Len(sCampos) - 1)
         sValores = Mid(sValores, 1, Len(sValores) - 1)
      End If
      
      If nCodigo = 0 Then
         rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
         
         If rsTabela.RecordCount = 0 Then
            nCodigo = 1
         Else
            nCodigo = CDbl(0 & rsTabela!n) + 1
         End If
         rsTabela.Close
      End If
      Dim v As String
      
      ID = DbService.GetUniqueValue("INSERT INTO " & sTabela _
      & "(" & sCampos & sCampoAdi & ") OUTPUT INSERTED.Id_Transacao VALUES " _
      & "(" & sValores & sValorAdi & ")", _
      IIf(Conexao1 Is Nothing, Conexao, Conexao1))
      
   Else
      If sAltera <> "" Then   'quando existe apenas 1 campo e este é chave
         sAltera = Mid(sAltera, 1, Len(sAltera) - 1)
         If sStringAux <> "" Then sStringAux = "," & sStringAux
         If bEmpresa = False Then
            If sCondicao <> "" Then
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE " & sCondicao & " and " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            Else
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            End If
         End If
         If DbService.Error Then MsgBox "Erro ao alterar registro!", vbCritical, "Atenção!"
      End If
   End If
   
   'ID = nCodigo
   
   Grava = Not DbService.Error
   
   Exit Function
   
erro:
   Call Util.logSistema("User", sTabela, nCodigo, vEmpresa, Err.Description)
   MsgBox Err.Description, vbCritical, "Atenção!"
   Grava = False
End Function

Public Function Exclui(ID, Optional bPergunta As Boolean = False, Optional Conexao1 As ADODB.Connection = Nothing) As Boolean
    On Error GoTo erro
    
    Set DbService = New cDBService
      
    If Trim$(CStr(ID)) = "" Then
        MsgBox "Nenhum registro foi selecionado!", vbExclamation, "Exclusão não realizada!"
        Exit Function
    End If
    
    If Not bPergunta Then
        If 7 = MsgBox("Confirma exclusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!!") Then Exit Function
    End If
        
    Call DbService.ExecuteSql("DELETE FROM " & sTabela & " Where " & sCampoSeq & " = " & ID, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
    
    Exclui = Not DbService.Error
    
    Exit Function
               
erro:
   Call Util.logSistema("User", sTabela, ID, vEmpresa, Err.Description)
   Call trataErros(Err.Number)
End Function

Private Sub Class_Initialize()
   Set oConexao = New cConecta
   Set Conexao = CreateObject("ADODB.Connection")
   Set rsTabela = New ADODB.Recordset
      
   Set Conexao = oConexao.Conexao
End Sub

Private Sub Class_Terminate()
'   Conexao.Close
   Set rsTabela = Nothing
   Set Conexao = Nothing
End Sub

Private Sub formatSQLAdi(ByRef sCampoAdi As String, ByRef sValorAdi As String)
   'formata sql adicional (p/ objetos não vinculados)
   Dim i As Long
   Dim retAux As Integer, pIgual As Integer
   
   sCampoAdi = ""
   sValorAdi = ""
   If stringAux = "" Then Exit Sub
   i = 1
   Do
      If f.Tag = "NEW" Then
         retAux = InStr(i, stringAux, ",")
         If retAux = 0 And sCampoAdi = "" Then
            sCampoAdi = ", " & Trim(Mid(stringAux, i, InStr(i, stringAux, "=") - 1))
            sValorAdi = ", " & Trim(Mid(stringAux, InStr(i, stringAux, "=") + 1, Len(stringAux)))
            Exit Do
         Else
            pIgual = InStr(i, stringAux, "=")
            sCampoAdi = sCampoAdi & ", " & Trim(Mid(stringAux, i, pIgual - i))
            If retAux = 0 Then
               sValorAdi = sValorAdi & ", " & Trim(Mid(stringAux, pIgual + 1, Len(stringAux)))
               Exit Do
            Else
               sValorAdi = sValorAdi & ", " & Trim(Mid(stringAux, pIgual + 1, retAux - pIgual - 1))
            End If
         End If
         i = 1 + retAux
      End If
   Loop Until i = Len(sStringAux)
End Sub

Private Function formataCampo(ByRef Controles As Control, ByRef vFieldReq As Boolean, ByRef vFieldDes As String, ByRef vFieldMid As Integer) As String
   Dim strNmCampo As String, vCampo As String
   
   Set Util = New cUtil
   
   If Right(Controles.Tag, 1) = ";" Then
      MsgBox "Campo '" & Controles.Name & "' possui separador ; sem o devido argumento!", vbCritical, "Atenção!"
      Exit Function
   End If
   vFieldMid = 0
   vCampo = Controles.Tag
   ' retira o parametro de campo requerido da Tag
   If InStr(vCampo, ";") > 0 Then
      vCampo = Util.getFieldCSV(Controles.Tag, ";", 0)
      vFieldReq = Util.getFieldCSV(Controles.Tag, ";", 1)
   End If
   
   If Util.getFieldCSV(Controles.Tag, ";", 2) <> "" Then
     vFieldDes = Util.getFieldCSV(Controles.Tag, ";", 2)
   End If
      
   If Util.getFieldCSV(Controles.Tag, ";", 3) <> "" Then
     vFieldMid = IIf(IsNumeric(Util.getFieldCSV(Controles.Tag, ";", 3)), Util.getFieldCSV(Controles.Tag, ";", 3), 0)
   End If
      
   If TypeOf Controles Is ComboBox Or TypeOf Controles Is DataCombo Then
      'verifica se tem -
      If Mid(Controles.Tag, Len(Controles.Tag), 1) = "_" Then
         strNmCampo = Mid(Controles.Tag, 1, Len(Controles.Tag) - 1)
      Else
         strNmCampo = Controles.Tag
      End If
   Else
      'qualquer outro tipo de obj
      strNmCampo = Controles.Tag
   End If
   
   formataCampo = Mid(Util.getFieldCSV(strNmCampo, ";", 0), 5, Len(Util.getFieldCSV(strNmCampo, ";", 0)))
End Function

Private Function Consulta(ID As Double)
   On Error Resume Next
   
   Dim vFieldReq As Boolean, vFieldDes As String, vFieldMid As Integer
   
   If sCondicao <> "" Then
      rsTabela.Open "Select * from " & sTabela & " Where " & sCondicao & " And " & sCampoSeq & " = " & ID, Conexao, adOpenDynamic, adLockOptimistic
   Else
      rsTabela.Open "Select * from " & sTabela & " Where " & sCampoSeq & " = " & ID, Conexao, adOpenDynamic, adLockOptimistic
   End If
   
   Dim Controles As Control
   Dim i As Integer
   Dim strCampo As String
   For Each Controles In f.Controls
       If TypeOf Controles Is TextBox Or TypeOf Controles Is DataCombo Or _
          TypeOf Controles Is CheckBox Or TypeOf Controles Is ComboBox Or TypeOf Controles Is Label Or _
          TypeOf Controles Is DTPicker Or TypeOf Controles Is OptionButton Then
          If Mid(UCase(Controles.Tag), 1, 3) = "FLD" Then
             For i = 0 To rsTabela.Fields.Count - 1
                 strCampo = formataCampo(Controles, vFieldReq, vFieldDes, vFieldMid)
                 If UCase(strCampo) = UCase(rsTabela(i).Name) Then
                    If TypeOf Controles Is DataCombo Then         'datacombo
                       If IsNumeric(rsTabela(i).value) Then
                           Controles.BoundText = CLng(0 & rsTabela(i).value)
                       Else
                           Controles.BoundText = "" & rsTabela(i).value
                       End If
                    ElseIf TypeOf Controles Is CheckBox Then      'checkbox
                       Controles.value = CSng(0 & rsTabela(i).value)
                    ElseIf TypeOf Controles Is OptionButton Then  'option
                       Controles.value = IIf(CSng(0 & rsTabela(i).value) = False, 0, 1)
                    ElseIf TypeOf Controles Is Label Then         'label
                       If rsTabela(i).Type = adCurrency Then
                           Controles.Caption = "" & ftValor(rsTabela(i).value, 2)
                       Else
                           Controles.Caption = "" & rsTabela(i).value
                       End If
                    ElseIf TypeOf Controles Is ComboBox Then      'combobox
                        If vFieldMid > 0 Then
                           Controles.ListIndex = Util.SetaCombo(Controles, "" & rsTabela(i).value)
                        Else
                           If Not (IsNull(rsTabela(i).value) Or rsTabela(i).value = "") Then
                              Controles.text = rsTabela(i).value
                           Else
                              Controles.text = ""
                           End If
                        End If
                    ElseIf TypeOf Controles Is DTPicker Then
                       If IsNull(rsTabela(i)) Then
                          Controles = 0
                       Else
                          Controles = rsTabela(i)
                       End If
                    Else
                       If Not (IsNull(rsTabela(i).value) Or rsTabela(i).value = "") Then
                          Controles = Trim(rsTabela(i).value)
                       Else
                          Controles = ""
                       End If
                    End If
                 End If
             Next
          End If
       End If
   Next
   
   nCodigo = ID
   f.Tag = "EDIT"
      
   rsTabela.Close
   
End Function

Public Function Primeiro() As Double
   On Error GoTo erro
   
   Dim ID As Double
   
   Dim sSql As String
   sSql = "SELECT MIN(" & sCampoSeq & ") AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " > 0"
   
   If sCondicao <> "" Then
      sSql = sSql & " AND " & sCondicao
   End If
   
   With rsTabela
        .Open sSql, Conexao, adOpenForwardOnly, adLockReadOnly
        
        If IsNull(!codigo) Then
           ID = 0
        Else
           ID = !codigo
        End If
        .Close
        If ID <> 0 Then
           Call Consulta(ID)
           Primeiro = ID
        End If
   End With
   
   Exit Function
   
erro:
   If Err.Number = -2147467259 Then
      MsgBox "Erro de Conexão: A rede foi perdida, você deve reiniciar o sistema!", vbCritical, "Conexão de Rede"
   Else
      MsgBox "Erro de Conexão: " & Err.Description, vbCritical, "Conexão de Rede"
   End If
End Function

Public Function Anterior(ID) As Double
   On Error GoTo erro
   
   If Val(ID) = 0 Then Call Primeiro: Exit Function
   
   Dim sSql As String
   sSql = "SELECT MAX(" & sCampoSeq & ") as Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " < " & ID
   
   If sCondicao <> "" Then
      sSql = sSql & " AND " & sCondicao
   End If
   
   With rsTabela
        .Open sSql, Conexao, adOpenForwardOnly, adLockReadOnly
        
        If .EOF Or .BOF Then
           ID = 0
        Else
           ID = 0 & !codigo
        End If
        .Close
        If ID <> 0 Then
           Call Consulta(Val(ID))
           Anterior = ID
        End If
   End With
   
   Exit Function
   
erro:
   If Err.Number = -2147467259 Then
      MsgBox "Erro de Conexão: A rede foi perdida, você deve reiniciar o sistema!", vbCritical, "Conexão de Rede"
   Else
      MsgBox "Erro de Conexão: " & Err.Description, vbCritical, "Conexão de Rede"
   End If
End Function

Public Function Proximo(ID) As Double
   On Error GoTo erro
   
   If Val(ID) = 0 Then Call Primeiro: Exit Function
   
   Dim sSql As String
   sSql = "SELECT MIN(" & sCampoSeq & ") AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " > " & ID
   
   If sCondicao <> "" Then
      sSql = sSql & " AND " & sCondicao
   End If
   
   With rsTabela
        .Open sSql, Conexao, adOpenForwardOnly, adLockReadOnly
        If IsNull(!codigo) Then
           ID = 0
        Else
           ID = !codigo
        End If
        .Close
        If ID <> 0 Then
           Call Consulta(Val(ID))
           Proximo = ID
        End If
   End With
   
   Exit Function
   
erro:
   If Err.Number = -2147467259 Then
      MsgBox "Erro de Conexão: A rede foi perdida, você deve reiniciar o sistema!", vbCritical, "Conexão de Rede"
   Else
      MsgBox "Erro de Conexão: " & Err.Description, vbCritical, "Conexão de Rede"
   End If
End Function

Public Function Ultimo() As Double
   On Error GoTo erro
   
   Dim ID As Double
   
   Dim sSql As String
   sSql = "SELECT MAX(" & sCampoSeq & ") AS Codigo FROM " & Tabela
   
   If sCondicao <> "" Then
      If bEmpresa = True Then
         sSql = sSql & " AND " & sCondicao
      Else
         sSql = sSql & " WHERE " & sCondicao
      End If
   End If
      
   With rsTabela
        .Open sSql, Conexao, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
           ID = 0
        Else
           ID = 0 & !codigo
        End If
        .Close
        If ID <> 0 Then
           Call Consulta(ID)
           Ultimo = ID
        End If
   End With
   
   Exit Function
   
erro:
   If Err.Number = -2147467259 Then
      MsgBox "Erro de Conexão: A rede foi perdida, você deve reiniciar o sistema!", vbCritical, "Conexão de Rede"
   Else
      MsgBox "Erro de Conexão: " & Err.Description, vbCritical, "Conexão de Rede"
   End If
End Function

Public Function Seta(ID As String) As Boolean
   'On Error GoTo erro
   
   ID = CDbl(0 & Trim(ID))
   If ID = 0 Then Exit Function
   
   Dim sSql As String
   sSql = "SELECT " & sCampoSeq & " AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " = " & ID
   
   If sCondicao <> "" Then
      sSql = sSql & " AND " & sCondicao
   End If
   
   With rsTabela
        .Open sSql, Conexao, adOpenDynamic, adLockReadOnly
        If .EOF Then
           ID = 0
        Else
           ID = !codigo
           Seta = True
        End If
        .Close
        If ID <> 0 Then Call Consulta(CDbl(ID))
   End With
   
   Exit Function
   
erro:
   If Err.Number = -2147467259 Then
      MsgBox "Erro de Conexão: A rede foi perdida, você deve reiniciar o sistema!", vbCritical, "Conexão de Rede"
   Else
      MsgBox "Erro de Conexão: " & Err.Description, vbCritical, "Conexão de Rede"
   End If
End Function

Private Function retCampos(sCampo As String, sValue As String, sFieldReq As String) As String
   Dim campo() As campos
   Dim sUpdate As String
   Dim i As Integer, sConteudo As String
   Dim vFieldReq As Boolean, vFieldDes As String, vFieldMid As Integer
   
   On Error GoTo erro
   
   sCampo = ""
   sValue = ""
   sFieldReq = ""
   
   If rsTabela.State = 1 Then rsTabela.Close
   rsTabela.Open "Select * from " & sTabela & " Where " & sCampoSeq & " < 0", Conexao, adOpenStatic, adLockOptimistic
   ReDim campo(rsTabela.Fields.Count) As campos
   
   Dim Controles As Control
   Dim strCampo As String
   For Each Controles In f.Controls
       If TypeOf Controles Is TextBox Or TypeOf Controles Is DataCombo Or _
          TypeOf Controles Is CheckBox Or _
          TypeOf Controles Is ComboBox Or TypeOf Controles Is Label Or _
          TypeOf Controles Is DTPicker Or TypeOf Controles Is OptionButton Then
             sConteudo = Replace(Controles, "'", "")
          End If
          If Mid(UCase(Controles.Tag), 1, 3) = "FLD" Then
             For i = 0 To rsTabela.Fields.Count - 1
                 If UCase(sCampoSeq) <> UCase(Mid(Controles.Tag, 5, Len(Controles.Tag))) Then
                    vFieldReq = False
                    vFieldDes = ""
                    strCampo = formataCampo(Controles, vFieldReq, vFieldDes, vFieldMid)
                    If UCase(strCampo) = UCase(rsTabela(i).Name) Then
                        If TypeOf Controles Is TextBox Then
                           If vFieldReq And Trim(Controles) = "" Then
                              sFieldReq = strCampo
                              MsgBox "Informação requerida não foi preenchida! " & vFieldDes, vbCritical, "Atenção!"
                              If Controles.Enabled And Controles.Visible Then Controles.SetFocus
                              Exit Function
                           End If
                        End If
                        If f.Tag = "NEW" Then
                           sCampo = sCampo & rsTabela(i).Name & ","
                        End If
                        sUpdate = sUpdate & rsTabela(i).Name & " = "
                        If sConteudo = "" Then           ' SE O CONTEUDO ESTIVER EM BRANCO
                           Select Case rsTabela(i).Type
                           Case adTinyInt, adSmallInt, adSingle, adInteger, adNumeric, adBigInt, adDouble
                             sUpdate = sUpdate & "0,"
                             sValue = sValue & "0,"
                           Case adDate, adDBDate, adDBTimeStamp
                                sUpdate = sUpdate & "null,"
                                sValue = sValue & "null,"
                           Case Else
                                 sUpdate = sUpdate & "'" & Null & "',"
                                 sValue = sValue & "'" & Null & "',"
                           End Select
                        Else                             ' SE O CONTEUDO FOI PREENCHIDO
                           If TypeOf Controles Is DataCombo Then
                              sUpdate = sUpdate & Controles.BoundText & ","
                              sValue = sValue & Controles.BoundText & ","
                           ElseIf TypeOf Controles Is CheckBox Then
                              sUpdate = sUpdate & Controles.value & ","
                              sValue = sValue & Controles.value & ","
                           ElseIf TypeOf Controles Is OptionButton Then
                              sUpdate = sUpdate & IIf(Controles.value = False, 0, 1) & ","
                              sValue = sValue & IIf(Controles.value = False, 0, 1) & ","
                           ElseIf TypeOf Controles Is ComboBox Then
                              If vFieldMid > 0 Then
                                 sUpdate = sUpdate & "'" & Mid(Controles.text, 1, vFieldMid) & "',"
                                 sValue = sValue & "'" & Mid(Controles.text, 1, vFieldMid) & "',"
                              Else
                                 If Mid(Controles.Tag, Len(Controles.Tag), 1) = "_" Then
                                    sUpdate = sUpdate & Controles.ListIndex & ","
                                    sValue = sValue & Controles.ListIndex & ","
                                 Else
                                    sUpdate = sUpdate & IIf(Controles.text = "", Null, IIf(rsTabela(i).Type = adChar Or rsTabela(i).Type = adVarChar Or rsTabela(i).Type = 202, "'" & Controles & "'", Controles.text)) & ","
                                    sValue = sValue & IIf(Controles.text = "", Null, IIf(rsTabela(i).Type = adChar Or rsTabela(i).Type = adVarChar Or rsTabela(i).Type = 202, "'" & Controles & "'", Controles.text)) & ","
                                 End If
                              End If
                           ElseIf TypeOf Controles Is DTPicker Then
                              sUpdate = sUpdate & ftDat(Controles) & ","
                              sValue = sValue & ftDat(Controles) & ","
                           Else
                              If rsTabela(i).Type = adDate Or rsTabela(i).Type = adDBDate Or rsTabela(i).Type = adDBTimeStamp Then
                                  sUpdate = sUpdate & ftDat(Controles) & ","
                                  sValue = sValue & ftDat(Controles) & ","
                              ElseIf rsTabela(i).Type = adSmallInt Or rsTabela(i).Type = adSingle Or rsTabela(i).Type = adInteger Or rsTabela(i).Type = adDouble Or rsTabela(i).Type = dbDecimal Then 'numerico sem decimais
                                  sUpdate = sUpdate & ftValor(Controles, 0, 1) & ","
                                  sValue = sValue & ftValor(Controles, 0, 1) & ","
                              ElseIf rsTabela(i).Type = adCurrency Or rsTabela(i).Type = adNumeric Then
                                  sUpdate = sUpdate & ftValor(Controles, 2, 1) & ","
                                  sValue = sValue & ftValor(Controles, 2, 1) & ","
                              Else
                                  sUpdate = sUpdate & "'" & Controles & "',"
                                  sValue = sValue & "'" & Controles & "',"
                              End If
                           End If
                        End If
                    End If
                 End If
             Next
          End If
       
   Next
   If f.Tag = "EDIT" Then retCampos = sUpdate
   rsTabela.Close
   
   Exit Function
   
erro:
   MsgBox Err.Description
End Function


Public Function Achou(sFiltro As String, Optional sTabelaDif As String) As Boolean
   Dim vSql As String
   
   If sTabelaDif = "" Then
      sTabelaDif = sTabela
   End If
      
   vSql = ""
   rsTabela.Open "Select * from " & sTabelaDif & " Where " & sFiltro & vSql, Conexao, adOpenStatic, adLockReadOnly
         
   If rsTabela.RecordCount > 0 Then Achou = True
   rsTabela.Close
End Function


