Attribute VB_Name = "cCadPad"
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

Public Property Let Condicao(ByVal PMT As String)
   sCondicao = PMT
End Property

Public Function Grava(Optional Conexao1 As ADODB.Connection = Nothing, Optional ByRef ID As Long) As Boolean
   Dim sCampos As String, sValores As String
   Dim sCampoAdi As String, sValorAdi As String
   Dim sAltera As String, vFieldReq As String
   
   On Error GoTo erro
   
   sAltera = retCampos(sCampos, sValores, vFieldReq)
   
   If vFieldReq <> "" Then
      rsTabela.Close
      Exit Function
   End If
   
   ' verificar inconsistência ao editar (leonardo)
   If f.Tag <> "NEW" Then
      If bEmpresa = False Then
         rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
      Else
         rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " WHERE EmpCodigo = " & EmpresaAtiva.ID, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
      End If
      If IsNull(rsTabela!n) Then
         MsgBox "A tabela não possui registros, clique em Novo para adicionar registros!", vbCritical, "Atenção!"
         rsTabela.Close
         Exit Function
      End If
      rsTabela.Close
   End If
   
   If f.Tag = "NEW" Then
      formatSQLAdi sCampoAdi, sValorAdi
      
      If sCampos <> "" Then
         sCampos = Mid(sCampos, 1, Len(sCampos) - 1)
         sValores = Mid(sValores, 1, Len(sValores) - 1)
      End If
      
      If nCodigo = 0 Then
         If bEmpresa = False Then
            rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
         Else
            rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " WHERE EmpCodigo = " & EmpresaAtiva.ID, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
         End If

'         If bEmpresa = False Then
'            If sCondicao <> "" Then
'               rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " where " & sCondicao, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
'               'MsgBox "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " where " & sCondicao ', IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
'            Else
'               rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
'            End If
'         Else
'            If sCondicao <> "" Then
'               rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " WHERE EmpCodigo = " & EmpresaAtiva.ID & " and " & sCondicao, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
'            Else
'               rsTabela.Open "Select MAX(" & sCampoSeq & ") as N from " & sTabela & " WHERE EmpCodigo = " & EmpresaAtiva.ID, IIf(Conexao1 Is Nothing, Conexao, Conexao1), adOpenStatic, adLockReadOnly
'            End If
'         End If
         
         If rsTabela.RecordCount = 0 Then
            nCodigo = 1
         Else
            nCodigo = CDbl(0 & rsTabela!n) + 1
         End If
         rsTabela.Close
      End If
      Dim v As String
      Call DbService.ExecuteSql("INSERT INTO " & sTabela _
      & "(" & sCampoSeq & IIf(sCampos = "", "", ",") & sCampos & sCampoAdi & ") VALUES " _
      & "(" & nCodigo & IIf(sCampos = "", "", ",") & sValores & sValorAdi & ")", , _
      IIf(Conexao1 Is Nothing, Conexao, Conexao1))
'     If DbService.Error Then MsgBox "Erro ao inserir registro!", vbCritical, "Atenção!"
   Else
      If sAltera <> "" Then   'quando existe apenas 1 campo e este é chave
         sAltera = Mid(sAltera, 1, Len(sAltera) - 1)
         'MsgBox sAltera
         If sStringAux <> "" Then sStringAux = "," & sStringAux
         If bEmpresa = False Then
            If sCondicao <> "" Then
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE " & sCondicao & " and " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            Else
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            End If
         Else
            If sCondicao <> "" Then
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE EmpCodigo=" & EmpresaAtiva.ID & " And " & sCondicao & " and " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            Else
               Call DbService.ExecuteSql("UPDATE " & sTabela & " SET " & sAltera & sStringAux & " WHERE EmpCodigo=" & EmpresaAtiva.ID & " And " & sCampoSeq & " = " & nCodigo, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
            End If
         End If
         If DbService.Error Then MsgBox "Erro ao alterar registro!", vbCritical, "Atenção!"
      End If
   End If
   
   ID = nCodigo
   
   Grava = Not DbService.Error
   
   Call Util.logSistema(Usuario.nome, sTabela, nCodigo, Conexao)
   
   Exit Function
   
erro:
   MsgBox Err.Description, vbCritical, "Atenção!"
   Call WriteFile("c:\log.log", Date & " " & Time & " " & Err.Description, False, True)
   'Resume
   Grava = False
End Function

Public Function Exclui(ID, Optional bPergunta As Boolean = False, Optional Conexao1 As ADODB.Connection = Nothing) As Boolean
    On Error GoTo erro

    If Trim$(CStr(ID)) = "" Then
        MsgBox "Nenhum registro foi selecionado!", vbExclamation, "Exclusão não realizada!"
        Exit Function
    End If
    
    If Not bPergunta Then
        If 7 = MsgBox("Confirma exclusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!!") Then Exit Function
    End If
    
    If bEmpresa = True Then
        'Conexao.Execute "DELETE FROM " & sTabela & " Where EmpCodigo = " & EmpresaAtiva.id & " AND " & sCampoSeq & " = " & id
        Call DbService.ExecuteSql("DELETE FROM " & sTabela & " Where EmpCodigo = " & EmpresaAtiva.ID & " AND " & sCampoSeq & " = " & ID, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
    Else
        'Conexao.Execute "DELETE FROM " & sTabela & " Where " & sCampoSeq & " = " & id
        Call DbService.ExecuteSql("DELETE FROM " & sTabela & " Where " & sCampoSeq & " = " & ID, , IIf(Conexao1 Is Nothing, Conexao, Conexao1))
    End If
    
    DbService.ExecuteSql ("Delete From GradeEstrutura Where GraCodigo = " & ID)
    
    Exclui = Not DbService.Error
    
    'Exclui = True
    Exit Function
               
erro:
   Call trataErros(Err.Number)
End Function

Private Sub Class_Initialize()
   Set oConexao = New cConecta
   Set Conexao = CreateObject("ADODB.Connection")
   Set rsTabela = New ADODB.Recordset
      
   Set Conexao = oConexao.Conexao
End Sub

Private Sub Class_Terminate()
   Conexao.Close
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
   
   If Right(Controles.Tag, 1) = ";" Then
      MsgBox "Campo '" & Controles.Name & "' possui separador ; sem o devido argumento!", vbCritical, "Atenção!"
      Exit Function
   End If
   vFieldMid = 0
   vCampo = Controles.Tag
'   Private Function getFieldReq(ByVal vCampo As String, ByRef vReq As Boolean) As String
      ' retira o parametro de campo requerido da Tag
   If InStr(vCampo, ";") > 0 Then
      'vCampo = Mid(vCampo, 1, InStr(vCampo, ";") - 1)
      vCampo = Util.getFieldCSV(Controles.Tag, ";", 0)
      'vFieldReq = IIf(Mid(Controles.Tag, InStr(Controles.Tag, ";") + 1, Len(Controles.Tag) - InStr(Controles.Tag, ";")) = "True", True, False)
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
   'If InStr(strNmCampo, ";") = 0 Then
      formataCampo = Mid(Util.getFieldCSV(strNmCampo, ";", 0), 5, Len(Util.getFieldCSV(strNmCampo, ";", 0)))
   'Else
   '   formataCampo = Mid(strNmCampo, 5, InStr(strNmCampo, ";") - 5)
   'End If
End Function

Private Function Consulta(ID As Double)
   On Error Resume Next
   
   Dim vFieldReq As Boolean, vFieldDes As String, vFieldMid As Integer
   
   If bEmpresa = True Then
      If sCondicao <> "" Then
         rsTabela.Open "Select * from " & sTabela & " Where " & sCondicao & " And " & sCampoSeq & " = " & ID & " AND EmpCodigo = " & EmpresaAtiva.ID, Conexao, adOpenDynamic, adLockOptimistic
      Else
         rsTabela.Open "Select * from " & sTabela & " Where " & sCampoSeq & " = " & ID & " AND EmpCodigo = " & EmpresaAtiva.ID, Conexao, adOpenDynamic, adLockOptimistic
      End If
   Else
      If sCondicao <> "" Then
         rsTabela.Open "Select * from " & sTabela & " Where " & sCondicao & " And " & sCampoSeq & " = " & ID, Conexao, adOpenDynamic, adLockOptimistic
      Else
         rsTabela.Open "Select * from " & sTabela & " Where " & sCampoSeq & " = " & ID, Conexao, adOpenDynamic, adLockOptimistic
      End If
   End If
   
   Dim Controles As Control
   Dim i As Integer
   Dim strCampo As String
   For Each Controles In f.Controls
       'MsgBox Controles.Name
       If TypeOf Controles Is TextBox Or TypeOf Controles Is DataCombo Or TypeOf Controles Is ActiveText Or _
          TypeOf Controles Is CheckBox Or TypeOf Controles Is ComboBox Or TypeOf Controles Is Label Or _
          TypeOf Controles Is DTPicker Or TypeOf Controles Is OptionButton Or TypeOf Controles Is TwPanel Then
          If Mid(UCase(Controles.Tag), 1, 3) = "FLD" Then
             For i = 0 To rsTabela.Fields.Count - 1
                 strCampo = formataCampo(Controles, vFieldReq, vFieldDes, vFieldMid)
                 If UCase(strCampo) = UCase(rsTabela(i).Name) Then
                    If TypeOf Controles Is DataCombo Then         'datacombo
                       If IsNumeric(rsTabela(i).Value) Then
                           Controles.BoundText = CLng(0 & rsTabela(i).Value)
                       Else
                           Controles.BoundText = "" & rsTabela(i).Value
                       End If
                    ElseIf TypeOf Controles Is CheckBox Then      'checkbox
                       Controles.Value = CSng(0 & rsTabela(i).Value)
                    ElseIf TypeOf Controles Is OptionButton Then  'option
                       Controles.Value = IIf(CSng(0 & rsTabela(i).Value) = False, 0, 1)
                    ElseIf TypeOf Controles Is Label Then         'label
                       If rsTabela(i).Type = adCurrency Then
                           Controles.Caption = "" & ftValor(rsTabela(i).Value, 2)
                       Else
                           Controles.Caption = "" & rsTabela(i).Value
                       End If
                    ElseIf TypeOf Controles Is ComboBox Then      'combobox
                        If InStr(5, Controles.Tag, "_") > 0 Then
                           Controles.ListIndex = CByte(0 & rsTabela(i).Value)
                        Else
                           If vFieldMid > 0 Then
                              'Controles.ListIndex = Util.SetaCombo(Controles, clng(0 & rsTabela(i).value))
                              Controles.ListIndex = Util.SetaCombo(Controles, "" & rsTabela(i).Value)
                           Else
                              If Not (IsNull(rsTabela(i).Value) Or rsTabela(i).Value = "") Then
                                 Controles.Text = rsTabela(i).Value
                              Else
                                 Controles.Text = ""
                              End If
                           End If
                        End If
                       'If InStr(5, Controles.Tag, "_") > 0 Then
                       '   Controles.ListIndex = CByte(0 & rsTabela(i).value)
                       'Else
                       '   If Not (IsNull(rsTabela(i).value) Or rsTabela(i).value = "") Then
                       '      Controles.text = rsTabela(i).value
                       '   Else
                       '      Controles.text = ""
                       '   End If
                       'End If
                    ElseIf TypeOf Controles Is DTPicker Then
                       If IsNull(rsTabela(i)) Then
                          Controles = 0
                       Else
                          Controles = rsTabela(i)
                       End If
                    ElseIf TypeOf Controles Is TwPanel Then
                       If rsTabela(i).Type = adCurrency Or rsTabela(i).Type = adNumeric Then
                          Controles.Caption = ftValor(IIf(IsNull(rsTabela(i).Value), 0, rsTabela(i).Value), 2)
                       ElseIf rsTabela(i).Type = adDate Or rsTabela(i).Type = adDBDate Or rsTabela(i).Type = adDBTimeStamp Then
                           Controles.Caption = Format(rsTabela(i).Value, "dd/mm/yyyy")
                       Else
                          Controles.Caption = rsTabela(i).Value
                       End If
                    Else
                       If Not (IsNull(rsTabela(i).Value) Or rsTabela(i).Value = "") Then
                          Controles = Trim(rsTabela(i).Value)
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
   
   If bEmpresa = True Then
      sSql = sSql & " AND EmpCodigo = " & EmpresaAtiva.ID
   End If
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
   
   If bEmpresa = True Then
      sSql = sSql & " AND EmpCodigo = " & EmpresaAtiva.ID
   End If
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
   
   If bEmpresa = True Then
      sSql = sSql & " AND EmpCodigo = " & EmpresaAtiva.ID
   End If
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
   
   If bEmpresa = True Then
      sSql = sSql & " WHERE EmpCodigo = " & EmpresaAtiva.ID
   End If
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
  'Resume
End Function

Public Function Seta(ID As String) As Boolean
   'On Error GoTo erro
   
   ID = CDbl(0 & Trim(ID))
   If ID = 0 Then Exit Function
   
   Dim sSql As String
   sSql = "SELECT " & sCampoSeq & " AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " = " & ID
   
   If bEmpresa = True Then
      sSql = sSql & " and EmpCodigo = " & EmpresaAtiva.ID
   End If
   If sCondicao <> "" Then
      'If bEmpresa = True Then
         sSql = sSql & " AND " & sCondicao
      'Else
      '   sSql = sSql & " WHERE " & sCondicao
      'End If
   End If
   
   With rsTabela
        'If bEmpresa = True Then
        '   .Open "SELECT " & sCampoSeq & " AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " = " & id & " And EmpCodigo = " & EmpresaAtiva.id, Conexao, adOpenDynamic, adLockReadOnly
        'Else
        '   .Open "SELECT " & sCampoSeq & " AS Codigo FROM " & Tabela & " WHERE " & sCampoSeq & " = " & id, Conexao, adOpenDynamic, adLockReadOnly
        'End If
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
   'Resume
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
       ' MsgBox Controles.Name
       If TypeOf Controles Is TextBox Or TypeOf Controles Is DataCombo Or _
          TypeOf Controles Is ActiveText Or TypeOf Controles Is CheckBox Or _
          TypeOf Controles Is ComboBox Or TypeOf Controles Is Label Or _
          TypeOf Controles Is DTPicker Or TypeOf Controles Is OptionButton Or _
          TypeOf Controles Is TwPanel Then
          If TypeOf Controles Is TwPanel Then
              sConteudo = Replace(Controles.Caption, "'", "")
              'Controles.Caption = Replace(Controles.Caption, "'", "")
          Else
              sConteudo = Replace(Controles, "'", "")
              'Controles = Replace(Controles, "'", "")
          End If
          If Mid(UCase(Controles.Tag), 1, 3) = "FLD" Then
             For i = 0 To rsTabela.Fields.Count - 1
                 If UCase(sCampoSeq) <> UCase(Mid(Controles.Tag, 5, Len(Controles.Tag))) Then
                    vFieldReq = False
                    vFieldDes = ""
                    strCampo = formataCampo(Controles, vFieldReq, vFieldDes, vFieldMid)
                    'If strCampo = "" Then MsgBox "campobranco"
                    If UCase(strCampo) = UCase(rsTabela(i).Name) Then
                        If TypeOf Controles Is ActiveText Or TypeOf Controles Is TextBox Then
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
                              sUpdate = sUpdate & Controles.Value & ","
                              sValue = sValue & Controles.Value & ","
                           ElseIf TypeOf Controles Is OptionButton Then
                              sUpdate = sUpdate & IIf(Controles.Value = False, 0, 1) & ","
                              sValue = sValue & IIf(Controles.Value = False, 0, 1) & ","
                           'ElseIf TypeOf Controles Is Label Then
                            '  sUpdate = sUpdate & IIf(Controles.Caption = "", Null, IIf(rsTabela(i).Type = adVarChar Or rsTabela(i).Type = adChar, "'" & Controles & "'", Controles.Caption)) & ","
                            '  sValue = sValue & IIf(Controles.Caption = "", Null, IIf(rsTabela(i).Type = adVarChar Or rsTabela(i).Type = adChar, "'" & Controles & "'", Controles.Caption)) & ","
                           ElseIf TypeOf Controles Is ComboBox Then
                              If vFieldMid > 0 Then
                                 sUpdate = sUpdate & "'" & Mid(Controles.Text, 1, vFieldMid) & "',"
                                 sValue = sValue & "'" & Mid(Controles.Text, 1, vFieldMid) & "',"
                              Else
                                 If Mid(Controles.Tag, Len(Controles.Tag), 1) = "_" Then
                                    sUpdate = sUpdate & Controles.ListIndex & ","
                                    sValue = sValue & Controles.ListIndex & ","
                                 Else
                                    sUpdate = sUpdate & IIf(Controles.Text = "", Null, IIf(rsTabela(i).Type = adChar Or rsTabela(i).Type = adVarChar Or rsTabela(i).Type = 202, "'" & Controles & "'", Controles.Text)) & ","
                                    sValue = sValue & IIf(Controles.Text = "", Null, IIf(rsTabela(i).Type = adChar Or rsTabela(i).Type = adVarChar Or rsTabela(i).Type = 202, "'" & Controles & "'", Controles.Text)) & ","
                                 End If
                              End If
                           ElseIf TypeOf Controles Is DTPicker Then
                              sUpdate = sUpdate & ftDat(Controles) & ","
                              sValue = sValue & ftDat(Controles) & ","
                           ElseIf TypeOf Controles Is TwPanel Then
                              If rsTabela(i).Type = adCurrency Or rsTabela(i).Type = adNumeric Then
                                 sUpdate = sUpdate & ftValor(Controles.Caption, 2, 1) & ","
                                 sValue = sValue & ftValor(Controles.Caption, 2, 1) & ","
                              ElseIf rsTabela(i).Type = adDate Or rsTabela(i).Type = adDBDate Or rsTabela(i).Type = adDBTimeStamp Then
                                  sUpdate = sUpdate & ftDat(Controles.Caption) & ","
                                  sValue = sValue & ftDat(Controles.Caption) & ","
                              Else
                                 sUpdate = sUpdate & IIf(Controles.Caption = "", Null, IIf(rsTabela(i).Type = adVarChar Or rsTabela(i).Type = adChar, "'" & Controles.Caption & "'", Controles.Caption)) & ","
                                 sValue = sValue & IIf(Controles.Caption = "", Null, IIf(rsTabela(i).Type = adVarChar Or rsTabela(i).Type = adChar, "'" & Controles.Caption & "'", Controles.Caption)) & ","
                              End If
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
                              'sUpdate = sUpdate & IIf(Controles.Text = "", Null, IIf(campo(i).Tipo = 1, "'" & Controles & "'", Controles.Text)) & ","
                           End If
                        End If
                    End If
                 End If
             Next
          End If
       End If
   Next
   If f.Tag = "EDIT" Then retCampos = sUpdate
   rsTabela.Close
   
   Exit Function
   
erro:
   MsgBox Err.Description
'   rsTabela.Close
'   Resume
End Function


Public Function Achou(sFiltro As String, Optional sTabelaDif As String) As Boolean
   Dim vSql As String
   
   If sTabelaDif = "" Then
      sTabelaDif = sTabela
   End If
      
   If bEmpresa = True Then
      vSql = " AND EmpCodigo = " & EmpresaAtiva.ID
   Else
      vSql = ""
   End If
   rsTabela.Open "Select * from " & sTabelaDif & " Where " & sFiltro & vSql, Conexao, adOpenStatic, adLockReadOnly
         
   If rsTabela.RecordCount > 0 Then Achou = True
   rsTabela.Close
End Function


