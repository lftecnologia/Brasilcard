Attribute VB_Name = "Geral"
'impressora
Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Function trataErros(idErr As Double) As Boolean
   'Dim olog As New LogObject
   Dim vDescricao As String
   Dim tblErr As String
   
   'olog.AppName = App.EXEName
   'olog.FileName = "log"
      
   Select Case idErr
          Case 0
               Exit Function
          Case 3001, 3021, -2147024809, -2147217900, -2147217904, -2147352571
               MsgBox "Conteúdo informado inválido.", 16, "Erro inesperado!"
          Case -2147467259, 3709  'Acesso à rede
               If Config.DbType = PostgreSQL Then
                  If InStr(Err.Description, "no connection to the server") > 0 Or InStr(Err.Description, "server closed the connection unexpectedly") > 0 Then
                     MsgBox "Erro de rede, o sistema será fechado!", vbCritical, "Perda de Conexão de Rede"
                     End
                  Else
                     MsgBox "Erro Ocorrido: " & Err.Description, vbCritical, "Atenção!"
                  End If
               ElseIf Config.DbType = SqlServer Then
                  If Err.Number = "-2147467259" Then
                     MsgBox "Erro de rede, o sistema será fechado!", vbCritical, "Perda de Conexão de Rede"
                     End
                  End If
               End If
          Case -2147217873
               tblErr = ErrExclusao("Existe Itens repetidos para serem incluidos ou alterados!")
               MsgBox Err.Description, 16, "Erro de Inlcusão/Alteração/Exclusão"
          Case -2147217887
               MsgBox "Digite o dado no formato requerido!", 16, "Erro de Consulta"
          Case 13
               MsgBox "Tipo de dado inválido!", 16, "Atenção!!"
          Case 75
               MsgBox "Arquivo não Encontrado!", vbInformation, "Atenção!"
               
          Case Else
               MsgBox "Erro não reportado, chame o supervisor do sistema!" & Chr(13) & Err.Description, 16, "Erro Ocorrido!"
   End Select
End Function

Public Function ftValor(ByVal parValor As Variant, Optional ByVal parDecimal As Byte = 2, Optional vSqlConf As Integer, Optional ByVal vProCodigo As Long = 0) As String
   Dim cfgArredondamento As Integer
   
   ' Parâmetro vSqlConf: Se = 1 troca , por .
   ' cfgArredondamento: 0- Arredondamento pdv 1-ABNT 2-Truncamento
   
   Dim sZero As String
   Dim i As Integer
   
   If IsMissing(parDecimal) Then parDecimal = 2
   
   ' #BALANCA
   If vProCodigo > 0 Then
      If DbService.GetUniqueValue("Select ProBalanca From Produto Where ProCodigo=" & vProCodigo & _
      IIf(ConfigGeral.ConfUsaProEmpresa = 1, " And EmpCodigo=" & EmpresaAtiva.ID, "")) = "1" Then
         cfgArredondamento = CZ(DbService.GetUniqueValue("select ConfArredondamento from Configuracao"))
      End If
   Else
      cfgArredondamento = 0
   End If
   
   If Not IsNumeric(parValor) Then
      ftValor = ""
   Else
      If parDecimal = 0 Then              ' Se informou Zero nas decimais é INTEIRO
         ftValor = Int(parValor)
        ' ftValor = Format(parValor, "###,###,##0.00; -###,###,##0.00")
      Else
         For i = 1 To parDecimal
            sZero = sZero & "0"
         Next
         If vSqlConf = 1 Then
            ftValor = Format(parValor, "##0." & sZero & "; -##0." & sZero)
            ftValor = Replace(ftValor, ",", ".")
         Else
            Select Case cfgArredondamento
               Case 1 ' abnt
                     ftValor = Format(Arredondamento_ABNT_NBR5891(parValor), "########0.00")
               Case 0 ' pdv
                  ftValor = Format(parValor, "########0." & sZero & "; -########0." & sZero)
               Case 2 ' truncamento
                  ftValor = Format(Trunca(parValor, 2), "########0." & sZero & "; -########0." & sZero)
            End Select
         End If
      End If
   End If
End Function

Public Function ftDat(ByVal vData As String, Optional ByVal UTC As Boolean) As String
    ' Formata a data
    If Not UTC Then
       If vData = "" Then
          ftDat = "null"
       Else
          ftDat = "'" & Format(vData, FormatoData) & "'"
      End If
   Else
      ftDat = Format(vData, "YYYY-MM-DD") & "T" & Format(Time, "hh:mm:ss") & IIf(DbHelper.IsHorarioVerao, "-02:00", "-03:00")
   End If
End Function

Public Function ErrExclusao(NomErr As String) As String
   Dim Result(1) As Integer
   Result(0) = (InStr(1, NomErr, "'")) + 1
   Result(1) = InStr(Result(0), NomErr, "'")
   ErrExclusao = Mid(NomErr, Result(0), IIf(Result(1) - Result(0) < 0, 0, Result(1) - Result(0)))
End Function

Public Function CZ(ValorEntrada As Variant, Optional ByVal IfZero As Long = 0) As Currency
   On Error GoTo erro
   
   'Currency Zero, antigo ccurneg
   
   If IsNull(ValorEntrada) Then
      CZ = 0
   ElseIf ValorEntrada = "" Or ValorEntrada = "," Or Not IsNumeric(ValorEntrada) Then
      CZ = 0
   Else
      CZ = ValorEntrada
   End If
   
   If IfZero <> 0 And CZ = 0 Then
      CZ = IfZero
   End If
      
   Exit Function
   
erro:
   MsgBox "Erro em Função CZ() " & Err.Description, vbInformation, "Chame o suporte"
   
End Function

Public Function Arredondamento_ABNT_NBR5891(ByVal valor As Currency) As Currency
   On Error GoTo Trata_Erros
   
   'TRANSFORMA E FORMATA O VALOR PARA STRING E 4 DECIMAIS
   Dim StrValor_Trabalhar As String
   StrValor_Trabalhar = Format(valor, "############0.0000")

   'DESCOBRE A POSI??O DA VIRGULA
   Dim Posicao_Virgula As Integer
   Posicao_Virgula = InStr(1, CStr(StrValor_Trabalhar), ",")
   Dim StrDecimal As String
   StrDecimal = Mid(StrValor_Trabalhar, Posicao_Virgula + 1, Len(StrValor_Trabalhar))
  
   'VERIFICA SE NA DECIMAL OS 2 ULTIMOS DIGITOS SÃO IGUAIS A 00, SE FOREM, NÃO SERÁ NECESS?RIO ARREDONDAR
   'POR EXEMPLO 2,5500
   If Mid(StrDecimal, 3, 2) = "00" Then
      Arredondamento_ABNT_NBR5891 = Format(CCur(StrValor_Trabalhar), "############0.00")
      Exit Function
   End If
  
   ' Default
   Dim StrValor_Retornar As String
   StrValor_Retornar = CStr(Format(valor, "#############0.00"))
  
   '********************************************************************************************************************************************
   '1- Quando o algarismo seguinte a 2a. CASA for INFERIOR a 5, A 2a. CASA permanecer? SEM modificação
   'ENTAO SE NA 3a. CASA O NUMERO FOR < 5 (MENOR QUE 5) ENTAO NAO ARREDONDA, MANTEM O VALOR ORIGINAL
   'EXEMPLO 2,5501 FICARA SOMENTE 2,55 POIS A TERCEIRA CASA (0) ? MENOR QUE 5
   '********************************************************************************************************************************************
   If CInt(Mid(StrDecimal, 3, 1)) < 5 Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5501  REMOVER? O 01 DO FINAL, RETORNANDO SOMENTE O 2,55
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
      
   End If
  
   '********************************************************************************************************************************************
   '2 - Quando o algarismo seguinte A 2? CASA for SUPERIOR a 5 ENTAO AUMENTARA EM UMA UNIDADE A 2? CASA, EXEMPLO: 2,556 (FICA 2,56)
   '********************************************************************************************************************************************
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   'VERIFICA SE A TERCEIRA CASA ? MAIOR QUE 5
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If CInt(Mid(StrDecimal, 3, 1)) > 5 Then
      'SE FOR MAIOR QUE 5, ENT?O ARREDONDA PRA MAIS O VALOR, EXEMPLO: 2,556 FICAR? 2,56
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5501  REMOVER? O 01 DO FINAL, RETORNANDO SOMENTE O 2,55
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
  
   '************************************************************************************************************************************************************************
   '3 - Quando a TERCEIRA CASA ? IGUAL A CINCO, TEREMOS 2 OPCOES (A e B):
   '************************************************************************************************************************************************************************
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '(A) - SE A SEGUNDA CASA FOR IMPAR ENT?O ARREDONDA PRA MAIS O VALOR, EXEMPLO: 2,3751 (o 7 dos 37 centavos ? IMPAR, neste caso arredonda pra mais)
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If IsImpar(CLng(Mid(StrDecimal, 2, 1))) = True Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,3751  REMOVER? O 51 DO FINAL, RETORNANDO SOMENTE O 2,37
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '(B) - SE A SEGUNDA CASA FOR PAR, ENT?O:
   'SE A QUARTA CASA FOR ALGARISMO ZERO, N?O HAVER? ALTERA??O NAS DECIMAIS, RETORNANDO O VALOR SEM ARREDONDAR, EXEMPLO: 2,5450 (FICARA 2,54)
   'SE A QUARTA CASA FOR ALGARISMO DIFERENTE DE ZERO, A 2? CASA  dever? ser AUMENTADA EM UMA unidade, EXEMPLO: 2,5451 (FICAR? 2,55)
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  
   'SE A QUARTA CASA FOR IGUAL A ZERO
   If CInt(Mid(StrDecimal, 4, 1)) = 0 Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5450  REMOVER? O 50 DO FINAL, RETORNANDO SOMENTE O 2,54
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
  
   'SE A QUARTA CASA FOR MAIOR QUE ZERO, ACRESCENTA EM 0,01 ARREDONDANDO PRA MAIS O VALOR DECIMAL COM 2 CASAS
   Else
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,3451  REMOVER? O 51 DO FINAL, RETORNANDO SOMENTE O 2,34
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")  'SOMA MAIS 1 CENTAVO
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
Trata_Erros:
   If Err.Number <> 0 Then
      MsgBox "Erro na funcao de ARREDONDAMENTO ABNT NBR 5891: " & Err.Source & Err.Description, vbCritical
      Exit Function
   End If
End Function

Public Function Trunca(ByVal valor As Currency, ByVal Decimais As Integer) As Double
   Dim nDec As Integer, i As Integer, sAux As String
   sAux = "1"
   For i = 1 To Decimais
      sAux = sAux & "0"
   Next i
   nDec = Val(sAux)
   If IsNull(valor) Then
      Trunca = 0
   Else
      Trunca = Int(valor * nDec) / nDec
   End If
End Function

Function IsImpar(ByVal iNum As Long) As Boolean
  IsImpar = (iNum Mod 2)
End Function

Sub LimpaTexto(f As Form)
    Dim Controles As Control
    For Each Controles In f.Controls
        If TypeOf Controles Is TextBox Or TypeOf Controles Is DBCombo Or TypeOf Controles Is DataCombo Or TypeOf Controles Is Label Or TypeOf Controles Is CheckBox Then
           If TypeOf Controles Is Label Then
              If Mid(Controles.Tag, 1, 3) = "FLD" Then Controles.Caption = ""
           ElseIf TypeOf Controles Is CheckBox Then
               Controles.value = 0
           ElseIf Controles.Tag <> "NO_CLEAN" Then
              Controles.text = ""
           End If
        End If
    Next Controles
End Sub

Public Function Avisa(f As Form) As Boolean
   If f.Tag = "EDIT" Then
      If 7 = MsgBox("Confirma saida do formulario!", vbYesNo + 32, "Atenção!") Then
         Avisa = True
      End If
   End If
End Function

Function formAberto(Texto As String) As Boolean
   Dim v As Integer
   For v = 0 To Forms.Count - 1
       If Forms(v).Name = Texto Then
          formAberto = True
          Exit For
       End If
   Next
End Function

Sub DescarregaForm(f As Form)
    Dim Controles As Control
    For Each Controles In f.Controls
        Set Controles = Nothing
    Next Controles
    Unload f
End Sub


Public Sub ChamaRel(Relatorio As String, Formula As String, Optional ByVal A4 As Boolean = True, _
                    Optional ByVal Terminal As String = "IsMissing")
                    
    On Error GoTo Trata
    
    Dim i As Integer
    
    With Principal.Crp1
        .WindowShowPrintSetupBtn = True
        .WindowShowGroupTree = True
        .WindowShowSearchBtn = True
        .WindowShowNavigationCtls = True
        .WindowTitle = "Relatório - Sistema de Automação Fusion"
        
        'i = .LogOnServer("pdsodbc.dll", "Fusion", "total_2", "sa", "12")
        
'        .RetrieveLogonInfo
         
        Call DefineImpressoraPadraoWindows
        
        .WindowState = crptMaximized
        
        .WindowLeft = 0
        .WindowTop = 0
        .Formulas(0) = "cred = 'Adriano Cobuccio'"
        
        If Config.DbType = SqlServer Then
           .UserName = "sa"
           .Password = "12"
           .Connect = "; Uid=sa; pwd=12"
        Else
           .UserName = "postgres"
           .Password = "12"
           .Connect = "; Uid=postgres; pwd=12"
        End If
        '.LogOnServer "Driver={SQL Server};Server=" & Config.Server & ";Database=" & Config.DatabaseName & ";Uid=sa;Pwd=12;"
        .SelectionFormula = Formula
        .ReportFileName = (Config.ReportPath & "\" & Relatorio & ".Rpt")
        '.RetrieveDataFiles
        .WindowShowCancelBtn = False
        .Action = 1
        .SelectionFormula = ""
        For i% = 0 To 45: .Formulas(i%) = "": Next i%
        .Reset
    End With
    
    Principal.Crp1.Destination = crptToWindow
    
    Exit Sub
   
Trata:
    Principal.Crp1.SelectionFormula = ""
    Principal.Crp1.Reset
    For i = 0 To 45: Principal.Crp1.Formulas(i%) = "": Next
    Select Case Err.Number
           Case 20526
                MsgBox "Não há impressora instalada!" & Chr(13) & "É necessário a instalação de uma impressora!", 16, "Atenção!!"
           Case 20504, 20507
                MsgBox "Relatório " & Relatorio & ".rpt não encontrado!" & Chr(13) & "Verifique o caminho: " & UCase(Config.ReportPath), 16, "Atenção!!"
           Case Else
                MsgBox Err.Number & " " & Err.Description, vbCritical, "Dispara Relatório: " & Relatorio
    End Select
    Principal.Crp1.Destination = crptToWindow
End Sub

Public Sub DefineImpressoraPadraoWindows()
   Dim objPrinter As Printer
   Set objPrinter = GetDefaultPrinter()
   
   With Principal.Crp1
      .PrinterDriver = objPrinter.DriverName
      .PrinterPort = objPrinter.Port
      .PrinterName = objPrinter.DeviceName
   End With
   
   Set objPrinter = Nothing
End Sub

Public Function GetDefaultPrinter() As Printer
   Dim strBuffer As String * 254
   Dim iRetValue As Long
   Dim strDefaultPrinterInfo As String
   Dim tblDefaultPrinterInfo() As String
   Dim objPrinter As Printer

' pega as informacoes da impressora padrao
  iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
  strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
  tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
  For Each objPrinter In Printers
        If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
          ' se achou a impressora padrao entao sai
          Exit For
        End If
   Next
   ' se nao achou retrona nothing
  If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then
      Set objPrinter = Nothing
  End If
  Set GetDefaultPrinter = objPrinter
End Function

Public Sub CenterForm(f As Form)
    Screen.MousePointer = 11
    If f.WindowState <> vbMaximized Then
      f.Move Screen.Width / 2 - f.Width / 2, (Screen.Height) / 2 - f.Height / 2
    End If
    Screen.MousePointer = 0
End Sub

Public Sub Sendkeys(text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys text, wait
      Set WshShell = Nothing
End Sub
