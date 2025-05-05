Attribute VB_Name = "DataBase"
Public Sub CriarTabelaTransacoes()
    Dim strSQL As String
    Dim Conexao As ADODB.Connection
    Dim objCon As New cConecta
    
    On Error GoTo TrataErro

    ' Montar o SQL
    strSQL = ""
    strSQL = strSQL & "IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Transacoes' AND xtype='U') BEGIN "
    strSQL = strSQL & "CREATE TABLE Transacoes ("
    strSQL = strSQL & "Id_Transacao INT IDENTITY(1,1) PRIMARY KEY, "
    strSQL = strSQL & "Numero_Cartao CHAR(16) NOT NULL, "
    strSQL = strSQL & "Valor_Transacao DECIMAL(18,2) NOT NULL CHECK (Valor_Transacao > 0), "
    strSQL = strSQL & "Data_Transacao DATETIME NOT NULL DEFAULT GETDATE(), "
    strSQL = strSQL & "Descricao VARCHAR(255), "
    strSQL = strSQL & "Status_Transacao VARCHAR(10) NOT NULL CHECK (Status_Transacao IN ('Aprovada', 'Pendente', 'Cancelada'))"
    strSQL = strSQL & ") END"

    ' Pegar a conexão da classe
    Set Conexao = objCon.Conexao()

    ' Executar o SQL
    Conexao.Execute strSQL
    Conexao.Close
    Set Conexao = Nothing
    
    Exit Sub

TrataErro:
    MsgBox "Erro ao criar tabela: " & Err.Description, vbCritical, "Erro"
End Sub

Public Sub CriarTabelaLogSis()
    Dim Conexao As ADODB.Connection
    Dim objCon As New cConecta
    Dim strSQL As String
    
    On Error GoTo TrataErro

    strSQL = ""
    strSQL = strSQL & "IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='logsis' AND xtype='U') BEGIN "
    strSQL = strSQL & "CREATE TABLE logsis ("
    strSQL = strSQL & "ID INT PRIMARY KEY, "
    strSQL = strSQL & "LogData DATE NOT NULL, "
    strSQL = strSQL & "LogHora CHAR(5) NOT NULL, "
    strSQL = strSQL & "LogUsuario VARCHAR(50) NOT NULL, "
    strSQL = strSQL & "LogTabela VARCHAR(100) NOT NULL, "
    strSQL = strSQL & "LogTabelaID VARCHAR(50) NOT NULL, "
    strSQL = strSQL & "Empresa VARCHAR(50) NOT NULL, "
    strSQL = strSQL & "Exception VARCHAR(MAX)"
    strSQL = strSQL & ") END"

    ' Pegar a conexão da classe
    Set Conexao = objCon.Conexao()
    
    Conexao.Execute strSQL
    
    Conexao.Close
    Set Conexao = Nothing
    
    Exit Sub

TrataErro:
    MsgBox "Erro ao criar tabela logsis: " & Err.Description, vbCritical
End Sub

