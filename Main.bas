Attribute VB_Name = "modMain"
Public vRetConsulta As String
Public vEmpresa As String
Global CPadraoColuna() As Long             ' vetor com tamanho das colunas do grid (Consulta Padrão)
Public FormatoData As String

Sub Main()
    vEmpresa = "Adriano Cobuccio"
    FormatoData = "dd/mm/yyyy"
    frmCard.Show 1
End Sub
