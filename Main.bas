Attribute VB_Name = "modMain"
Public vRetConsulta As String
Public vEmpresa As String
Global CPadraoColuna() As Long             ' vetor com tamanho das colunas do grid (Consulta Padr�o)
Public FormatoData As String
Public DbService As cDBService

Sub Main()
    vEmpresa = "Adriano Cobuccio"
    FormatoData = "dd/mm/yyyy"
    Set DbService = New cDBService
    frmCard.Show 1
End Sub
