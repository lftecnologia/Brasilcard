VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Cartões"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7920
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1860
      TabIndex        =   3
      Tag             =   "FLD_Data_Transacao"
      Top             =   2100
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   85524481
      CurrentDate     =   45782
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      ItemData        =   "frmCadCard.frx":0000
      Left            =   1860
      List            =   "frmCadCard.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "FLD_Status_Transacao;True;O campo Status é obrigatório"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtDescricao 
      Height          =   915
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "FLD_Descricao"
      Top             =   2580
      Width           =   5475
   End
   Begin VB.TextBox txtValor 
      Height          =   375
      Left            =   1860
      TabIndex        =   2
      Tag             =   "FLD_Valor_Transacao;True;O campo valor da transação é obrigatório!"
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox txtCartao 
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Tag             =   "FLD_Numero_Cartao;True;O campo cartão é obrigatório"
      Top             =   1140
      Width           =   2235
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1860
      TabIndex        =   0
      Tag             =   "FLD_Id_Transacao"
      Top             =   660
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7200
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":0030
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":034C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":0668
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":0984
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":0CA0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":0FBC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":1410
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":1864
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":1CB8
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":210C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":2560
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":287C
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCard.frx":2B98
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo Pedido"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir Pedido"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Gravar Pedido"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta Pedidos"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprime Pedido Atual"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fechar"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "IMG7"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   1140
      TabIndex        =   10
      Top             =   3660
      Width           =   450
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   2700
      Width           =   720
   End
   Begin VB.Label lblValor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label lblCartao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número Cartão"
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      Height          =   195
      Left            =   1620
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oCadPad As cCadPad

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
   
   ' Criar as tabelas se não existirem
   CriarTabelaTransacoes
   CriarTabelaLogSis
   
   Set oCadPad = New cCadPad

   oCadPad.formulario = Me
   oCadPad.Tabela = "Transacoes"
   oCadPad.Chave = "ID_Transacao"

   Call Novo
   
End Sub

Private Sub Novo()
   Call LimpaTexto(Me)
   txtID.Enabled = True
   Tag = "NEW"
   If txtID.Visible = True Then txtID.SetFocus
End Sub

Private Sub Consulta()
   Tag = "EDIT"
   txtID.Enabled = False
   txtCartao.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Avisa(Me) = True Then
      Cancel = True
      Exit Sub
   End If
    
   If formAberto("frmConPadrao") = True Then
      Unload frmConPadrao
   End If
      
   Call DescarregaForm(Me)
End Sub

Sub CreateBar()
    vcbPadrao.ToolbarImageList = ilsIcons
    vcbPadrao.MenuImageList = ilsIcons
    vcbPadrao.Toolbar = Principal.cmdBar(0).CommandBars("DEFAULT_TOOLBAR")
End Sub

Private Sub TwTitleBar_BarClose()
    If formAberto("frmConPadrao") = True Then
       Unload frmConPadrao
    End If
    Call DescarregaForm(Me)
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
   'On Error GoTo Trata
   
   Select Case Button.Index
    Case 1 'novo
         Call Novo
               
    Case 2 'exclui
         If oCadPad.Exclui(txtID.text) = True Then
            Call Novo
         End If
    
    Case 3 'grava
         oCadPad.codigo = CDbl(0 & txtID.text)
         
         If CZ(oCadPad.codigo) > 0 And cmbStatus.text = "Aprovada" Then
            MsgBox "Transações com status de 'Aprovada' não podem ser editadas!", vbInformation, "Atenção!"
         ElseIf cmbStatus.text <> "Aprovada" And cmbStatus.text <> "Cancelada" And cmbStatus.text <> "Pendente" Then
            MsgBox "Selecione um status válido, 'Aprovada', 'Pendente' ou 'Cancelada'", vbInformation, "Atenção!"
         Else
            If oCadPad.Grava = False Then Exit Sub
            Call Novo
         End If
    
    Case 4 'consulta
         frmConPadrao.Chave = "ID_Transacao"
         frmConPadrao.largura = 9500
         frmConPadrao.UsaColunaSistema = True
         frmConPadrao.CampoOrdenacao = 1
         ReDim CPadraoColuna(5) As Long
         CPadraoColuna(0) = 800
         CPadraoColuna(1) = 2000
         CPadraoColuna(2) = 1300
         CPadraoColuna(3) = 1300
         CPadraoColuna(4) = 1500
    
         vRetConsulta = 0
         ' Uso de top 400 para que não congele na carga do form em alto volume de dados
         ' internamente no form de consulta trabalhará com top 400 nos filtros
         frmConPadrao.vSql = "Select Top 400 ID_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Status_Transacao From Transacoes"
         frmConPadrao.Edicao = True
         frmConPadrao.Show 1
         If vRetConsulta = 0 Then Exit Sub

         If oCadPad.Seta(CStr(vRetConsulta)) = True Then
            Call Consulta
         End If
    
    Case 5 'relatório
 '        Crp1.Destination = crptToWindow
 '        Call ChamaRel("Card", "")
    Case 8, 9, 10, 11
        Select Case Button.Index
        Case 8 'primeiro
            oCadPad.Primeiro
        Case 9 'anterior
            oCadPad.Anterior CDbl(0 & txtID.text)
        Case 10 'próximo
            oCadPad.Proximo CDbl(0 & txtID.text)
        Case 11 'ultimo
            oCadPad.Ultimo
        End Select
        Call Consulta
    End Select
    
    Exit Sub
    
Trata:
    Call trataErros(Err.Number)
End Sub

Private Sub txtCartao_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
