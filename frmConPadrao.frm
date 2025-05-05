VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConPadrao 
   Caption         =   "Consulta"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCon 
      Height          =   345
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   4875
   End
   Begin MSDataGridLib.DataGrid grdPad 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   10504
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCol 
      Left            =   720
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":08B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":0D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":1180
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":15D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":1A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":1E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":22D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMon 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":2B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":2FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":357C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":3898
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":3CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":4140
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":4594
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConPadrao.frx":49E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   9120
      TabIndex        =   5
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblCon 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   195
      Left            =   1380
      TabIndex        =   4
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de registros"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   7140
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblCol 
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblTam 
      Height          =   195
      Left            =   1260
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "frmConPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vSql As String
Public Chave As String
Public Edicao As Boolean
Public sCondicao As String
Public Titulo As String               ' titulo da janela
Public UsaColunaSistema As Boolean    ' usa os tamanhos especificados no vetor de colunas
Public largura As Integer             ' Largura do form
Public CampoOrdenacao As String       ' Index do Campo para ordenar

Private sTabela As String

Private oConexao As cConecta
Private Conexao As ADODB.Connection
Private rsTabela As ADODB.Recordset

Private Sub formataGrid()
   Dim i As Integer
   
   If Not UsaColunaSistema Then
      ReDim CPadraoColuna(rsTabela.Fields.Count - 1) As Long
   Else
      With grdPad
         For i = 0 To rsTabela.Fields.Count - 1
            .Columns(i).Width = CPadraoColuna(i)
         Next i
      End With
   End If
   
   With grdPad
        For i = 0 To rsTabela.Fields.Count - 1
            If InStr(1, rsTabela(i).Name, ".") > 0 Then
               .Columns(i).Width = 0
            Else
               Select Case rsTabela(i).Type
                      Case 5                ' FLOAT
                           .Columns(i).Alignment = dbgRight
                           .Columns(i).Width = IIf(Not UsaColunaSistema, 1100, .Columns(i).Width)
                      Case 19, 3              '  inteiro
                           .Columns(i).Alignment = dbgRight
                           .Columns(i).Width = IIf(Not UsaColunaSistema, 700, .Columns(i).Width)
                      Case 131, 6           ' decimal
                           .Columns(i).NumberFormat = "STANDARD"
                           .Columns(i).Alignment = dbgRight
                           .Columns(i).Width = IIf(Not UsaColunaSistema, 1000, .Columns(i).Width)
                      Case 133, 135             ' data
                           .Columns(i).NumberFormat = FormatoData
                           .Columns(i).Width = IIf(Not UsaColunaSistema, 1300, .Columns(i).Width)
                           .Columns(i).Alignment = dbgCenter
                      Case 129, 200, 202        ' 200=varchar,129=char
                           .Columns(i).Width = IIf(Not UsaColunaSistema, rsTabela(i).DefinedSize * 70, .Columns(i).Width)
                           '.Columns(i).Width = rsTabela(i).DefinedSize * 70
                           .Columns(i).Alignment = dbgLeft
                      Case Else
                           '.Columns(i).Width = IIf(Not UsaColunaSistema, rsTabela(i).DefinedSize * 70, .Columns(i).Width)
                           .Columns(i).Alignment = dbgLeft
               End Select
            End If
        Next
        
        Select Case UCase(sTabela)
            Case "COMPRA"
               For i = 0 To 1
                  grdPad.Columns(i).Locked = True
               Next
               For i = 7 To 8
                  grdPad.Columns(i).Locked = True
               Next
            Case ""
        End Select
        lblTot.Caption = rsTabela.RecordCount
   End With
   
End Sub

Private Sub txtCon_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then grdPad.SetFocus
End Sub


Private Sub Form_Activate()
   'txtCon.SetFocus
   rsTabela.Requery
   Call formataGrid
   If CampoOrdenacao = "" Then
      Call grdPad_HeadClick(1)
   Else
      Call grdPad_HeadClick(CampoOrdenacao)
   End If
         
   lblCon.Caption = StrConv((grdPad.Columns(IIf(CampoOrdenacao = "", 1, CampoOrdenacao)).Caption), vbProperCase)
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If vbKeyF2 = KeyCode Then txtCon.SetFocus
   If KeyCode = 27 Then Call DescarregaForm(Me)
   If Shift = 1 And txtCon.text <> "*" Then
      If KeyCode < 65 Then
         Call grdPad_HeadClick(KeyCode - 49)
      Else
         Call grdPad_HeadClick(KeyCode - 56)
      End If
      If KeyCode <> 106 Then txtCon.text = ""
   End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   If largura > 0 Then frmConPadrao.Width = largura
   Call CenterForm(Me)
   
   Set oConexao = New cConecta
   Set Conexao = CreateObject("ADODB.Connection")
   Set rsTabela = New ADODB.Recordset
   Set Conexao = oConexao.Conexao
   If Chave = "" Then Chave = "Codigo"
   sTabela = vSql
   
   rsTabela.Open vSql, Conexao, adOpenDynamic, IIf(Edicao = True, adLockOptimistic, adLockReadOnly)
   If Edicao = False Then
      grdPad.MarqueeStyle = dbgHighlightRow
   Else
      grdPad.MarqueeStyle = dbgFloatingEditor
   End If
   
   Set grdPad.DataSource = rsTabela.DataSource
   
   Tag = 1
End Sub

Private Sub txtCon_KeyPress(KeyAscii As Integer)
     On Error GoTo Trata

     If KeyAscii = 13 Then
         Dim campoFiltro As String
         Dim condicao As String
         Dim novoSQL As String
         Dim i As Integer

         If rsTabela.State = 1 Then rsTabela.Close

         If Trim(txtCon.text) = "" Then
             novoSQL = vSql
         Else
             For i = 0 To rsTabela.Fields.Count - 1
                 If UCase(rsTabela(i).Name) = UCase(lblCon.Caption) Then
                     campoFiltro = rsTabela(i).Name
                     Select Case rsTabela(i).Type
                       Case 3, 131, 20    ' Inteiros e decimais
                          condicao = campoFiltro & " = " & Replace(txtCon.text, ",", ".")
                       Case 5, 6, 17, 19  ' Números de ponto flutuante
                          condicao = campoFiltro & " = " & txtCon.text
                       Case 135          ' Data
                          condicao = campoFiltro & " = '" & Format(CDate(txtCon.text), FormatoData) & "'"
                       Case Else         ' Textos
                          condicao = campoFiltro & " LIKE '" & Replace(txtCon.text, "'", "''") & "%'"
                     End Select
                     Exit For
                End If
             Next
             novoSQL = vSql & " WHERE " & condicao
         End If
         If rsTabela.State = 1 Then rsTabela.Close
         rsTabela.Open novoSQL, Conexao, adOpenStatic, adLockReadOnly
         Set grdPad.DataSource = rsTabela.DataSource
         lblTot.Caption = rsTabela.RecordCount
     End If
    
     Exit Sub

Trata:
    trataErros Err.Number
End Sub

Private Sub Form_Resize()
   grdPad.Width = Me.Width - 285
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set grdPad.DataSource = Nothing
   Set rsTabela = Nothing
   Set Conexao = Nothing
   UsaColunaSistema = False
   Titulo = ""
   largura = 0
   CampoOrdenacao = ""
   sCondicao = ""
   Call DescarregaForm(Me)
End Sub

Private Sub grdPad_AfterColUpdate(ByVal ColIndex As Integer)
   On Error GoTo erro
   
   Dim n
   
   n = rsTabela.Bookmark
   rsTabela.MoveNext
   rsTabela.Bookmark = n
   
   Exit Sub
   
erro:
   MsgBox "Erro Ocorrido: " & Err.Description
   
End Sub

Private Sub grdPad_DblClick()
   On Error GoTo erro
   
   If rsTabela.RecordCount = 0 Then Exit Sub
   vRetConsulta = rsTabela.Fields(Chave)
   Tag = 0
   Call DescarregaForm(Me)
   
   Exit Sub
   
erro:
   
   MsgBox "Erro Ocorrido: " & Err.Description
   
End Sub

Private Sub grdPad_HeadClick(ByVal ColIndex As Integer)
   If ColIndex < 0 Or ColIndex > grdPad.Columns.Count - 1 Then Exit Sub

   If UCase(grdPad.Columns(ColIndex).DataField) <> "ID" Then
      lblCon.Caption = grdPad.Columns(ColIndex).Caption
   End If

End Sub

Private Sub grdPad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
          Case 13
               If grdPad.MarqueeStyle <> dbgFloatingEditor Then
                  Call grdPad_DblClick
               Else
                  'n = grdPad.Bookmark
                  'grdPad.Col = n - 1
                  'grdPad.Col = n
               End If
          Case 39
               KeyAscii = 0
          Case 46
               KeyAscii = 44
          
   End Select
End Sub

Private Sub grdPad_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
          Case vbKeyF2
               txtCon.SetFocus
   End Select
End Sub

Private Sub grdPad_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If LastCol = -1 Then LastCol = 0
   If grdPad.Col = -1 Then grdPad.Col = 0
   If grdPad.Columns(grdPad.Col).Locked = True Then
      If LastCol < grdPad.Col Then
         grdPad.Col = LastCol + 2
      Else
         grdPad.Col = LastCol
      End If
   End If
End Sub

Private Sub tmrRequery_Timer()
   rsTabela.Requery
   Call formataGrid
End Sub

Private Sub TwTitleBar_BarClose()
    Tag = 0
    Me.Width = 11760
    Titulo = ""
    largura = 0
    Call DescarregaForm(Me)
End Sub



