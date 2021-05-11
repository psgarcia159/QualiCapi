VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form FrmCartas 
   Caption         =   "Cadastro de Cartas"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7740
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgOperacoes 
      Left            =   720
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".rtf"
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   60
      TabIndex        =   8
      Top             =   4680
      Width           =   9135
   End
   Begin VB.TextBox txtDescricao 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   7935
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   60
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0AB4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0BC6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0CD8
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0DEA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":0EFE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":1442
            Key             =   "Painter"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":1986
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":1ECA
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":240E
            Key             =   "Bullets"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":29F6
            Key             =   "FSreen"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCartas.frx":30C2
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1296
      BandCount       =   2
      _CBWidth        =   9060
      _CBHeight       =   735
      _Version        =   "6.7.9782"
      Child1          =   "tbToolBar"
      MinHeight1      =   330
      Width1          =   2295
      NewRow1         =   0   'False
      MinHeight2      =   315
      Width2          =   4410
      NewRow2         =   -1  'True
      Begin VB.ComboBox cboFonteNome 
         Height          =   315
         Left            =   165
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   2790
      End
      Begin VB.ComboBox cboFonteTam 
         Height          =   315
         Left            =   3060
         TabIndex        =   3
         Top             =   390
         Width           =   855
      End
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Novo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Abrir"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Salvar"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cortar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copiar"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Colar"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Negrito"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Itálico"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Sublinhado"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Left"
               Object.ToolTipText     =   "Alinhado à Esquerda"
               ImageIndex      =   11
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Centralizado"
               ImageIndex      =   12
               Style           =   2
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Right"
               Object.ToolTipText     =   "Alinhado à Direita"
               ImageIndex      =   13
               Style           =   2
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullets"
               Object.ToolTipText     =   "Marcadores"
               ImageIndex      =   19
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtfTexto 
      Height          =   3075
      Left            =   0
      TabIndex        =   2
      Top             =   1500
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5424
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"FrmCartas.frx":36FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCorpo 
      Caption         =   "Corpo da Carta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label lblDescricao 
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   915
   End
End
Attribute VB_Name = "FrmCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constantes
Const PHYSICALOFFSETX As Long = 112
Const PHYSICALOFFSETY As Long = 113
Const WM_USER = &H400
Const EM_FORMATRANGE As Long = WM_USER + 57

'Tipos definidos pelo usuário
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type
Private Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As Rect        ' Region of the DC to draw to (in twips)
  rcPage As Rect    ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type

'Variáveis globais
Dim XLB_ALTEROU As Boolean
Dim XLI_ALTURAANT As Integer, XLI_LARGURAANT As Integer, XLI_DIFALTURA As Integer, XLI_DIFLARGURA As Integer
Dim XFT_SQL As String
Dim XFO_RSCARTA As ADODB.Recordset

'Funções Externas
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, Ip As Any) As Long

Private Sub cboFonteNome_Click()
  'Muda a fonte do texto selecionado
  rtfTexto.SelFontName = cboFonteNome.Text 'cboFonteNome.List(cboFonteNome.ListIndex)
End Sub

Private Sub cboFonteTam_Click()
'  Muda o tamanho do texto selecionado
  rtfTexto.SelFontSize = cboFonteTam.Text
End Sub

Private Sub CmdCancelar_Click()
  XGB_RESULTADO = False
  Unload Me
End Sub

Private Sub cmdGravar_Click()
  On Error GoTo Cancelou
  'Inserção de Carta
  If XInserir = "I" Then
    XFT_SQL = "SELECT * FROM Cartas WHERE cart_cd_carta = -1"
    SubQOpenRecordset XFO_RSCARTA, XFT_SQL, Dinamico
    XFO_RSCARTA.AddNew
    XFO_RSCARTA.Fields("cart_tx_Descricao") = txtDescricao.Text
    XFO_RSCARTA.Fields("cart_tx_texto") = rtfTexto.TextRTF
    XFO_RSCARTA.Update
  'ALteracao de Carta
  Else
    XFO_RSCARTA.Fields("cart_tx_Descricao") = txtDescricao.Text
    XFO_RSCARTA.Fields("cart_tx_texto").Value = rtfTexto.TextRTF
    XFO_RSCARTA.Update
  End If
  Set XFO_RSCARTA = Nothing
  
  XGB_RESULTADO = True
  Unload Me

Cancelou:
End Sub

Private Sub Form_Load()
  Dim XLT_SQL As String
  Dim x As Integer
  
  'Inicializa as variáveis de redimensionamento
  XLI_ALTURAANT = Me.Height
  XLI_LARGURAANT = Me.Width
  
  'Preenche o combo de fontes
  For x = 1 To Screen.FontCount
    cboFonteNome.AddItem Screen.Fonts(x)
  Next
  
  'Preenche o combo de tamanho das fontes
  'com os tamanhos padrões utilizados pelo word
  cboFonteTam.AddItem Str$(8)
  cboFonteTam.AddItem Str$(9)
  cboFonteTam.AddItem Str$(10)
  cboFonteTam.AddItem Str$(11)
  cboFonteTam.AddItem Str$(12)
  cboFonteTam.AddItem Str$(14)
  cboFonteTam.AddItem Str$(16)
  cboFonteTam.AddItem Str$(18)
  cboFonteTam.AddItem Str$(20)
  cboFonteTam.AddItem Str$(22)
  cboFonteTam.AddItem Str$(24)
  cboFonteTam.AddItem Str$(26)
  cboFonteTam.AddItem Str$(28)
  cboFonteTam.AddItem Str$(36)
  cboFonteTam.AddItem Str$(48)
  cboFonteTam.AddItem Str$(72)

  'Apresenta a fonte padrão do texto no combo de fontes
  For x = 0 To cboFonteNome.ListCount - 1
    If rtfTexto.SelFontName = cboFonteNome.List(x) Then
      cboFonteNome.ListIndex = x
      Exit For
    End If
  Next
  
  'Apresenta o tamanho padrão do texto no combo de tamanhos
  rtfTexto.SelFontSize = 10
  For x = 0 To cboFonteTam.ListCount - 1
    If CInt(rtfTexto.SelFontSize) = CInt(cboFonteTam.List(x)) Then
      cboFonteTam.ListIndex = x
      Exit For
    End If
  Next

  'Carrega os dados caso seja uma alteração de carta
  If XInserir = "A" Then
    XLT_SQL = "SELECT * FROM Cartas WHERE cart_cd_carta = " & Chave
    SubQOpenRecordset XFO_RSCARTA, XLT_SQL, Dinamico
    txtDescricao.Text = XFO_RSCARTA("cart_tx_Descricao").Value
    rtfTexto.TextRTF = XFO_RSCARTA("cart_tx_texto").Value
  End If
End Sub

Private Sub DefinirMenu()
  'Só habilita estes botões caso o exista algum texto selecionado
  If rtfTexto.SelLength > 0 Then
    tbToolBar.Buttons(7).Enabled = True
    tbToolBar.Buttons(8).Enabled = True
  Else
    tbToolBar.Buttons(7).Enabled = False
    tbToolBar.Buttons(8).Enabled = False
  End If
  'Só habilita estes botões caso exista dados na área de transf.
  If Clipboard.GetFormat(vbCFText) Then
    tbToolBar.Buttons(9).Enabled = True
  Else
    tbToolBar.Buttons(9).Enabled = False
  End If
End Sub

Private Sub chkBullets()
  If rtfTexto.SelBullet = True Then
    Me.tbToolBar.Buttons("Bullets").Value = tbrPressed
  Else
    Me.tbToolBar.Buttons("Bullets").Value = tbrUnpressed
  End If
End Sub

Private Sub Form_Resize()
  SubAjustaTela
End Sub

Private Sub rtfTexto_Change()
  XLB_ALTEROU = True
End Sub

Private Sub rtfTexto_SelChange()
  chkBullets
  ' Atualiza o estado dos botões da barra de ferramentas de acordo
  ' com o  estado das propriedas do texto selecionado.
  tbToolBar.Buttons("Bold").Value = IIf(rtfTexto.SelBold, tbrPressed, tbrUnpressed)
  tbToolBar.Buttons("Italic").Value = IIf(rtfTexto.SelItalic, tbrPressed, tbrUnpressed)
  tbToolBar.Buttons("Underline").Value = IIf(rtfTexto.SelUnderline, tbrPressed, tbrUnpressed)
  tbToolBar.Buttons("Align Left").Value = IIf(rtfTexto.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
  tbToolBar.Buttons("Center").Value = IIf(rtfTexto.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
  tbToolBar.Buttons("Align Right").Value = IIf(rtfTexto.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
  On Error Resume Next
  cboFonteNome.Text = rtfTexto.SelFontName
  cboFonteTam.Text = rtfTexto.SelFontSize
  DefinirMenu
End Sub

Private Sub SubAjustaTela()
  If (Me.Width >= 9180) And (Me.Height >= 5025) Then
    XLI_DIFALTURA = Me.Height - XLI_ALTURAANT
    XLI_DIFLARGURA = Me.Width - XLI_LARGURAANT
    XLI_ALTURAANT = Me.Height
    XLI_LARGURAANT = Me.Width
    
    rtfTexto.Width = Me.ScaleWidth
    rtfTexto.Height = rtfTexto.Height + XLI_DIFALTURA
    txtDescricao.Width = txtDescricao.Width + XLI_DIFLARGURA
    Frame1.Width = Frame1.Width + XLI_DIFLARGURA
    Frame1.Top = Frame1.Top + XLI_DIFALTURA
    cmdCancelar.Left = cmdCancelar.Left + XLI_DIFLARGURA
    cmdCancelar.Top = cmdCancelar.Top + XLI_DIFALTURA
    cmdGravar.Left = cmdGravar.Left + XLI_DIFLARGURA
    cmdGravar.Top = cmdGravar.Top + XLI_DIFALTURA
  Else
    On Error GoTo RotuloErro
    If Me.Width < 9180 Then
      Me.Width = 9180
    Else
      Me.Height = 5025
    End If
  End If
  
RotuloErro:

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim bBullets As Boolean
  Dim XLI_RESPOSTA As Integer
  
  On Error Resume Next
  Select Case Button.Key
    Case "New"
      'Verifica se o documento foi alterado.
      'Se foi, perguta se deseja salvar e depois cria um novo documento em branco.
      If XLB_ALTEROU Then
        XLI_RESPOSTA = MsgBox("Deseja salvar as alterações no documento?", vbYesNoCancel, _
          "Salvar as alterações")
        If XLI_RESPOSTA = vbOK And FunSalvarArquivo Then
          SubNovoDocumento
          XLB_ALTEROU = False
        ElseIf XLI_RESPOSTA = vbNo Then
          SubNovoDocumento
          XLB_ALTEROU = False
        End If
      Else
        SubNovoDocumento
      End If
    Case "Open"
      'Verifica se o documento foi alterado.
      'Se foi, perguta se deseja salvar e depois pede o documento para abrir.
      If XLB_ALTEROU Then
        XLI_RESPOSTA = MsgBox("Deseja salvar as alterações no documento?", vbYesNoCancel, _
          "Salvar as alterações")
        If XLI_RESPOSTA = vbYes And FunSalvarArquivo Then
          If FunAbrirArquivo Then XLB_ALTEROU = False
        ElseIf XLI_RESPOSTA = vbNo Then
          If FunAbrirArquivo Then XLB_ALTEROU = False
        End If
      Else
        FunAbrirArquivo
      End If
    Case "Save"
      If FunSalvarArquivo Then XLB_ALTEROU = False
    Case "Print"
      'Chama o rotina de impressão
      PrintRTF rtfTexto, 720, 720, 720, 720
    Case "Cut"
      On Error Resume Next
      Clipboard.Clear
      Clipboard.SetText rtfTexto.SelText
      rtfTexto.SelText = vbNullString
    Case "Copy"
      On Error Resume Next
      Clipboard.SetText rtfTexto.SelText
    Case "Paste"
      On Error Resume Next
      rtfTexto.SelRTF = Clipboard.GetText
    Case "Bold"
      rtfTexto.SelBold = Not rtfTexto.SelBold
      Button.Value = IIf(rtfTexto.SelBold, tbrPressed, tbrUnpressed)
    Case "Italic"
      rtfTexto.SelItalic = Not rtfTexto.SelItalic
      Button.Value = IIf(rtfTexto.SelItalic, tbrPressed, tbrUnpressed)
    Case "Underline"
      rtfTexto.SelUnderline = Not rtfTexto.SelUnderline
      Button.Value = IIf(rtfTexto.SelUnderline, tbrPressed, tbrUnpressed)
    Case "Align Left"
      rtfTexto.SelAlignment = rtfLeft
    Case "Center"
      rtfTexto.SelAlignment = rtfCenter
    Case "Align Right"
      rtfTexto.SelAlignment = rtfRight
    Case "Bullets"
      With rtfTexto
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
          'A seleção está mista (com e sem marcador) ou não possui marcador algum.
          'Adiciona o marcador.
          .SelBullet = True
          tbToolBar.Buttons("Bullets").Value = tbrPressed
        ElseIf .SelBullet = True Then
          'A seleção possui marcador então remove.
          .SelBullet = False
          .SelHangingIndent = False
          tbToolBar.Buttons("Bullets").Value = tbrUnpressed
        End If
      End With
  End Select
End Sub

Private Sub SubNovoDocumento()
  rtfTexto.Text = ""
  XLB_ALTEROU = False
End Sub

Private Function FunAbrirArquivo() As Boolean
  FunAbrirArquivo = False
  On Error GoTo Cancelou
  
  ' Tipos de arquivos a serem abertos
  dlgOperacoes.Filter = "Ritch Text Files (*.rtf)|*.rtf|Arquivos Texto (*.txt)|*.txt|All files (*.*)|*.*"
  
  'Não permite que arquivos somente para leitura sejam abertos
  dlgOperacoes.Flags = cdlOFNHideReadOnly
  dlgOperacoes.ShowOpen
  
  If dlgOperacoes.FileName <> "" Then
    If UCase(Right(dlgOperacoes.FileName, 3)) = "RTF" Then
      rtfTexto.LoadFile dlgOperacoes.FileName, rtfRTF
    Else
      rtfTexto.LoadFile dlgOperacoes.FileName, rtfText
    End If
    FunAbrirArquivo = True
  End If

Cancelou:
End Function

Private Function FunSalvarArquivo() As Boolean
  FunSalvarArquivo = False
  On Error GoTo Cancelou
  'Não permite salvar em arquivos somente para leitura e
  'pede a confirmação caso o arquivo já exista.
  dlgOperacoes.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
  dlgOperacoes.Filter = "RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"
  dlgOperacoes.ShowSave
  
  If dlgOperacoes.FileName <> "" Then
    If UCase(Right(dlgOperacoes.FileName, 3)) = "RTF" Then
      rtfTexto.SaveFile dlgOperacoes.FileName, rtfRTF
    Else
      rtfTexto.SaveFile dlgOperacoes.FileName, rtfText
    End If
    FunSalvarArquivo = True
  End If

Cancelou:
End Function

Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
    '** Description:
    '** Print the active document
    On Error GoTo PrintError
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long

    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips

    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin

    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text

    ' Get length of text in RTF
    TextLength = Len(RTF.Text)

    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        Printer.NewPage                  ' Move on to next page
        Printer.Print Space(1) ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
    Loop

    ' Commit the print job
    Printer.EndDoc

    ' Allow the RTF to free up memory
    r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    Exit Sub

PrintError:
    MsgBox "Erro Inesperado ao Imprimir o Documento", vbCritical, "Erro de Impressão"
End Sub

