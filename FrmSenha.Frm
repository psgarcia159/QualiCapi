VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSenha 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   2850
   ClientTop       =   3615
   ClientWidth     =   3960
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1090
   Icon            =   "FrmSenha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   3960
   Begin Threed.SSPanel SSPanel2 
      Height          =   1980
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   3492
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      Begin VB.TextBox Senha_Ant 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         HelpContextID   =   1090
         IMEMode         =   3  'DISABLE
         Left            =   1305
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "Senha:"
         Top             =   885
         Width           =   2115
      End
      Begin VB.CommandButton Botao_Des 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Desiste"
         Height          =   300
         HelpContextID   =   1090
         Left            =   2115
         TabIndex        =   4
         Top             =   1455
         Width           =   1245
      End
      Begin VB.CommandButton Botao_Con 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Confirma"
         Default         =   -1  'True
         Height          =   300
         HelpContextID   =   1090
         Left            =   615
         TabIndex        =   3
         Top             =   1455
         Width           =   1245
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Left            =   225
         TabIndex        =   2
         Top             =   210
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Senha de Acesso"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   13.49
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelInner      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Senha:"
         Height          =   285
         Left            =   525
         TabIndex        =   5
         Top             =   930
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cont As Integer
Dim Tb_Operadores As Object
Dim sen As String
'Public Protecao As New clsProtecao

Private Sub Botao_Con_Click()
  Dim XLO_ESTOQUE As ADODB.Recordset
  Dim XLT_LETRA As String
  Dim XLT_LETRAASCII As String
  Dim XLT_SENHA As String
  Dim XLT_EMPRESA As String
  XLT_SENHA = Senha_Ant
  sen = ""
  'Cont = 1
  Dim XLI_CONT As Integer
  XLI_CONT = 1
  While (XLI_CONT <= Len(XLT_SENHA))
    XLT_LETRA = (Mid(XLT_SENHA, XLI_CONT, 1))
    XLT_LETRAASCII = Asc(XLT_LETRA) + 1
    XLT_LETRA = Chr$(XLT_LETRAASCII)
    sen = sen + XLT_LETRA
    XLI_CONT = XLI_CONT + 1
  Wend
    
  SubQOpenRecordset Tb_Operadores, _
    "SELECT " & _
    " func_cd_operador," & _
    " func_tx_senha," & _
    " func_tx_nome," & _
    " func_tx_funcao," & _
    " CentroResultados.* " & _
    "FROM" & _
    " CentroResultados RIGHT JOIN Funcionarios" & _
    " ON CentroResultados.cere_cd_Pcr = Funcionarios.cere_cd_Pcr " & _
    "WHERE func_tx_senha = '" & sen & "'", Estatico
  
  If Tb_Operadores.EOF Then
    MsgBox "Senha Inválida", 48, "ERRO"
    Senha_Ant.Text = ""
    sen = ""
    Senha_Ant.SetFocus
    Cont = Cont + 1
    If Cont = 3 Then
      Conexao.Close
      Set Conexao = Nothing
      'Setando a Proteção
     ' Protecao.SetLicense ("\\qdserver\Sistemas\Protecao\Qualidados\protecao.dat")
      End
    End If
  Else
    VariasVezes = True
    func_cd_operador = Trim$(Tb_Operadores!func_cd_operador)
    XGT_USUARIO = Trim$(Tb_Operadores!func_tx_nome)
    XGT_CARGO = Trim$(Tb_Operadores!func_tx_funcao)
     
    If (XGT_NomeAplic = "QUALICONT") Then
        MdiPrincipal.BarraStatus.Panels(8).Text = XGT_USUARIO
    Else
        MdiPrincipal.BarraStatus.Panels(7).Text = XGT_USUARIO
    End If
          
    If Not IsNull(Tb_Operadores("cere_cd_PCR").Value) Then XGL_CODPCR = Trim$(Tb_Operadores!cere_cd_Pcr)
    If Not IsNull(Tb_Operadores("cere_tx_descricao").Value) Then XGT_PCR = Trim$(Tb_Operadores!cere_tx_descricao)
    'A variável identificando o almoxarifado do usuário é o XGL_CODPCR
    XGT_ALMOXARIFADOUSUARIO = ""
    If Tb_Operadores!cere_bl_ContrAlmoxarifado = "S" Then
        If Not IsNull(Tb_Operadores!cere_cd_Pcr) Then XGT_ALMOXARIFADOUSUARIO = Trim$(Tb_Operadores!cere_cd_Pcr)
    End If

    Tb_Operadores.Close
    Set Tb_Operadores = Nothing
    

    'Leva ao formulário de multiempresa
    If pTipoCopia = "Multi" Then
        FrmMultiEmpresa.Show vbModal
    Else
        Unload Me
    End If
    
    'Definir variável identificando o código do Almoxarifado Central
    XGT_SQL = "SELECT * FROM CentroResultados WHERE empr_cd_Empresa = " & PCodEmpresa & _
              " AND cere_bl_ContrAlmoxarifado = 'S' ORDER BY cere_bl_AlmoxCentral DESC"
    SubQOpenRecordset XLO_ESTOQUE, XGT_SQL, Estatico
    If Not XLO_ESTOQUE.EOF Then XGT_ALMOXARIFADOCENTRAL = FunNulo(XLO_ESTOQUE!cere_cd_Pcr)
    XLO_ESTOQUE.Close
    Set XLO_ESTOQUE = Nothing
    
    Unload Me
  End If
End Sub
Private Sub Botao_Des_Click()
    MdiPrincipal.MousePointer = vbDefault
    Conexao.Close
    Set Conexao = Nothing
  '  Protecao.SetLicense ("\\qdserver\Sistemas\Protecao\Qualidados\protecao.dat")
    End
End Sub

Private Sub Form_Load()
  Cont = 0
  FrmSenha.Left = Int((ResolucaoX - FrmSenha.ScaleWidth) / 2)
  FrmSenha.Top = Int((Resolucaoy - FrmSenha.ScaleHeight) / 2)
  sen = ""
   If (XGT_NomeAplic = "QUALICONT") Then
        MdiPrincipal.BarraStatus.Panels(8).Text = ""
    Else
        MdiPrincipal.BarraStatus.Panels(7).Text = ""
    End If
End Sub

Private Sub Senha_Ant_KeyPress(KeyAscii As Integer)
  'Aceita apenas caracteres entre 0 - 9; A - Z ; a - z
'  If (KeyAscii = vbKeyBack) Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Then ' 97 = a ; 122 = z
'          KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'     sen = sen + Chr$(KeyAscii + 1)
'     'sen = sen + Chr$(KeyAscii)
'  Else
'     'KeyAscii = 0
'  End If
'  If (KeyAscii = vbKeyBack) Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Then ' 97 = a ; 122 = z
'    If (KeyAscii = vbKeyBack) Then
'        sen = Left(sen, (Len(sen) - 1))
'    Else
'        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'        sen = sen + Chr$(KeyAscii + 1)
'    End If
'  End If

End Sub

