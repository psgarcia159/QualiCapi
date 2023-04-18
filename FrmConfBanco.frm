VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
																	   
Begin VB.Form FrmConfBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações do Sistema"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnlProgresso 
      Height          =   855
      Left            =   660
      TabIndex        =   34
      Top             =   3780
      Visible         =   0   'False
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   5212
      _ExtentY        =   1508
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ProgressBar prgBar 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   420
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblProgresso 
         Caption         =   "Registrando..."
         Height          =   195
         Left            =   480
         TabIndex        =   36
         Top             =   -1440
         Width           =   2715
      End
   End
   Begin MSComDlg.CommonDialog dlgLocalBanco 
      Left            =   180
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3660
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4980
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   7329
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conexão do Banco"
      TabPicture(0)   =   "FrmConfBanco.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmDadosBanco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmTipoBanco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Registro do Banco"
      TabPicture(1)   =   "FrmConfBanco.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmRegistro"
											 
      Tab(1).Control(1)=   "Label4"
											 
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Atualização"
      TabPicture(2)   =   "FrmConfBanco.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraCaminho"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      Begin VB.Frame FraCaminho 
         Height          =   2895
         Left            =   -74940
         TabIndex        =   37
         Top             =   1260
         Visible         =   0   'False
         Width           =   6135
         Begin VB.DirListBox dirPasta 
            Height          =   1665
            Left            =   1440
            TabIndex        =   31
            Top             =   660
            Width           =   3195
         End
         Begin VB.DriveListBox drvPasta 
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Top             =   300
            Width           =   3255
         End
         Begin VB.CommandButton CmdCaminhoOk 
            Caption         =   "&OK"
            Height          =   315
            Left            =   1440
            TabIndex        =   32
            Top             =   2460
            Width           =   1035
         End
         Begin VB.CommandButton CmdCaminhoCancelar 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   3600
            TabIndex        =   33
            Top             =   2460
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   6135
         Begin VB.TextBox txtPastaAtualizacoes 
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   540
            Width           =   5415
         End
         Begin VB.CommandButton cmdPastaAtualizacoes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            Picture         =   "FrmConfBanco.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   540
            Width           =   390
         End
         Begin VB.Label Label3 
            Caption         =   "Pasta dos arquivos de atualizações:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame frmTipoBanco 
         Caption         =   " Tipo do Banco: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   5955
         Begin VB.OptionButton optSql 
            Caption         =   "SQL Server 2000 / MSDE"
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton optAccess 
            Caption         =   "Access 2000"
            Height          =   255
            Left            =   3360
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame frmDadosBanco 
         Caption         =   " Dados para acesso: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   180
         TabIndex        =   4
         Top             =   1440
         Width           =   5955
         Begin VB.TextBox txtSenha 
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1620
            PasswordChar    =   "*"
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtUsuario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   38
            Top             =   1260
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtLocalBanco 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            TabIndex        =   6
            Top             =   420
            Width           =   3855
         End
         Begin VB.CommandButton cmdProcurarBanco 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   420
            Width           =   315
         End
         Begin VB.TextBox txtNomeBanco 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Text            =   "QualiAdmFin"
            Top             =   840
            Width           =   3855
         End
         Begin VB.CommandButton txtVerifBanco 
            Caption         =   "Verificar..."
            Height          =   315
            Left            =   4520
            TabIndex        =   10
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblSenha 
            Alignment       =   1  'Right Justify
            Caption         =   "Senha:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblUsuario 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1380
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblLocalBanco 
            Alignment       =   1  'Right Justify
            Caption         =   "Local do Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome do Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame frmRegistro 
         Caption         =   " Dados para o registro: "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74880
         TabIndex        =   14
         Top             =   1020
         Width           =   6075
         Begin VB.TextBox txtSenhaSA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2640
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   2400
            Width           =   1755
         End
         Begin VB.TextBox txtLoginSA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Text            =   "sa"
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CommandButton cmdDesRegistrar 
            Caption         =   "Desregistrar"
            Height          =   315
            Left            =   4680
            TabIndex        =   25
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtNomeServidor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   16
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdRegistrar 
            Caption         =   "Registrar"
            Height          =   315
            Left            =   4680
            TabIndex        =   24
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtLocalBancoServ 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   18
            Top             =   780
            Width           =   3615
         End
         Begin VB.Label Label5 
            Caption         =   "Senha de acesso:"
            Height          =   195
            Left            =   2640
            TabIndex        =   22
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Nome do Usuário:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome do Servidor:"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label7 
            Caption         =   $"FrmConfBanco.frx":0103
            Height          =   615
            Left            =   240
            TabIndex        =   19
            Top             =   1500
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Local do Banco no Servidor:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   900
            Width           =   2055
         End
      End
      Begin VB.Label Label4 
         Caption         =   $"FrmConfBanco.frx":018A
         Height          =   435
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   5835
      End
   End
End
Attribute VB_Name = "FrmConfBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFO_CONEXAO As New ADODB.Connection

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Sub CmdCaminhoCancelar_Click()
  
  FraCaminho.Visible = False
  
End Sub

Private Sub CmdCaminhoOk_Click()
  
  txtPastaAtualizacoes.Text = dirPasta.List(dirPasta.ListIndex)
  FraCaminho.Visible = False
  
End Sub

Private Sub CmdCancelar_Click()
  XGB_RESPOSTA = False
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim XLT_TIPO As String, XLT_ERRO As String

  Screen.MousePointer = vbHourglass
  If Not funTestaConexao(XLT_ERRO) Then
    MsgBox "Não foi possível se conectar ao banco de dados." & Chr(13) & XLT_ERRO, _
      vbCritical, "Erro"
  Else
    If optAccess.Value = True Then
      XLT_TIPO = 1
    Else
      XLT_TIPO = 2
    End If
    
    'GRAVA O TIPO DO BANCO
    If WritePrivateProfileString("BancoDeDados", "TipoBanco", XLT_TIPO, App.Path + "\QualiAdmFin.INI") = 0 Then
      MsgBox "Problemas na gravação do arquivo de configuração", vbCritical, "Configuração"
    End If
    'GRAVA O ENDEREÇO DO BANCO
    WritePrivateProfileString "BancoDeDados", "Endereco", txtLocalBanco.Text, App.Path + "\QualiAdmFin.INI"
    'GRAVA O NOME DO BANCO
    WritePrivateProfileString "BancoDeDados", "NomeBanco", txtNomeBanco.Text, App.Path + "\QualiAdmFin.INI"
    'GRAVA O USUARIO
    WritePrivateProfileString "BancoDeDados", "Usuario", txtUsuario.Text, App.Path + "\QualiAdmFin.INI"
    'GRAVA A SENHA
    WritePrivateProfileString "BancoDeDados", "Senha", txtSenha.Text, App.Path + "\QualiAdmFin.INI"
    
    'Grava o local da atualização
    If Dir(txtPastaAtualizacoes.Text, vbDirectory) <> "" Then
      WritePrivateProfileString "Atualizacao", "Pasta", txtPastaAtualizacoes.Text, App.Path + "\QualiAdmFin.INI"
    End If
      
    MsgBox "O sistema deve ser reiniciado para que as alterações sejam efetivadas", vbInformation, "Reiniciar"
    
    XGB_RESPOSTA = True
    Unload Me
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPastaAtualizacoes_Click()
  
  FraCaminho.Visible = True
  FraCaminho.Left = 120
  FraCaminho.Top = 360
  FraCaminho.Height = 2895
  FraCaminho.Width = 6135
  
End Sub

Private Sub cmdProcurarBanco_Click()
  dlgLocalBanco.DialogTitle = "Local do Banco de Dados"
  dlgLocalBanco.DefaultExt = "mdb"
  dlgLocalBanco.Filter = "Access (*.mdb)|*.mdb"
  dlgLocalBanco.InitDir = XGT_DirSistema
  dlgLocalBanco.CancelError = False
  dlgLocalBanco.ShowOpen
  txtLocalBanco.Text = dlgLocalBanco.FileName
End Sub

Private Sub cmdRegistrar_Click()
  On Error GoTo Erro
  
  Screen.MousePointer = vbHourglass
  subMostraProgresso
  'Registra o banco
  lblProgresso.Caption = "Registrando o banco..."
  lblProgresso.Refresh
  subAnexaBanco 'Atacha o banco usando o usuário sa
  prgBar.Value = 25
  'Cria o usuário QualiAdmFin no servidor
  lblProgresso.Caption = "Criando o usuário..."
  lblProgresso.Refresh
  subRevokeDBAccess 'Retira o usuário qualiadmfin
  subAddLogin 'Cria o usuário qualiadmfin
  prgBar.Value = 50
  'Adiciona o usuário QualiAdmFin ao banco QualiAdmFin
  lblProgresso.Caption = "Adicionando o usuário..."
  lblProgresso.Refresh
  subAddUser
  prgBar.Value = 75
  'Habilita o usuário no banco QualiAdmFin
  lblProgresso.Caption = "Adicionando permissões..."
  lblProgresso.Refresh
  subAddRoleMember
  prgBar.Value = 100
  'Finaliza
  MsgBox "Banco de dados registrado no servidor com sucesso.", vbInformation, "Registro"
  pnlProgresso.Visible = False
  Screen.MousePointer = vbDefault
  Exit Sub

Erro:
  MsgBox "Ocorreu um erro ao tentar registrar o banco no servidor." & Chr(13) & _
    Err.Number & " - " & Err.Description, vbCritical, "Erro"
  Screen.MousePointer = vbDefault
  pnlProgresso.Visible = False
End Sub

Private Sub cmdDesRegistrar_Click()
  On Error GoTo Erro
  subDesanexaBanco
  subApagaLogin
  'subRemoveUsuario
  MsgBox "Banco de dados desregistrado no servidor com sucesso.", vbInformation, "Registro"
  Exit Sub

Erro:
  MsgBox "Ocorreu um erro ao tentar desregistrar o banco no servidor." & Chr(13) & _
    Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub drvPasta_Change()

    On Error GoTo Erro

    dirPasta.Path = drvPasta.Drive

Erro:
    If Len(Err.Description) > 0 Then
        MsgBox Err.Description, vbCritical, "Quali Update"
    End If
End Sub

Private Sub Form_Load()
  Dim XLT_STRINGCONEXAO
  Dim XLT_TIPOBANCO
  Dim XLT_NOMEBANCO
  Dim XLT_USUARIO
  Dim XLT_SENHA
  Dim XLT_TEMP As String * 254
  Dim XLT_PASTA As String
  
  'PEGA O TIPO DO BANCO
  If GetPrivateProfileString("BancoDeDados", "TipoBanco", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
    MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
  End If
  XLT_TIPOBANCO = funTiraCaracteresNulos(XLT_TEMP)
  
  'PEGA O ENDEREÇO DO BANCO
  If GetPrivateProfileString("BancoDeDados", "Endereco", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
    MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
  End If
  XLT_STRINGCONEXAO = funTiraCaracteresNulos(XLT_TEMP)
    
  'PEGA O NOME DO BANCO
  If GetPrivateProfileString("BancoDeDados", "NomeBanco", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
    MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
  End If
  XLT_NOMEBANCO = funTiraCaracteresNulos(XLT_TEMP)
  
  'PEGA O USUARIO
  If GetPrivateProfileString("BancoDeDados", "Usuario", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
    MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
  End If
  XLT_USUARIO = funTiraCaracteresNulos(XLT_TEMP)
    
  'PEGA A SENHA
  If GetPrivateProfileString("BancoDeDados", "Senha", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
    MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
  End If
  XLT_SENHA = funTiraCaracteresNulos(XLT_TEMP)
    
  
  If GetPrivateProfileString("Atualizacao", "Pasta", "", XLT_TEMP, 255, App.Path + "\QualiAdmFin.INI") <> 0 Then
    XLT_PASTA = Left$(XLT_TEMP, InStr(XLT_TEMP, Chr$(0)) - 1)
  End If
    

  If XLT_TIPOBANCO = 1 Then
    optAccess.Value = True
  Else
    optSql.Value = True
  End If
  
  txtLocalBanco.Text = Trim(XLT_STRINGCONEXAO)
  txtNomeBanco.Text = Trim(XLT_NOMEBANCO)
  txtNomeServidor.Text = funNomeComp
  txtLocalBancoServ.Text = App.Path
  txtPastaAtualizacoes.Text = XLT_PASTA
  txtUsuario.Text = Trim(XLT_USUARIO)
  txtSenha.Text = Trim(XLT_SENHA)
      
End Sub

Private Sub optAccess_Click()
  lblLocalBanco.Caption = "Local do Banco:"
  txtLocalBanco.Text = ""
  frmRegistro.Enabled = False
  cmdProcurarBanco.Visible = True
  SSTab1.TabVisible(1) = False
  lblUsuario.Visible = False
  txtUsuario.Enabled = False
  txtUsuario.Visible = False
  LblSenha.Visible = False
  txtSenha.Enabled = False
  txtSenha.Visible = False
End Sub

Private Sub optSql_Click()
  lblLocalBanco.Caption = "Nome do Servidor:"
  txtLocalBanco.Text = ""
  frmRegistro.Enabled = True
  cmdProcurarBanco.Visible = False
  SSTab1.TabVisible(1) = True
  lblUsuario.Visible = True
  txtUsuario.Enabled = True
  txtUsuario.Visible = True
  LblSenha.Visible = True
  txtSenha.Enabled = True
  txtSenha.Visible = True
End Sub

Private Sub txtVerifBanco_Click()
  Dim XLT_ERRO As String
  
  Screen.MousePointer = vbHourglass
  If funTestaConexao(XLT_ERRO) Then
    MsgBox "Conexão efetuada com sucesso", vbInformation
  Else
    MsgBox "Não foi possível se conectar ao banco de dados." & Chr(13) & XLT_ERRO, _
      vbCritical, "Erro"
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub subMostraProgresso()
  pnlProgresso.Visible = True
  lblProgresso.Caption = ""
  prgBar.Value = 0
End Sub

Private Function funTestaConexao(XLT_ERRO As String) As Boolean
  On Error GoTo Erro
  
  If optAccess.Value Then
    XFO_CONEXAO.Provider = "Microsoft.Jet.OLEDB.4.0"
    XFO_CONEXAO.Open txtLocalBanco.Text, "Admin", ""
  Else
    XFO_CONEXAO.Provider = "sqloledb"
    XFO_CONEXAO.Open _
      "Data Source=" & txtLocalBanco.Text & _
      ";Initial Catalog=" & txtNomeBanco.Text & _
      ";User Id=" & txtUsuario.Text & "; Password=" & txtSenha.Text & ";"
  End If
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  funTestaConexao = True
  Exit Function

Erro:
  XLT_ERRO = Err.Number & " - " & Err.Description
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  funTestaConexao = False
End Function

Private Sub subAnexaBanco()
  Dim XLO_COMANDO As ADODB.Command
  
  On Error GoTo Erro
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=master" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_attach_db"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria automaticamente todos os parâmetros dentro do objeto Comando
    .Parameters.Refresh
    'Passa os valores de alguns parâmetros
    'Nome do Banco
    .Parameters(1).Value = "QualiAdmFin"
    'Local do MDF (data file)
    .Parameters(2).Value = txtLocalBancoServ.Text & "\QualiAdmFin_data.mdf"
    'Local do LDF (log file)
    .Parameters(3).Value = txtLocalBancoServ.Text & "\QualiAdmFin_log.ldf"
    'Executa a stored procedure
    .Execute
    If .Parameters(0).Value = 1 Then
      If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
      Err.Raise 6, , "Erro ao executar sp_attach_db"
    End If
  End With
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub

Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Sub subDesanexaBanco()
  Dim XLO_COMANDO As ADODB.Command
  
  On Error GoTo Erro
  
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=master" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_detach_db"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria automaticamente todos os parâmetros dentro do objeto Comando
    .Parameters.Refresh
    'Passa os valores de alguns parâmetros
    'Nome do Banco
    .Parameters(1).Value = "QualiAdmFin"
    'Executa a stored procedure
    .Execute
    If .Parameters(0).Value = 1 Then
      If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
      Err.Raise 6, , "Erro ao executar sp_attach_db"
    End If
  End With
  'Finaliza o processo
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub
  
Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Sub subAddLogin()
  Dim XLO_COMANDO As ADODB.Command
  
  On Error GoTo Erro
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=master" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_addlogin"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria automaticamente todos os parâmetros dentro do objeto Comando
    .Parameters.Refresh
    'Passa os valores de alguns parâmetros
    'Nome do usuário a ser criado
    .Parameters(1).Value = "QualiAdmFin"
    'Senha (opcional)
    .Parameters(2).Value = "qd"
    'Executa a stored procedure
    .Execute
    If .Parameters(0).Value = 1 Then
      If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
      Err.Raise 6, , "Erro ao executar sp_addlogin"
    End If
  End With
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub

Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Sub subApagaLogin()
  Dim XLO_COMANDO As ADODB.Command
  
  On Error GoTo Erro

  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=master" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_droplogin"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria automaticamente todos os parâmetros dentro do objeto Comando
    .Parameters.Refresh
    'Passa os valores de alguns parâmetros
    'Nome do usuário a ser criado
    .Parameters(1).Value = "QualiAdmFin"
    'Executa a stored procedure
    .Execute
    If .Parameters(0).Value = 1 Then
      If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
'      Err.Raise 6, , "Erro ao executar sp_droplogin"
    End If
  End With
  Set XLO_COMANDO = Nothing
  Exit Sub

Erro:
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Err.Raise Err.Number
End Sub

Private Sub subAddUser()
  Dim XLO_COMANDO As ADODB.Command
  Dim XLO_PARAMETRO As ADODB.Parameter
  
  On Error GoTo Erro
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=QualiAdmFin" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_grantdbaccess"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria o parâmetro necessário e passa o valor
    Set XLO_PARAMETRO = .CreateParameter("loginame", adVarChar, adParamInput, 255)
    .Parameters.Append XLO_PARAMETRO
    XLO_PARAMETRO.Value = "QualiAdmFin"
    'Executa a stored procedure
    .Execute
  End With
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub

Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Sub subRevokeDBAccess()
  Dim XLO_COMANDO As ADODB.Command
  Dim XLO_PARAMETRO As ADODB.Parameter
  
  On Error GoTo Erro
  
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=QualiAdmFin" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_revokedbaccess"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria o parâmetro necessário e passa o valor
    Set XLO_PARAMETRO = .CreateParameter("name_in_db", adVarChar, adParamInput, 255)
    .Parameters.Append XLO_PARAMETRO
    XLO_PARAMETRO.Value = "QualiAdmFin"
    'Executa a stored procedure
    .Execute
  End With
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub
  
Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Sub subAddRoleMember()
  Dim XLO_COMANDO As ADODB.Command
  Dim XLO_PARAMETRO As ADODB.Parameter
  
  On Error GoTo Erro
  'Abre a conexao
  XFO_CONEXAO.Provider = "sqloledb"
  XFO_CONEXAO.Open _
    "Data Source=" & txtNomeServidor.Text & _
    ";Initial Catalog=QualiAdmFin" & _
    ";User Id=" & txtLoginSA.Text & ";Password=" & txtSenhaSA.Text & ";"
  
  'Cria o objeto Comando para executar a stored procedure
  Set XLO_COMANDO = New ADODB.Command
  With XLO_COMANDO
    .CommandText = "sp_addrolemember"
    .CommandType = adCmdStoredProc
    Set XLO_COMANDO.ActiveConnection = XFO_CONEXAO
    'Cria o parâmetro necessário e passa o valor
    'Tipo de permissão
    Set XLO_PARAMETRO = .CreateParameter("rolename", adVarChar, adParamInput, 255)
    .Parameters.Append XLO_PARAMETRO
    XLO_PARAMETRO.Value = "db_owner"
    'Nome do usuário que terá a permissão
    Set XLO_PARAMETRO = .CreateParameter("membername", adVarChar, adParamInput, 255)
    .Parameters.Append XLO_PARAMETRO
    XLO_PARAMETRO.Value = "QualiAdmFin"
    'Executa a stored procedure
    .Execute
  End With
  If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  Exit Sub

Erro:
  If Err.Number <> -2147217900 Then
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
    Err.Raise Err.Number
  Else
    If XFO_CONEXAO.State = adStateOpen Then XFO_CONEXAO.Close
  End If
End Sub

Private Function funNomeComp() As String
  Dim strString As String
  
  'Create a buffer
  strString = String(255, Chr$(0))
  'Get the computer name
  GetComputerName strString, 255
  'remove the unnecessary chr$(0)'s
  strString = Left$(strString, InStr(1, strString, Chr$(0)) - 1)
  'Show the computer name
  funNomeComp = strString
End Function
