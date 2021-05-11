VERSION 5.00
Object = "{0DF767BD-777E-11D5-AF6E-008048E23503}#1.0#0"; "SHELLVIEW.OCX"
Begin VB.Form FrmSelecionarPasta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Características do Arquivo de Exportação"
   ClientHeight    =   4815
   ClientLeft      =   4290
   ClientTop       =   2955
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNomeArquivo 
      Height          =   330
      Left            =   1785
      TabIndex        =   0
      Top             =   150
      Width           =   2820
   End
   Begin ShellView.FolderTree FolderTree1 
      Height          =   2835
      Left            =   180
      TabIndex        =   5
      Top             =   1440
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   5001
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Indentation     =   120,189
   End
   Begin VB.TextBox TxtCaminho 
      Height          =   330
      Left            =   180
      TabIndex        =   4
      Top             =   945
      Width           =   4455
   End
   Begin VB.CommandButton CmdRetornar 
      Caption         =   "&Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3675
      TabIndex        =   3
      Top             =   4395
      Width           =   975
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2595
      TabIndex        =   2
      Top             =   4395
      Width           =   1005
   End
   Begin VB.CommandButton CmdNovaPasta 
      Caption         =   "&Nova Pasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   195
      TabIndex        =   1
      Top             =   4410
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Arquivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   7
      Top             =   210
      Width           =   1845
   End
   Begin VB.Label LblConvenio 
      Caption         =   "Selecione uma pasta para gravação do arquivo de Exportação:"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   4500
   End
End
Attribute VB_Name = "FrmSelecionarPasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdNovaPasta_Click()
  FolderTree1.CreateNewFolder
End Sub

Private Sub CmdConfirmar_Click()
    If TxtNomeArquivo.Text = "" Then
        MsgBox "Nome do arquivo deve ser definido.", vbCritical + vbOKOnly, "ATENÇÃO"
        TxtNomeArquivo.SetFocus
        Exit Sub
    End If
    If TypeName(FolderTree1.SelectedFolder) = "Folder" Then
        m_SelectedPath = FolderTree1.SelectedFolder.Items.Item.Path
    Else
        m_SelectedPath = FolderTree1.SelectedFolder.Path
    End If
  
    XGT_ARQUIVO = TxtNomeArquivo.Text
    XGT_LOCALARQ = TxtCaminho.Text & "\"
    
    Unload Me
End Sub

Private Sub CmdRetornar_Click()
  m_SelectedPath = ""
  Unload Me
End Sub

Private Sub FolderTree1_SelectionChange(ByVal SelFolder As Object)
   On Error Resume Next
   TxtCaminho.Text = SelFolder.Items.Item.Path
   CmdConfirmar.Enabled = SelFolder.Items.Item.IsFileSystem
End Sub

Private Sub Form_Load()
   CmdConfirmar.Enabled = False
   With FolderTree1
        Set .RootFolder = .ShellNameSpace(0) '0 - desktop
        Set .SelectedFolder = .ShellNameSpace(App.Path)
        TxtNomeArquivo.Text = XGT_ARQUIVO
   End With
End Sub
