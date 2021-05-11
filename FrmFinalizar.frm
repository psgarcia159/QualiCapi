VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmFinalizar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1920
   ClientLeft      =   2445
   ClientTop       =   2535
   ClientWidth     =   3975
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
   Icon            =   "FrmFinalizar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   3975
   Begin Threed.SSPanel PanEncerrar 
      Height          =   1950
      Left            =   0
      TabIndex        =   4
      Top             =   -30
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   3440
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton CmdNao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Não"
         Height          =   300
         HelpContextID   =   1090
         Left            =   2145
         TabIndex        =   3
         Top             =   1440
         Width           =   1245
      End
      Begin VB.CommandButton CmdSim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Sim"
         Default         =   -1  'True
         Height          =   300
         HelpContextID   =   1090
         Left            =   630
         TabIndex        =   2
         Top             =   1440
         Width           =   1245
      End
      Begin VB.OptionButton OptFinalizar 
         Caption         =   "Finalizar o Programa?"
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
         Left            =   465
         TabIndex        =   0
         Top             =   690
         Value           =   -1  'True
         Width           =   3090
      End
      Begin VB.OptionButton OptReiniciar 
         Caption         =   "Efetuar logon com usuário diferente?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   465
         TabIndex        =   1
         Top             =   1050
         Width           =   3090
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Left            =   210
         TabIndex        =   5
         Top             =   150
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Encerrar WorkFlow"
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
   End
End
Attribute VB_Name = "FrmFinalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public Protecao As New clsProtecao

Private Sub CmdNao_Click()
    Finaliza = True
    Unload Me
End Sub
Private Sub CmdSim_Click()
    If OptFinalizar.Value Then
        'Encerra a conexao com o banco de dados e Finaliza aplicação
        Conexao.Close
        ConexaoRelatorio.Close
       ' Protecao.SetLicense ("\\qdserver\Sistemas\Protecao\Qualidados\protecao.dat")
        End
    Else
        'Altera usuário
        Finaliza = True
        Conexao.Close
        ConexaoRelatorio.Close
        Unload Me
        Call Main
        FrmSenha.Show 1
    End If
End Sub
Private Sub Form_Load()
    SSPanel1.Caption = "Encerra " & App.ProductName
    FrmFinalizar.Left = Int((ResolucaoX - FrmFinalizar.ScaleWidth) / 2)
    FrmFinalizar.Top = Int((Resolucaoy - FrmFinalizar.ScaleHeight) / 2)
End Sub
