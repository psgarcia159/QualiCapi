VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmFINVencResiduo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resíduo"
   ClientHeight    =   2070
   ClientLeft      =   3765
   ClientTop       =   2820
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel PanVencimento 
      Height          =   1620
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   2857
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Begin VB.Frame FraEscopo 
         Height          =   1035
         Left            =   270
         TabIndex        =   4
         Top             =   405
         Width           =   3300
         Begin MSComCtl2.DTPicker DtpVencResiduo 
            Height          =   300
            Left            =   1425
            TabIndex        =   5
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   63766529
            CurrentDate     =   37180
         End
         Begin VB.Label LblValor2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1425
            TabIndex        =   8
            Top             =   615
            Width           =   1245
         End
         Begin VB.Label LblValor 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor:"
            Height          =   255
            Left            =   450
            TabIndex        =   7
            Top             =   645
            Width           =   915
         End
         Begin VB.Label LblVencResiduo 
            Alignment       =   1  'Right Justify
            Caption         =   "Vencimento:"
            Height          =   255
            Left            =   450
            TabIndex        =   6
            Top             =   255
            Width           =   915
         End
      End
      Begin Threed.SSPanel PanInfo 
         Height          =   255
         Left            =   390
         TabIndex        =   3
         Top             =   165
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Deseja gerar parcela de resíduo?"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Não"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2093
      TabIndex        =   1
      Top             =   1740
      Width           =   1125
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Sim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   713
      TabIndex        =   0
      Top             =   1740
      Width           =   1125
   End
End
Attribute VB_Name = "FrmFINVencResiduo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdConfirmar_Click()
    FrmFINVencResiduo.Visible = False
    XGB_GERARESIDUO = True
End Sub

Private Sub Command1_Click()
    FrmFINVencResiduo.Visible = False
    XGB_GERARESIDUO = False
End Sub


