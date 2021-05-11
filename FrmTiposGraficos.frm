VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form FrmTiposGraficos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Gráficos"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4845
      _Version        =   65536
      _ExtentX        =   8546
      _ExtentY        =   7435
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
      Begin Threed.SSCommand CmdLine 
         Height          =   855
         Left            =   2580
         TabIndex        =   2
         Top             =   1440
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":0000
      End
      Begin Threed.SSCommand CmdArea 
         Height          =   855
         Left            =   3780
         TabIndex        =   3
         Top             =   1440
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":0512
      End
      Begin Threed.SSCommand Cmd3DBars 
         Height          =   855
         Left            =   180
         TabIndex        =   4
         Top             =   2760
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":0A24
      End
      Begin Threed.SSCommand CmdPie 
         Height          =   855
         Left            =   1380
         TabIndex        =   5
         Top             =   2760
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":0F36
      End
      Begin Threed.SSCommand CmdMultiplePie 
         Height          =   855
         Left            =   2580
         TabIndex        =   6
         Top             =   2760
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":1448
      End
      Begin Threed.SSCommand CmdWeightedPie 
         Height          =   855
         Left            =   3780
         TabIndex        =   7
         Top             =   2760
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":195A
      End
      Begin Threed.SSCommand Cmd3DPercentBar 
         Height          =   855
         Left            =   1380
         TabIndex        =   8
         Top             =   1440
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   1508
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":1E6C
      End
      Begin Threed.SSCommand CmdPercentBar 
         Height          =   855
         Left            =   180
         TabIndex        =   9
         Top             =   1440
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":237E
      End
      Begin Threed.SSCommand cmd3DStackedbar 
         Height          =   855
         Left            =   3780
         TabIndex        =   10
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":2890
      End
      Begin Threed.SSCommand CmdStackedBar 
         Height          =   855
         Left            =   2580
         TabIndex        =   11
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":2DA2
      End
      Begin Threed.SSCommand Cmd3DSide 
         Height          =   855
         Left            =   1380
         TabIndex        =   12
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":32B4
      End
      Begin Threed.SSCommand CmdSideBySide 
         Height          =   855
         Left            =   180
         TabIndex        =   13
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "FrmTiposGraficos.frx":37C6
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Colunas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pizza Ponderada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3660
         TabIndex        =   24
         Top             =   3720
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pizzas Múltiplas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2460
         TabIndex        =   23
         Top             =   3720
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pizza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   22
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Blocos 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Área"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3660
         TabIndex        =   20
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Linha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   19
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Percentual 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1260
         TabIndex        =   18
         Top             =   2400
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Percentual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Empilhado 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3720
         TabIndex        =   16
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Empilhado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2520
         TabIndex        =   15
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Colunas 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1260
         TabIndex        =   14
         Top             =   1065
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3540
      TabIndex        =   0
      Top             =   4380
      Width           =   1215
   End
End
Attribute VB_Name = "FrmTiposGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd3DBars_Click()
  XGI_GRAFICO = crThreeDRegularGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub Cmd3DPercentBar_Click()
  XGI_GRAFICO = crFaked3DPercentBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub Cmd3DSide_Click()
  XGI_GRAFICO = crFaked3DSideBySideBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub cmd3DStackedbar_Click()
  XGI_GRAFICO = crFaked3DStackedBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdArea_Click()
  XGI_GRAFICO = 120
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
  XGB_RESPOSTA = False
  Unload Me
End Sub

Private Sub CmdLine_Click()
  XGI_GRAFICO = crRegularLineGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdMultiplePie_Click()
  XGI_GRAFICO = crMultiplePieGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdPercentBar_Click()
  XGI_GRAFICO = crPercentBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdPie_Click()
  XGI_GRAFICO = crRegularPieGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdSideBySide_Click()
  XGI_GRAFICO = crSideBySideBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdStackedBar_Click()
  XGI_GRAFICO = crStackedBarGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub CmdWeightedPie_Click()
  XGI_GRAFICO = crMultipleProportionalPieGraph
  XGB_RESPOSTA = True
  Unload Me
End Sub

Private Sub Form_Load()
  If XGI_GRAFICO_NRSUMS <= 1 Then
    Cmd3DBars.Enabled = False
    CmdPie.Enabled = False
    CmdMultiplePie.Enabled = False
    CmdWeightedPie.Enabled = False
  End If
End Sub
