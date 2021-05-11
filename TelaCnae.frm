VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaCnae 
   Caption         =   "Tabela de Seção"
   ClientHeight    =   4545
   ClientLeft      =   2010
   ClientTop       =   1245
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   9390
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel panClasse 
      Height          =   2475
      Left            =   600
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   8160
      _Version        =   65536
      _ExtentX        =   14393
      _ExtentY        =   4366
      _StockProps     =   15
      Caption         =   " Cadastro de Classe"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin MSMask.MaskEdBox txtCodigoClasse 
         Height          =   345
         Left            =   1140
         TabIndex        =   47
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "9999-9"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdRetornarClasse 
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
         Height          =   375
         Left            =   7080
         TabIndex        =   51
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdConfirmarClasse 
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
         Height          =   375
         Left            =   6000
         TabIndex        =   50
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtDescricaoClasse 
         Height          =   555
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   900
         Width           =   6885
      End
      Begin MSDataListLib.DataCombo CboGrupo 
         Bindings        =   "TelaCnae.frx":0000
         Height          =   330
         Left            =   1140
         TabIndex        =   49
         Top             =   1560
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cngr_tx_Codigo"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblGrupo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblClasse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   36
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblDescricaoClasse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   930
         Width           =   915
      End
   End
   Begin Threed.SSPanel panSubClasse 
      Height          =   2475
      Left            =   600
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   8160
      _Version        =   65536
      _ExtentX        =   14393
      _ExtentY        =   4366
      _StockProps     =   15
      Caption         =   " Cadastro de Sub-Classe"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin MSMask.MaskEdBox txtCodigoSubclasse 
         Height          =   395
         Left            =   1140
         TabIndex        =   39
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9/99"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDescricaoSubclasse 
         Height          =   555
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   960
         Width           =   6885
      End
      Begin VB.CommandButton cmdConfirmarSubclasse 
         Caption         =   "&Confirmar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         TabIndex        =   42
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdRetornarSubclasse 
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
         Height          =   375
         Left            =   7080
         TabIndex        =   44
         Top             =   2040
         Width           =   975
      End
      Begin MSDataListLib.DataCombo CboClasse 
         Bindings        =   "TelaCnae.frx":0017
         Height          =   330
         Left            =   1140
         TabIndex        =   41
         Top             =   1560
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cncl_tx_Codigo"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDescricaoSubclasse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   46
         Top             =   930
         Width           =   915
      End
      Begin VB.Label cmdCodigoSubclasse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   540
         Width           =   915
      End
      Begin VB.Label txtClasse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Classe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   915
      End
   End
   Begin Threed.SSPanel PanGrupo 
      Height          =   2475
      Left            =   600
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   8160
      _Version        =   65536
      _ExtentX        =   14393
      _ExtentY        =   4366
      _StockProps     =   15
      Caption         =   " Cadastro de Grupo"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtDescricaoGrupo 
         Height          =   555
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   900
         Width           =   6885
      End
      Begin VB.TextBox txtCodigoGrupo 
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   26
         Top             =   510
         Width           =   765
      End
      Begin VB.CommandButton cmdConfirmarGrupo 
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
         Height          =   375
         Left            =   5880
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdRetornarGrupo 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   30
         Top             =   2040
         Width           =   975
      End
      Begin MSDataListLib.DataCombo CboDivisao 
         Bindings        =   "TelaCnae.frx":002F
         Height          =   330
         Left            =   1140
         TabIndex        =   28
         Top             =   1560
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cndi_tx_Codigo"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDescricaoGrupo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   930
         Width           =   915
      End
      Begin VB.Label lblCodigoGrupo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblDivisao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Divisão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   915
      End
   End
   Begin Threed.SSPanel PanDivisao 
      Height          =   2715
      Left            =   600
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   8160
      _Version        =   65536
      _ExtentX        =   14393
      _ExtentY        =   4789
      _StockProps     =   15
      Caption         =   " Cadastro de Divisão"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.CommandButton cmdRetornarDivisao 
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
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton CmdConfirmarDivisao 
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
         Height          =   375
         Left            =   6000
         TabIndex        =   20
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtCodigoDivisao 
         Height          =   315
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   17
         Top             =   510
         Width           =   525
      End
      Begin VB.TextBox txtDescricaoDivisao 
         Height          =   795
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   900
         Width           =   6885
      End
      Begin MSDataListLib.DataCombo cboSecao 
         Bindings        =   "TelaCnae.frx":0048
         Height          =   330
         Left            =   1140
         TabIndex        =   19
         Top             =   1800
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cnse_tx_Codigo"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSecao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seção:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblCodigoDivisao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblDescricaoDivisao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   930
         Width           =   915
      End
   End
   Begin Threed.SSPanel PanSecao 
      Height          =   2595
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   6960
      _Version        =   65536
      _ExtentX        =   12277
      _ExtentY        =   4577
      _StockProps     =   15
      Caption         =   " Cadastro de Seção"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtDescricaoSecao 
         Height          =   1035
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   5685
      End
      Begin VB.TextBox TxtCodigoSecao 
         Height          =   315
         Left            =   1140
         MaxLength       =   1
         TabIndex        =   9
         Top             =   510
         Width           =   405
      End
      Begin VB.CommandButton CmdConfirmarSecao 
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
         Height          =   375
         Left            =   4770
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton CmdRetornarSecao 
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
         Height          =   375
         Left            =   5850
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label LblDescricaoSecao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   930
         Width           =   915
      End
      Begin VB.Label LblCodigoSecao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   540
         Width           =   915
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaCnae.frx":005F
      Height          =   3690
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   27
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "cnse_tx_Codigo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descrição"
      Columns(1).DataField=   "cnse_tx_Descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Código"
      Columns(2).DataField=   "cndi_tx_Codigo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Descrição"
      Columns(3).DataField=   "cndi_tx_Descricao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Código"
      Columns(4).DataField=   "cngr_tx_Codigo"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Descrição"
      Columns(5).DataField=   "cngr_tx_Descricao"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Código"
      Columns(6).DataField=   "cncl_tx_Codigo"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Descrição"
      Columns(7).DataField=   "cncl_tx_Descricao"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Código"
      Columns(8).DataField=   "cnsc_tx_Codigo"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Descrição"
      Columns(9).DataField=   "cnsc_tx_Descricao"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=20"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=20"
      Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=20"
      Splits(0)._ColumnProps(28)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=20"
      Splits(0)._ColumnProps(36)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(39)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=20"
      Splits(0)._ColumnProps(44)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(47)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(6)._ColStyle=20"
      Splits(0)._ColumnProps(52)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(54)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(55)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(56)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(58)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(7)._ColStyle=20"
      Splits(0)._ColumnProps(60)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(62)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(63)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(64)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(66)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(8)._ColStyle=20"
      Splits(0)._ColumnProps(68)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(69)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(70)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(71)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(75)=   "Column(9)._ColStyle=20"
      Splits(0)._ColumnProps(76)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(77)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(78)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14,.bold=-1,.fontsize=825"
      _StyleDefs(38)  =   ":id=67,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(39)  =   ":id=67,.fontname=MS Sans Serif"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=28,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(83)  =   ":id=34,.strikethrough=0,.charset=0"
      _StyleDefs(84)  =   ":id=34,.fontname=MS Sans Serif"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(91)  =   ":id=37,.strikethrough=0,.charset=0"
      _StyleDefs(92)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   150
      TabIndex        =   7
      Top             =   4140
      Width           =   9225
      _Version        =   65536
      _ExtentX        =   16272
      _ExtentY        =   635
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   8640
         Picture         =   "TelaCnae.frx":0074
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7590
         TabIndex        =   6
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Im&primir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6345
         TabIndex        =   5
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdOrdem 
         Caption         =   "&Ordem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5100
         TabIndex        =   4
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton Botao_Fil 
         Caption         =   "Se&leções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3855
         TabIndex        =   3
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "&Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2610
         TabIndex        =   2
         Top             =   0
         Width           =   1020
      End
      Begin VB.CommandButton CmdInserir 
         BackColor       =   &H00000000&
         Caption         =   "&Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdAlterar 
         BackColor       =   &H00000000&
         Caption         =   "&Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   1
         Top             =   15
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc DatSecao 
      Height          =   375
      Left            =   -120
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatBanco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatDivisao 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatBanco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatGrupo 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatBanco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatClasse 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatBanco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatSubClasse 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatBanco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaCnae.frx":0223
      Height          =   315
      Left            =   2760
      TabIndex        =   52
      Top             =   3840
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "iden_tx_descricao"
      BoundColumn     =   "iden_cd_Projeto"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Descrição:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   3840
      Visible         =   0   'False
      Width           =   2670
   End
End
Attribute VB_Name = "TelaCnae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Tipos() As String * 1
Dim Quantid As Long
Dim ResFormulario As Object
Dim KeyAscii As Integer
Dim errLoop As ADODB.Error
Dim StrError As String
Dim x As Integer

Dim XLT_TELA As String ' Indica se é a tela de Seção, Grupo, Divisão, Classe ou Sub-Classe
    
Dim Chave As Integer
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean
Dim ResExiste As ADODB.Recordset
Dim FRM_ContFil As Integer ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long  ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer    ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean    'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem

Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        FRM_TipoVar = TipoVar
        FRM_FiltroOrig = FiltroOrig
        FRM_FiltroAtual = FiltroAtual
        FRM_filtra_especial = filtra_especial
        FRM_Tem_Filtro = Tem_Filtro
        FRM_ConTelaOrdem = ConTelaOrdem
        FRM_Quantid = Quantid
        FRM_Filtrou = Filtrou
        FRM_BarraFerramentasPressionada = BarraFerramentasPressionada
        FRM_VariasVezes = VariasVezes
        FRM_XInserir = XInserir
        FRM_XOrigem = XOrigem
        FRM_SQLAtivo = SQLAtivo
        
        ReDim FRM_Filtros(UBound(Filtros, 1), UBound(Filtros, 2))
        ReDim FRM_Ordens(UBound(Ordens, 1)) As Integer
        
        For Count = LBound(Filtros, 1) To UBound(Filtros, 1)
            For Count2 = LBound(Filtros, 2) To UBound(Filtros, 2)
                FRM_Filtros(Count, Count2) = Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(Ordens) To UBound(Ordens)
            FRM_Ordens(Count) = Ordens(Count)
        Next Count
    ElseIf GuardaOUPegaValores = "P" Then  'Regravando os valores relativos a tela atual variáveis de definição da tela
        ContFil = FRM_ContFil
        Ordem = FRM_Ordem
        Campos_Obrig = FRM_Campos_Obrig
        TipoVar = FRM_TipoVar
        FiltroOrig = FRM_FiltroOrig
        FiltroAtual = FRM_FiltroAtual
        filtra_especial = FRM_filtra_especial
        Tem_Filtro = FRM_Tem_Filtro
        ConTelaOrdem = FRM_ConTelaOrdem
        Quantid = FRM_Quantid
        Filtrou = FRM_Filtrou
        BarraFerramentasPressionada = FRM_BarraFerramentasPressionada
        VariasVezes = FRM_VariasVezes
        XInserir = FRM_XInserir
        XOrigem = FRM_XOrigem
        SQLAtivo = FRM_SQLAtivo

        ReDim Filtros(UBound(FRM_Filtros, 1), UBound(FRM_Filtros, 2))
        ReDim Ordens(UBound(FRM_Ordens, 1))
        
        For Count = LBound(FRM_Filtros, 1) To UBound(FRM_Filtros, 1)
            For Count2 = LBound(FRM_Filtros, 2) To UBound(FRM_Filtros, 2)
                Filtros(Count, Count2) = FRM_Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(FRM_Ordens) To UBound(FRM_Ordens)
            Ordens(Count) = FRM_Ordens(Count)
        Next Count
    End If
End Sub

Private Sub subHabilitaBotoes()
    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       Botao_Fil.Enabled = False
    Else
       Botao_Fil.Enabled = True
    End If
End Sub
Private Sub subDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    Botao_Fil.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub

Private Sub Botao_Fil_Click()
    Dim Fil As Integer
    subDesabilitaBotoes
    
    Set Formulario = TelaCnae
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subContaReg Adodc1
        'Adodc1.Refresh TestesubrecarregadadosNV
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
    End If

    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.SetFocus
End Sub

Private Sub Botao_Fil_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub
Private Sub CboClasse_LostFocus()
    DatClasse.Recordset.bookmark = CboClasse.SelectedItem
End Sub

Private Sub CboDivisao_LostFocus()
    DatDivisao.Recordset.bookmark = CboDivisao.SelectedItem
End Sub

Private Sub CboGrupo_LostFocus()
    DatGrupo.Recordset.bookmark = CboGrupo.SelectedItem
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If

End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If

End Sub


Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If

End Sub
Private Sub cboSecao_LostFocus()
    DatSecao.Recordset.bookmark = cboSecao.SelectedItem
End Sub

Private Sub CmdAlterar_Click()
  
  Dim XLO_ALTERAR As ADODB.Recordset
  
  If Not FunVerifica_Permissao("TelaIdentificadores", "3") Then
     Exit Sub
  End If
  
  TelaCnae.MousePointer = vbHourglass
  PanBotoes.Enabled = False
  
  Select Case XLT_TELA

        Case "S" 'Seção
            
            PanSecao.Enabled = True
            PanSecao.Tag = "A"
            PanSecao.Top = (TDBGrid1.Height - PanSecao.Height) / 2
            PanSecao.Left = (TDBGrid1.Width - PanSecao.Width) / 2
            PanSecao.Caption = " Alteração de Seção"
            PanSecao.Visible = True
            TxtCodigoSecao.Text = Adodc1.Recordset.Fields("cnse_tx_Codigo").Value
            txtDescricaoSecao.Text = Adodc1.Recordset.Fields("cnse_tx_Descricao").Value
            TxtCodigoSecao.SetFocus
            
        Case "D" 'Divisão
        
            PanDivisao.Enabled = True
            PanDivisao.Tag = "A"
            PanDivisao.Top = (TDBGrid1.Height - PanDivisao.Height) / 2
            PanDivisao.Left = (TDBGrid1.Width - PanDivisao.Width) / 2
            PanDivisao.Caption = " Alteração de Divisão"
            PanDivisao.Visible = True
            txtCodigoDivisao.Text = Adodc1.Recordset.Fields("cndi_tx_Codigo").Value
            txtDescricaoDivisao.Text = Adodc1.Recordset.Fields("cndi_tx_Descricao").Value
            
            SubQOpenRecordset XLO_ALTERAR, _
                "SELECT cnse_tx_codigo + ' - ' + cnse_tx_descricao as cnse_tx_codigo FROM CnaeSecao WHERE cnse_cd_CnaeSecao = '" & _
                Adodc1.Recordset.Fields("cnse_cd_cnaesecao").Value & "'", Dinamico
            If Not XLO_ALTERAR.EOF Then
                cboSecao.Text = XLO_ALTERAR!cnse_tx_codigo
            End If
            
            XLO_ALTERAR.Close
            
            txtCodigoDivisao.SetFocus
            
        Case "G" 'Grupo
            
            PanGrupo.Enabled = True
            PanGrupo.Tag = "A"
            PanGrupo.Top = (TDBGrid1.Height - PanGrupo.Height) / 2
            PanGrupo.Left = (TDBGrid1.Width - PanGrupo.Width) / 2
            PanGrupo.Caption = " Alteração de Grupo"
            PanGrupo.Visible = True
            txtCodigoGrupo.Text = Adodc1.Recordset.Fields("cngr_tx_Codigo").Value
            txtDescricaoGrupo.Text = Adodc1.Recordset.Fields("cngr_tx_Descricao").Value
            txtCodigoGrupo.SetFocus
            
            SubQOpenRecordset XLO_ALTERAR, _
                "SELECT cndi_tx_codigo + ' - ' + cndi_tx_descricao as cndi_tx_codigo FROM CnaeDivisao WHERE cndi_cd_Cnaedivisao= '" & _
            Adodc1.Recordset.Fields("cndi_cd_cnaedivisao").Value & "'", Dinamico
            
            If Not XLO_ALTERAR.EOF Then
                CboDivisao.Text = XLO_ALTERAR!cndi_tx_codigo
            End If
            
            XLO_ALTERAR.Close
            
            
        Case "C" 'Classe
    
            panClasse.Enabled = True
            panClasse.Tag = "A"
            panClasse.Top = (TDBGrid1.Height - panClasse.Height) / 2
            panClasse.Left = (TDBGrid1.Width - panClasse.Width) / 2
            panClasse.Caption = " Alteração de Classe"
            panClasse.Visible = True
            txtCodigoClasse.Mask = ""
            txtCodigoClasse.Text = Adodc1.Recordset.Fields("cncl_tx_Codigo").Value
            txtCodigoClasse.Mask = "9999-9"
            txtDescricaoClasse.Text = Adodc1.Recordset.Fields("cncl_tx_Descricao").Value
            txtCodigoClasse.SetFocus
            
            SubQOpenRecordset XLO_ALTERAR, _
            "SELECT cngr_tx_codigo + ' - ' + cngr_tx_descricao as cngr_tx_codigo FROM CnaeGrupo WHERE cngr_cd_Cnaegrupo = '" & _
            Adodc1.Recordset.Fields("cngr_cd_cnaegrupo").Value & "'", Dinamico
            
            If Not XLO_ALTERAR.EOF Then
                CboGrupo.Text = XLO_ALTERAR!cngr_tx_codigo
            End If
            
            XLO_ALTERAR.Close
            
        Case "SC" ' Sub-Classe
        
            panSubClasse.Enabled = True
            panSubClasse.Tag = "A"
            panSubClasse.Top = (TDBGrid1.Height - panSubClasse.Height) / 2
            panSubClasse.Left = (TDBGrid1.Width - panSubClasse.Width) / 2
            panSubClasse.Caption = " Alteração de Sub-Classe"
            panSubClasse.Visible = True
            txtCodigoSubclasse.Text = Adodc1.Recordset.Fields("cnsc_tx_Codigo").Value
            txtDescricaoSubclasse.Text = Adodc1.Recordset.Fields("cnsc_tx_Descricao").Value
            txtCodigoSubclasse.SetFocus
            
            SubQOpenRecordset XLO_ALTERAR, _
                "SELECT cncl_tx_codigo + ' - ' + cncl_tx_descricao as cncl_tx_codigo FROM CnaeClasse WHERE cncl_cd_Cnaeclasse = '" & _
            Adodc1.Recordset.Fields("cncl_cd_cnaeclasse").Value & "'", Dinamico
            
            If Not XLO_ALTERAR.EOF Then
                CboClasse.Text = XLO_ALTERAR!cncl_tx_codigo
            End If
            
            XLO_ALTERAR.Close
            
    End Select
    
  
  subDesabilitaBotoes
  cboLocalizar.Enabled = False
  'TDBGrid1.Enabled = False
  TelaCnae.MousePointer = vbDefault
  
End Sub
'
'Private Sub CmdCancelar_Click()
'  If Not FunTabelaVazia(Adodc1) Then Adodc1.Recordset.CancelUpdate
'  PanBotoes.Enabled = True
'  PanTipo.Visible = False
'  subHabilitaBotoes
'  CmdConfirmar.Caption = "Confirmar"
'  cboLocalizar.Enabled = True
'  TDBGrid1.Enabled = True
'  TDBGrid1.SetFocus
'End Sub

Private Sub cmdConfirmarClasse_Click()
    
'    'Validação da Máscara
    If Len(Trim(funTiraAcento(txtCodigoClasse, True))) <> 6 Then
        MsgBox "Máscara de Classe Inválida!", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

      'verificar preenchimento dos campos obrigatórios
    If txtCodigoClasse.Text = "" Then
      MsgBox "O preenchimento do código da classe é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
      Exit Sub
    End If
    
    If txtDescricaoClasse.Text = "" Then
        MsgBox "O preenchimento da descrição da classe é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    If CboGrupo.Text = "" Then
        MsgBox "O preenchimento do grupo é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
       
      'verifica se já existe este identificador
    If panClasse.Tag = "I" Then
        SubQOpenRecordset ResExiste, _
        "SELECT * FROM CnaeClasse WHERE cncl_tx_codigo = '" & _
        txtCodigoClasse.Text & "'", Dinamico
        
        If Not ResExiste.EOF Then
            MsgBox "Esta classe já foi cadastrada. Tente outra.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCodigoClasse.SetFocus
            ResExiste.Close
            Exit Sub
        End If
        
        ResExiste.Close
        Set ResExiste = Nothing
      End If
    'Trim$(Replace(FunNulo(txtCodigoClasse.Text), "-", ""))
      'inicia a gravação / alteração
    Conexao.BeginTrans
    If panClasse.Tag = "I" Then
        Sql = _
        "INSERT INTO CnaeClasse VALUES (" & _
        "'" & FunNulo(txtCodigoClasse.Text) & "', '" & _
        txtDescricaoClasse.Text & "', " & DatGrupo.Recordset.Fields("cngr_cd_CnaeGrupo") & ")"
        Conexao.Execute Sql
        MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) & _
        "/" & CStr((Adodc1.Recordset.RecordCount + 1))
    Else 'Alteração
     
        Sql = _
         "UPDATE CnaeClasse SET " & _
         "cncl_tx_codigo = '" & FunNulo(txtCodigoClasse.Text) & "', " & _
         "cncl_tx_descricao = '" & FunNulo(txtDescricaoClasse.Text) & "', " & _
         "cngr_cd_cnaegrupo = " & DatGrupo.Recordset.Fields("cngr_cd_CnaeGrupo") & "" & _
         " WHERE cncl_cd_cnaeclasse = " & Adodc1.Recordset.Fields("cncl_cd_cnaeclasse")
         
    Conexao.Execute Sql
      
    End If
    
    Conexao.CommitTrans
       
    panClasse.Visible = False
    panClasse.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
       
    cboLocalizar.Enabled = True
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
       
    If panClasse.Tag <> "I" Then
        cboLocalizar.BoundText = panClasse.Tag
    '   Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    Exit Sub
       
MostraErro:
      ' Enumera a coleção de erros e apresenta
      ' as propriedades de cada erro ocorrido
    For Each errLoop In Conexao.Errors
         StrError = "Error #" & errLoop.Number & vbCr & _
            "   " & errLoop.Description & vbCr & _
            "   (Source: " & errLoop.Source & ")" & vbCr & _
            "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
            "   (NativeError: " & errLoop.NativeError & ")" & vbCr
    
    MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
    Next
    
    Resume Next

    
End Sub

Private Sub CmdConfirmarDivisao_Click()
    
    'Validação da Máscara
    If Len(Trim$(txtCodigoDivisao)) <> 2 Then
        MsgBox "Máscara de Divisão Inválida!", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

'verificar preenchimento dos campos obrigatórios
  If txtCodigoDivisao.Text = "" Then
    MsgBox "O preenchimento do código da Divisao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If txtDescricaoDivisao.Text = "" Then
    MsgBox "O preenchimento da descrição da Divisao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If cboSecao.Text = "" Then
    MsgBox "O preenchimento da Secao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
   
  'verifica se já existe esta divisao
  If PanDivisao.Tag = "I" Then
    SubQOpenRecordset ResExiste, _
      "SELECT * FROM CnaeDivisao WHERE cndi_tx_codigo = '" & _
      txtCodigoDivisao.Text & "'", Dinamico
    If Not ResExiste.EOF Then
      MsgBox "Esta Divisao já foi cadastrada. Tente outra.", vbCritical + vbOKOnly, "ATENÇÃO"
      txtCodigoDivisao.SetFocus
      ResExiste.Close
      Exit Sub
    End If
    ResExiste.Close
    Set ResExiste = Nothing
  End If

  'inicia a gravação / alteração
  Conexao.BeginTrans
  If PanDivisao.Tag = "I" Then
    Sql = _
      "INSERT INTO CnaeDivisao VALUES (" & _
      "'" & FunNulo(txtCodigoDivisao.Text) & "', '" & _
      txtDescricaoDivisao.Text & "', " & DatSecao.Recordset.Fields("cnse_cd_CnaeSecao") & ")"
    Conexao.Execute Sql
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) & _
      "/" & CStr((Adodc1.Recordset.RecordCount + 1))
  Else    'Alteração
    Sql = _
      "UPDATE CnaeDivisao SET " & _
      "cndi_tx_codigo = '" & FunNulo(txtCodigoDivisao.Text) & "', " & _
      "cndi_tx_descricao = '" & FunNulo(txtDescricaoDivisao.Text) & "', " & _
      "cnse_cd_CnaeSecao = " & DatSecao.Recordset.Fields("cnse_cd_CnaeSecao") & "" & _
      " WHERE cndi_cd_cnaeDivisao = " & Adodc1.Recordset.Fields("cndi_cd_cnaeDivisao")
    Conexao.Execute Sql
  End If
  Conexao.CommitTrans
   
  PanDivisao.Visible = False
  PanDivisao.Enabled = False
  PanBotoes.Enabled = True
  subHabilitaBotoes
   
  cboLocalizar.Enabled = True
  subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   
  If PanDivisao.Tag <> "I" Then
    cboLocalizar.BoundText = PanDivisao.Tag
'    Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
  End If
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  Exit Sub

MostraErro:
  ' Enumera a coleção de erros e apresenta
  ' as propriedades de cada erro ocorrido
  For Each errLoop In Conexao.Errors
     StrError = "Error #" & errLoop.Number & vbCr & _
        "   " & errLoop.Description & vbCr & _
        "   (Source: " & errLoop.Source & ")" & vbCr & _
        "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
        "   (NativeError: " & errLoop.NativeError & ")" & vbCr

  MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
  Next

  Resume Next


End Sub

Private Sub cmdConfirmarGrupo_Click()
 
 'Validação da Máscara
    If Len(Trim$(txtCodigoGrupo)) <> 3 Then
        MsgBox "Máscara do Grupo Inválida!", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

'verificar preenchimento dos campos obrigatórios
  If txtCodigoGrupo.Text = "" Then
    MsgBox "O preenchimento do código do Grupo é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If txtDescricaoGrupo.Text = "" Then
    MsgBox "O preenchimento da descrição do Grupo é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If CboDivisao.Text = "" Then
    MsgBox "O preenchimento da Divisao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
   
  'verifica se já existe este Grupo
  If PanGrupo.Tag = "I" Then
    SubQOpenRecordset ResExiste, _
      "SELECT * FROM CnaeGrupo WHERE cngr_tx_codigo = '" & _
      txtCodigoGrupo.Text & "'", Dinamico
    If Not ResExiste.EOF Then
      MsgBox "Este Grupo já foi cadastrado. Tente outro.", vbCritical + vbOKOnly, "ATENÇÃO"
      txtCodigoGrupo.SetFocus
      ResExiste.Close
      Exit Sub
    End If
    ResExiste.Close
    Set ResExiste = Nothing
  End If

  'inicia a gravação / alteração
  Conexao.BeginTrans
  If PanGrupo.Tag = "I" Then
    Sql = _
      "INSERT INTO CnaeGrupo VALUES (" & _
      "'" & FunNulo(txtCodigoGrupo.Text) & "', '" & _
      txtDescricaoGrupo.Text & "', " & DatDivisao.Recordset.Fields("cndi_cd_CnaeDivisao") & ")"
    Conexao.Execute Sql
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) & _
      "/" & CStr((Adodc1.Recordset.RecordCount + 1))
  Else    'Alteração
    Sql = _
      "UPDATE CnaeGrupo SET " & _
      "cngr_tx_codigo = '" & FunNulo(txtCodigoGrupo.Text) & "', " & _
      "cngr_tx_descricao = '" & FunNulo(txtDescricaoGrupo.Text) & "', " & _
      "cndi_cd_CnaeDivisao = " & DatDivisao.Recordset.Fields("cndi_cd_CnaeDivisao") & "" & _
      " WHERE cngr_cd_cnaeGrupo = " & Adodc1.Recordset.Fields("cngr_cd_cnaeGrupo")
    Conexao.Execute Sql
  End If
  Conexao.CommitTrans
   
  PanGrupo.Visible = False
  PanGrupo.Enabled = False
  PanBotoes.Enabled = True
  subHabilitaBotoes
   
  cboLocalizar.Enabled = True
  subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   
  If PanGrupo.Tag <> "I" Then
    cboLocalizar.BoundText = PanGrupo.Tag
'    Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
  End If
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  Exit Sub

MostraErro:
  ' Enumera a coleção de erros e apresenta
  ' as propriedades de cada erro ocorrido
  For Each errLoop In Conexao.Errors
     StrError = "Error #" & errLoop.Number & vbCr & _
        "   " & errLoop.Description & vbCr & _
        "   (Source: " & errLoop.Source & ")" & vbCr & _
        "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
        "   (NativeError: " & errLoop.NativeError & ")" & vbCr

  MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
  Next

  Resume Next

End Sub


Private Sub CmdConfirmarSecao_Click()
      
'verificar preenchimento dos campos obrigatórios
  If TxtCodigoSecao.Text = "" Then
    MsgBox "O preenchimento do código da Secao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If txtDescricaoSecao.Text = "" Then
    MsgBox "O preenchimento da descrição da Secao é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If

   
  'verifica se já existe esta Secao
  If PanSecao.Tag = "I" Then
    SubQOpenRecordset ResExiste, _
      "SELECT * FROM CnaeSecao WHERE cnse_tx_codigo = '" & _
      TxtCodigoSecao.Text & "'", Dinamico
    If Not ResExiste.EOF Then
      MsgBox "Esta Secao já foi cadastrada. Tente outra.", vbCritical + vbOKOnly, "ATENÇÃO"
      TxtCodigoSecao.SetFocus
      ResExiste.Close
      Exit Sub
    End If
    ResExiste.Close
    Set ResExiste = Nothing
  End If

  'inicia a gravação / alteração
  Conexao.BeginTrans
  
  If PanSecao.Tag = "I" Then
    
    Sql = _
      "INSERT INTO CnaeSecao VALUES (" & _
      "'" & FunNulo(TxtCodigoSecao.Text) & "', '" & _
      txtDescricaoSecao.Text & "'" & ")"
    Conexao.Execute Sql
    
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) & _
      "/" & CStr((Adodc1.Recordset.RecordCount + 1))
      
  Else    'Alteração
    
    Sql = _
      "UPDATE CnaeSecao SET " & _
      "cnse_tx_codigo = '" & FunNulo(TxtCodigoSecao.Text) & "', " & _
      "cnse_tx_descricao = '" & FunNulo(txtDescricaoSecao.Text) & "'" & _
      " WHERE cnse_cd_cnaeSecao = " & Adodc1.Recordset.Fields("cnse_cd_cnaeSecao")
    Conexao.Execute Sql
    
  End If
  
  Conexao.CommitTrans
   
  PanSecao.Visible = False
  PanSecao.Enabled = False
  PanBotoes.Enabled = True
  subHabilitaBotoes
   
  cboLocalizar.Enabled = True
  subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   
  If PanSecao.Tag <> "I" Then
    cboLocalizar.BoundText = PanSecao.Tag
  '  Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
  End If
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  Exit Sub

MostraErro:
  ' Enumera a coleção de erros e apresenta
  ' as propriedades de cada erro ocorrido
  For Each errLoop In Conexao.Errors
     StrError = "Error #" & errLoop.Number & vbCr & _
        "   " & errLoop.Description & vbCr & _
        "   (Source: " & errLoop.Source & ")" & vbCr & _
        "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
        "   (NativeError: " & errLoop.NativeError & ")" & vbCr

  MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
  Next

  Resume Next

End Sub

Private Sub cmdConfirmarSubClasse_Click()

      'Validação da Máscara
    If Len(Trim(funTiraAcento(txtCodigoSubclasse, True))) <> 9 Then
        MsgBox "Máscara de Sub-Classe Inválida!", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

'verificar preenchimento dos campos obrigatórios
  If txtCodigoSubclasse.Text = "" Then
    MsgBox "O preenchimento do código da SubClasse é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If txtDescricaoSubclasse.Text = "" Then
    MsgBox "O preenchimento da descrição da SubClasse é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
  If CboClasse.Text = "" Then
    MsgBox "O preenchimento da Classe é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Sub
  End If
   
  'verifica se já existe este identificador
  If panSubClasse.Tag = "I" Then
    SubQOpenRecordset ResExiste, _
      "SELECT * FROM CnaeSubClasse WHERE cnsc_tx_codigo = '" & _
      txtCodigoSubclasse.Text & "'", Dinamico
    If Not ResExiste.EOF Then
      MsgBox "Esta SubClasse já foi cadastrada. Tente outra.", vbCritical + vbOKOnly, "ATENÇÃO"
      txtCodigoSubclasse.SetFocus
      ResExiste.Close
      Exit Sub
    End If
    ResExiste.Close
    Set ResExiste = Nothing
  End If

  'inicia a gravação / alteração
  Conexao.BeginTrans
  If panSubClasse.Tag = "I" Then
    Sql = _
      "INSERT INTO CnaeSubClasse VALUES (" & _
      "'" & FunNulo(txtCodigoSubclasse.Text) & "', '" & _
      txtDescricaoSubclasse.Text & "', " & DatClasse.Recordset.Fields("cncl_cd_CnaeClasse") & ")"
      Conexao.Execute Sql
      MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) & _
      "/" & CStr((Adodc1.Recordset.RecordCount + 1))
  Else    'Alteração
    Sql = _
      "UPDATE CnaeSubClasse SET " & _
      "cnsc_tx_codigo = '" & FunNulo(txtCodigoSubclasse.Text) & "', " & _
      "cnsc_tx_descricao = '" & FunNulo(txtDescricaoSubclasse.Text) & "', " & _
      "cncl_cd_CnaeClasse = " & DatClasse.Recordset.Fields("cncl_cd_CnaeClasse") & "" & _
      " WHERE cnsc_cd_cnaeSubClasse = " & Adodc1.Recordset.Fields("cnsc_cd_cnaeSubClasse")
    Conexao.Execute Sql
  End If
  Conexao.CommitTrans
   
  panSubClasse.Visible = False
  panSubClasse.Enabled = False
  PanBotoes.Enabled = True
  subHabilitaBotoes
   
  cboLocalizar.Enabled = True
  subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   
  If panSubClasse.Tag <> "I" Then
    cboLocalizar.BoundText = panSubClasse.Tag
'    Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
  End If
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  Exit Sub

MostraErro:
  ' Enumera a coleção de erros e apresenta
  ' as propriedades de cada erro ocorrido
  For Each errLoop In Conexao.Errors
     StrError = "Error #" & errLoop.Number & vbCr & _
        "   " & errLoop.Description & vbCr & _
        "   (Source: " & errLoop.Source & ")" & vbCr & _
        "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
        "   (NativeError: " & errLoop.NativeError & ")" & vbCr

  MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
  Next

  Resume Next

    

End Sub

Private Sub cmdExport_Click()
    Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    TDBGrid1.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To TDBGrid1.Columns.Count - 1
         If i = 0 And TDBGrid1.Columns(i).Visible Then
             sLine = Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         ElseIf TDBGrid1.Columns(i).Visible Then
             sLine = sLine & ";" & Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & Adodc1.Recordset.RecordCount & " registros", vbInformation, "Informação"
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Identificadores")
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub cmdRetornarClasse_Click()
    
    panClasse.Visible = False
    TDBGrid1.Enabled = True
    PanBotoes.Enabled = True

    
End Sub

Private Sub cmdRetornarDivisao_Click()
    
    PanDivisao.Visible = False
    TDBGrid1.Enabled = True
    PanBotoes.Enabled = True

End Sub

Private Sub cmdRetornarGrupo_Click()
    
    PanGrupo.Visible = False
    TDBGrid1.Enabled = True
    PanBotoes.Enabled = True

    
End Sub

Private Sub CmdRetornarSecao_Click()
    
    PanSecao.Visible = False
    TDBGrid1.Enabled = True
    PanBotoes.Enabled = True

    
End Sub

Private Sub cmdRetornarSubclasse_Click()
    
    panSubClasse.Visible = False
    TDBGrid1.Enabled = True
    PanBotoes.Enabled = True

    
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    
    Set Formulario = TelaCnae
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.SetFocus
    
End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    On Error GoTo RotuloErro
    Dim opcao As Integer
    Dim XLT_MENSAGEM As String
    Dim Posicao As Variant

    If Not FunVerifica_Permissao("TelaIdentificadores", "4") Then
        Exit Sub
    End If
    
    Posicao = Adodc1.Recordset.bookmark
    subDesabilitaBotoes
    
    opcao = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If opcao = 1 Then  'Confirma exclusão
      
      Conexao.BeginTrans
      
       Select Case XLT_TELA

        Case "S" 'Seção
        
            XLT_MENSAGEM = "Existem divisões associadas a esta seção portanto ela não pode ser excluída!"
            Conexao.Execute "DELETE FROM CnaeSecao WHERE cnse_cd_cnaesecao = " & Adodc1.Recordset.Fields("cnse_cd_cnaesecao")
            
        Case "D" 'Divisão
        
            XLT_MENSAGEM = "Existem grupos associados a esta divisão portanto ela não pode ser excluída!"
            Conexao.Execute "DELETE FROM CnaeDivisao WHERE cndi_cd_cnaedivisao = " & Adodc1.Recordset.Fields("cndi_cd_cnaedivisao")
        
        Case "G" 'Grupo
        
            XLT_MENSAGEM = "Existem classes associadas a este grupo portanto ele não pode ser excluído!"
            Conexao.Execute "DELETE FROM CnaeGrupo WHERE cngr_cd_cnaegrupo = " & Adodc1.Recordset.Fields("cngr_cd_cnaegrupo")
            
        Case "C" 'Classe
        
            XLT_MENSAGEM = "Existem subclasses associadas a esta classe portanto ela não pode ser excluída!"
            Conexao.Execute "DELETE FROM CnaeClasse WHERE cncl_cd_cnaeclasse = " & Adodc1.Recordset.Fields("cncl_cd_cnaeclasse")
         
        Case "SC" 'SubClasse
        
            Conexao.Execute "DELETE FROM CnaesubClasse WHERE cnsc_cd_cnaesubclasse = " & Adodc1.Recordset.Fields("cnsc_cd_cnaesubclasse")
        
      End Select

      Conexao.CommitTrans
       
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      
      MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
      
      If ContFil <> 0 Then
         subContaReg Adodc1
      End If
      
      subHabilitaBotoes
      
      If FunTabelaVazia(Adodc1) = True Then
         CmdAlterar.Enabled = False
         CmdExcluir.Enabled = False
         CmdImprimir.Enabled = False
         CmdOrdem.Enabled = False
      End If
       
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If
    
    Exit Sub
    
RotuloErro:
   Conexao.RollbackTrans
    If Err.Number = -2147467259 Then
      MsgBox "Esse item não pode ser excluído pois já possui algum registro relacionado!", vbCritical, "ATENÇÃO"
    ElseIf Err.Number = "-2147217873" Then
        MsgBox XLT_MENSAGEM, vbCritical, "ATENÇÃO"
    Else
        MsgBox "Não foi possível efetuar a exclusão!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
    
    If Not FunVerifica_Permissao("TelaIdentificadores", "2") Then
        Exit Sub
    End If
    
    TelaCnae.MousePointer = vbHourglass
       
    Select Case XLT_TELA
            
        Case "S" 'Seção
        
            PanSecao.Enabled = True
            PanSecao.Tag = "I"
            PanSecao.Top = (TDBGrid1.Height - PanSecao.Height) / 2
            PanSecao.Left = (TDBGrid1.Width - PanSecao.Width) / 2
            PanSecao.Caption = " Inserção de Seção"
            PanSecao.Visible = True
            TxtCodigoSecao.Text = ""
            txtDescricaoSecao.Text = ""
            TxtCodigoSecao.SetFocus
            
        Case "D" 'Divisão
        
            PanDivisao.Enabled = True
            PanDivisao.Tag = "I"
            PanDivisao.Top = (TDBGrid1.Height - PanDivisao.Height) / 2
            PanDivisao.Left = (TDBGrid1.Width - PanDivisao.Width) / 2
            PanDivisao.Caption = " Inserção de Divisão"
            PanDivisao.Visible = True
            txtCodigoDivisao = ""
            txtDescricaoDivisao.Text = ""
            cboSecao.Text = ""
            txtCodigoDivisao.SetFocus
            
            
        Case "G" 'Grupo
            
            PanGrupo.Enabled = True
            PanGrupo.Tag = "I"
            PanGrupo.Top = (TDBGrid1.Height - PanGrupo.Height) / 2
            PanGrupo.Left = (TDBGrid1.Width - PanGrupo.Width) / 2
            PanGrupo.Caption = " Inserção de Grupo"
            PanGrupo.Visible = True
            txtCodigoGrupo.Text = ""
            txtDescricaoGrupo.Text = ""
            CboDivisao.Text = ""
            txtCodigoGrupo.SetFocus
            
        Case "C" 'Classe
    
            panClasse.Enabled = True
            panClasse.Tag = "I"
            panClasse.Top = (TDBGrid1.Height - panClasse.Height) / 2
            panClasse.Left = (TDBGrid1.Width - panClasse.Width) / 2
            panClasse.Caption = " Inserção de Classe"
            panClasse.Visible = True
            
            txtCodigoClasse.Mask = ""
            txtCodigoClasse = ""
            txtCodigoClasse.Mask = "9999-9"
            
            txtDescricaoClasse.Text = ""
            CboGrupo.Text = ""
            txtCodigoClasse.SetFocus
            
        Case "SC" ' Sub-Classe
        
            panSubClasse.Enabled = True
            panSubClasse.Tag = "I"
            panSubClasse.Top = (TDBGrid1.Height - panSubClasse.Height) / 2
            panSubClasse.Left = (TDBGrid1.Width - panSubClasse.Width) / 2
            panSubClasse.Caption = " Inserção de Sub-Classe"
            panSubClasse.Visible = True
        
            txtCodigoSubclasse.Mask = ""
            txtCodigoSubclasse = ""
            txtCodigoSubclasse.Mask = "9999-9/99"
            
            txtDescricaoSubclasse.Text = ""
            CboClasse.Text = ""
            txtCodigoSubclasse.SetFocus
    
    End Select
        
    TDBGrid1.Enabled = False
    PanBotoes.Enabled = False
    subDesabilitaBotoes
    cboLocalizar.BoundText = ""
    cboLocalizar.Enabled = False
    TelaCnae.MousePointer = vbDefault

    
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Grava na tabela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    End If
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaCnae
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False 'Desabilita o botão Excluir
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
   
End Sub

Private Sub Form_Activate()
  Set Formulario = TelaCnae
  If Not PrimeiraVez Then
      subTelaValoresGlobais "P"
  End If
  
  PrimeiraVez = False
  
  MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaCnae
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaCnae"
    
    If XGT_CNAE = "Secao" Then
        'Associa um banco de dados à tabela CnaeSecao
        subCarregaDadosNV Adodc1, "", "cnse_tx_Codigo", "", "SELECT * FROM CnaeSecao"
        TelaCnae.Caption = "Tabela de Seção do Cnae"
        XLT_TELA = "S"
        
    ElseIf XGT_CNAE = "Divisao" Then
        'Associa um banco de dados à tabela CnaeDivisao
        subCarregaDadosNV Adodc1, "", "cndi_tx_Codigo", "", "SELECT * FROM CnaeDivisao"
        TelaCnae.Caption = "Tabela de Divisão do Cnae"
        XLT_TELA = "D"
        
        subConectarControleDadosNV DatSecao, "SELECT cnse_cd_cnaesecao, cnse_tx_codigo + ' - ' + cnse_tx_descricao as cnse_tx_codigo FROM CnaeSecao ORDER BY cnse_tx_Codigo", Estatico
        
        TDBGrid1.Columns(0).Visible = False
        TDBGrid1.Columns(1).Visible = False
        TDBGrid1.Columns(2).Visible = True
        TDBGrid1.Columns(3).Visible = True
        
    ElseIf XGT_CNAE = "Grupo" Then
        'Associa um banco de dados à tabela CnaeGrupo
        subCarregaDadosNV Adodc1, "", "cngr_tx_Codigo", "", "SELECT * FROM CnaeGrupo"
        TelaCnae.Caption = "Tabela de Grupo do Cnae"
        XLT_TELA = "G"
        
        'subConectarControleDadosNV DatDivisao, "SELECT * FROM CnaeDivisao ORDER BY cndi_tx_Codigo", Estatico
        
        subConectarControleDadosNV DatDivisao, "SELECT cndi_cd_cnaedivisao, cndi_tx_Codigo + ' - ' + cndi_tx_descricao as cndi_tx_Codigo FROM CnaeDivisao ORDER BY cndi_tx_Codigo", Estatico
        
        TDBGrid1.Columns(0).Visible = False
        TDBGrid1.Columns(1).Visible = False
        TDBGrid1.Columns(4).Visible = True
        TDBGrid1.Columns(5).Visible = True
        
    ElseIf XGT_CNAE = "Classe" Then
        'Associa um banco de dados à tabela CnaeClasse
        subCarregaDadosNV Adodc1, "", "cncl_tx_Codigo", "", "SELECT * FROM CnaeClasse"
        TelaCnae.Caption = "Tabela de Classe do Cnae"
        XLT_TELA = "C"
        
        'subConectarControleDadosNV DatGrupo, "SELECT * FROM CnaeGrupo ORDER BY cngr_tx_Codigo", Estatico
        subConectarControleDadosNV DatGrupo, "SELECT cngr_cd_cnaegrupo, cngr_tx_Codigo + ' - ' + cngr_tx_descricao as cngr_tx_Codigo FROM CnaeGrupo ORDER BY cngr_tx_Codigo", Estatico
        
        
        TDBGrid1.Columns(0).Visible = False
        TDBGrid1.Columns(1).Visible = False
        TDBGrid1.Columns(6).Visible = True
        TDBGrid1.Columns(7).Visible = True
        
    ElseIf XGT_CNAE = "SubClasse" Then
        'Associa um banco de dados à tabela CnaeSubClasse
        subCarregaDadosNV Adodc1, "", "cnsc_tx_Codigo", "", "SELECT * FROM CnaeSubClasse"
        TelaCnae.Caption = "Tabela de Sub-Classe do Cnae"
        XLT_TELA = "SC"
        
        'subConectarControleDadosNV DatClasse, "SELECT * FROM CnaeClasse ORDER BY cncl_tx_Codigo", Estatico
         subConectarControleDadosNV DatClasse, "SELECT cncl_cd_cnaeclasse, cncl_tx_Codigo + ' - ' + cncl_tx_descricao as cncl_tx_Codigo FROM CnaeClasse ORDER BY cncl_tx_Codigo", Estatico
        
        
        TDBGrid1.Columns(0).Visible = False
        TDBGrid1.Columns(1).Visible = False
        
        TDBGrid1.Columns(8).Visible = True
        TDBGrid1.Columns(9).Visible = True
        
    End If
    
    subHabilitaBotoes
    
    filtra_especial = False
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaCnae"
End Sub

Private Sub txtCodigoClasse_Change()
    Call subSelecionaMSK(txtCodigoClasse)
End Sub

Private Sub txtCodigoSubclasse_Change()
   Call subSelecionaMSK(txtCodigoSubclasse)
End Sub
