VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "TDBG7.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNotasFiscais 
   Caption         =   "Cadastro de Contas a Pagar"
   ClientHeight    =   5295
   ClientLeft      =   75
   ClientTop       =   1350
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   9375
   Begin Threed.SSPanel PanEntMultiplas 
      Height          =   1590
      Left            =   4425
      TabIndex        =   118
      Top             =   3300
      Visible         =   0   'False
      Width           =   3105
      _Version        =   65536
      _ExtentX        =   5477
      _ExtentY        =   2805
      _StockProps     =   15
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Enabled         =   0   'False
      Begin VB.CommandButton CmdDesiste 
         BackColor       =   &H00000000&
         Caption         =   "&Desiste"
         Height          =   240
         Left            =   390
         TabIndex        =   123
         Top             =   1260
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox TxtNParcelas 
         Height          =   315
         Left            =   1545
         TabIndex        =   119
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdSairEntMult 
         BackColor       =   &H00000000&
         Caption         =   "&Retornar"
         Height          =   240
         Left            =   2160
         TabIndex        =   125
         Top             =   1260
         Width           =   840
      End
      Begin VB.CommandButton CmdConfEntMult 
         BackColor       =   &H00000000&
         Caption         =   "&Confirmar"
         Height          =   240
         Left            =   1275
         TabIndex        =   124
         Top             =   1260
         Width           =   840
      End
      Begin VB.TextBox TxtPeriodicidade 
         Height          =   315
         Left            =   1545
         TabIndex        =   120
         Top             =   720
         Width           =   630
      End
      Begin VB.OptionButton OptDia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dia"
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
         Left            =   2235
         TabIndex        =   121
         Top             =   675
         Width           =   690
      End
      Begin VB.OptionButton OptMes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mês"
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
         Left            =   2235
         TabIndex        =   122
         Top             =   930
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.Label LblParcelas 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Parcelas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   375
         TabIndex        =   127
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label LblPeriodo 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodicidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   225
         TabIndex        =   126
         Top             =   735
         Width           =   1230
      End
   End
   Begin VB.CheckBox ChkDistribuicao 
      Caption         =   "Com &Distribuição"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   5000
      Width           =   1815
   End
   Begin TabDlg.SSTab TabNotasFiscais 
      Height          =   4785
      Left            =   0
      TabIndex        =   51
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8440
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "FrmNotasFiscais.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PanFornecedor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TabAuxiliar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSPanel1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PanNotasFiscais"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Composição de Duplicata"
      TabPicture(1)   =   "FrmNotasFiscais.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PnlCompDuplicata"
      Tab(1).Control(1)=   "CmdInsereduplicata"
      Tab(1).Control(2)=   "CmdAlteraduplicata"
      Tab(1).Control(3)=   "CmdExcluiduplicata"
      Tab(1).Control(4)=   "TDBGrid2"
      Tab(1).Control(5)=   "PnlDados1"
      Tab(1).Control(6)=   "PnlDados1_1"
      Tab(1).Control(7)=   "CmdDuplMultiplas"
      Tab(1).Control(8)=   "LblValorTotal"
      Tab(1).Control(9)=   "Label14"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Descontos"
      TabPicture(2)   =   "FrmNotasFiscais.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdAlteraDesconto"
      Tab(2).Control(1)=   "CmdExcluiDesconto"
      Tab(2).Control(2)=   "CmdInsereDesconto"
      Tab(2).Control(3)=   "pandescontos"
      Tab(2).Control(4)=   "TDBGrid1"
      Tab(2).Control(5)=   "PnlDados2"
      Tab(2).Control(6)=   "PnlDados2_2"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Distribuição"
      TabPicture(3)   =   "FrmNotasFiscais.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblValorNota"
      Tab(3).Control(1)=   "Label13"
      Tab(3).Control(2)=   "LblValorTotalDistrib"
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(4)=   "PnlDados3_3"
      Tab(3).Control(5)=   "PnlDados3"
      Tab(3).Control(6)=   "CmdInsereDistribuicao"
      Tab(3).Control(7)=   "CmdRemoverDistribuicao"
      Tab(3).Control(8)=   "CmdAlteraDistribuicao"
      Tab(3).Control(9)=   "TDBGrid3"
      Tab(3).Control(10)=   "PanDistribuicao"
      Tab(3).ControlCount=   11
      Begin Threed.SSPanel PanDistribuicao 
         Height          =   1845
         Left            =   -74835
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   9060
         _Version        =   65536
         _ExtentX        =   15981
         _ExtentY        =   3254
         _StockProps     =   15
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
         Begin VB.TextBox TxtPercDistribuicao 
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   44
            Top             =   1155
            Width           =   1260
         End
         Begin VB.TextBox TxtVldistribuicao 
            Height          =   330
            Left            =   4170
            TabIndex        =   45
            Top             =   1155
            Width           =   1695
         End
         Begin VB.CommandButton CmdSairDistribuicao 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   7695
            TabIndex        =   47
            Top             =   1380
            Width           =   1140
         End
         Begin VB.CommandButton Botao_Conf 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   6450
            TabIndex        =   46
            Top             =   1380
            Width           =   1140
         End
         Begin MSMask.MaskEdBox MskConta2 
            Height          =   330
            Left            =   1800
            TabIndex        =   42
            Top             =   705
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContas2 
            Bindings        =   "FrmNotasFiscais.frx":0070
            Height          =   330
            Left            =   2745
            TabIndex        =   43
            Top             =   705
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "plco_cd_codreduzido"
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
         Begin MSMask.MaskEdBox MskPcr2 
            Height          =   330
            Left            =   1800
            TabIndex        =   40
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr2 
            Bindings        =   "FrmNotasFiscais.frx":0087
            Height          =   330
            Left            =   2745
            TabIndex        =   41
            Top             =   255
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "cere_cd_estruturado"
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
         Begin VB.Label LblPercentual 
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   78
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label LblVlDist 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3225
            TabIndex        =   70
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label LblCR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Custo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   69
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label LblCC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Conta Contábil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   720
            Width           =   1350
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   3135
         Left            =   -74400
         TabIndex        =   79
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
         _LayoutType     =   4
         _RowHeight      =   12
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodPcr"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Centro Custo"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "CodConta"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Conta Contábil"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Devolução"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor Líquido"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Percentual"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "Percent"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1048848"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=3307"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3228"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1049104"
         Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(1)._MinWidth=33"
         Splits(0)._ColumnProps(18)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=1048848"
         Splits(0)._ColumnProps(23)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(26)=   "Column(2)._MinWidth=86574240"
         Splits(0)._ColumnProps(27)=   "Column(3).Width=3334"
         Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=3254"
         Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=1049104"
         Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(35)=   "Column(3)._MinWidth=86574240"
         Splits(0)._ColumnProps(36)=   "Column(4).Width=3043"
         Splits(0)._ColumnProps(37)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._WidthInPix=2963"
         Splits(0)._ColumnProps(39)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(4)._ColStyle=1049362"
         Splits(0)._ColumnProps(41)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(42)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(43)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(44)=   "Column(4)._MinWidth=86573588"
         Splits(0)._ColumnProps(45)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(5)._ColStyle=1049362"
         Splits(0)._ColumnProps(50)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(51)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(52)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(53)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(57)=   "Column(6)._ColStyle=1049362"
         Splits(0)._ColumnProps(58)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(59)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(60)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(61)=   "Column(7).Width=2408"
         Splits(0)._ColumnProps(62)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._WidthInPix=2328"
         Splits(0)._ColumnProps(64)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(7)._ColStyle=1049106"
         Splits(0)._ColumnProps(66)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(67)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(68)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=48,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=16,.parent=43"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=44"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=45"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=47"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=45"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=47"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=20,.parent=43"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=44"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=45"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=47"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=45"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=47"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=45"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=47"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=24,.parent=43,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=44,.alignment=1"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=45"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=47"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=45"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=47"
         _StyleDefs(62)  =   "Named:id=33:Normal"
         _StyleDefs(63)  =   ":id=33,.parent=0"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   ":id=34,.wraptext=-1"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton CmdAlteraDesconto 
         Caption         =   "A&lterar"
         Height          =   330
         Left            =   -67080
         TabIndex        =   33
         Top             =   1380
         Width           =   1080
      End
      Begin VB.CommandButton CmdAlteraDistribuicao 
         Caption         =   "&Al&terar"
         Height          =   330
         Left            =   -66885
         TabIndex        =   95
         Top             =   1380
         Width           =   1080
      End
      Begin Threed.SSPanel PnlCompDuplicata 
         Height          =   2775
         Left            =   -74400
         TabIndex        =   85
         Top             =   1305
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   4895
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   0
         Enabled         =   0   'False
         Begin Threed.SSPanel PanPagto2 
            Height          =   1545
            Left            =   60
            TabIndex        =   108
            Top             =   120
            Width           =   8145
            _Version        =   65536
            _ExtentX        =   14367
            _ExtentY        =   2725
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Begin VB.TextBox TxtDocPagto2 
               Height          =   315
               Left            =   6600
               MaxLength       =   7
               TabIndex        =   25
               Top             =   570
               Width           =   1515
            End
            Begin VB.TextBox TxtDuplicata 
               Height          =   315
               Left            =   1860
               MaxLength       =   1
               TabIndex        =   21
               Top             =   105
               Width           =   630
            End
            Begin VB.TextBox TxtValorDuplicata 
               Height          =   315
               Left            =   3465
               TabIndex        =   22
               Top             =   105
               Width           =   1605
            End
            Begin VB.ComboBox CboTipoPagto2 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":009C
               Left            =   1875
               List            =   "FrmNotasFiscais.frx":00B5
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   570
               Width           =   3000
            End
            Begin VB.ComboBox CboStatusduplicata 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":010F
               Left            =   6180
               List            =   "FrmNotasFiscais.frx":0119
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   105
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo CboCCorrente2 
               Bindings        =   "FrmNotasFiscais.frx":0137
               Height          =   330
               Left            =   1860
               TabIndex        =   26
               Top             =   1050
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "coco_cd_codigo"
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
            Begin VB.Label LblDocPagRec 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Doc. Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4740
               TabIndex        =   114
               Top             =   585
               Width           =   1815
            End
            Begin VB.Label LblCodDupl 
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1140
               TabIndex        =   113
               Top             =   135
               Width           =   720
            End
            Begin VB.Label LblValordupl 
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2865
               TabIndex        =   112
               Top             =   120
               Width           =   615
            End
            Begin VB.Label LblTipoPagamento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Tipo Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   111
               Top             =   585
               Width           =   1725
            End
            Begin VB.Label LblCCorrente2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Conta Corrente:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   420
               TabIndex        =   110
               Top             =   1050
               Width           =   1365
            End
            Begin VB.Label LblStatus3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Status:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   5355
               TabIndex        =   109
               Top             =   120
               Width           =   765
            End
         End
         Begin VB.CommandButton CmdSairDuplicata 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   7005
            TabIndex        =   31
            Top             =   2310
            Width           =   1140
         End
         Begin VB.CommandButton CmdConfirmaduplicata 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   5700
            TabIndex        =   30
            Top             =   2310
            Width           =   1140
         End
         Begin Threed.SSPanel PanDatas 
            Height          =   495
            Left            =   120
            TabIndex        =   86
            Top             =   2160
            Width           =   5385
            _Version        =   65536
            _ExtentX        =   9499
            _ExtentY        =   873
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin MSComCtl2.DTPicker DtpVencimento 
               Height          =   315
               Left            =   1215
               TabIndex        =   28
               Top             =   90
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               DateIsNull      =   -1  'True
               Format          =   24641537
               CurrentDate     =   37460
            End
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3855
               TabIndex        =   29
               Top             =   105
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   24641537
               CurrentDate     =   37460
            End
            Begin VB.Label LblVencimento2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   30
               TabIndex        =   89
               Top             =   120
               Width           =   1125
            End
            Begin VB.Label LblPagamento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2520
               TabIndex        =   87
               Top             =   120
               Width           =   1245
            End
         End
         Begin MSDataListLib.DataCombo CboBancoForn2 
            Bindings        =   "FrmNotasFiscais.frx":0156
            Height          =   330
            Left            =   1920
            TabIndex        =   27
            Top             =   1680
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "bafo_cd_codigo"
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
         Begin VB.Label LblBcFornecedor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Banco Fornecedor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   90
            Top             =   1680
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdInsereduplicata 
         Caption         =   "&Inserir"
         Height          =   330
         Left            =   -74790
         TabIndex        =   18
         Top             =   4305
         Width           =   1080
      End
      Begin VB.CommandButton CmdAlteraduplicata 
         Caption         =   "&Alterar"
         Height          =   330
         Left            =   -73480
         TabIndex        =   19
         Top             =   4305
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiduplicata 
         Caption         =   "&Excluir"
         Height          =   330
         Left            =   -72170
         TabIndex        =   20
         Top             =   4305
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiDesconto 
         Caption         =   "E&xcluir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   34
         Top             =   1800
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDesconto 
         Caption         =   "I&nserir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   32
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton CmdRemoverDistribuicao 
         Caption         =   "Excl&uir"
         Height          =   330
         Left            =   -66885
         TabIndex        =   72
         Top             =   1800
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDistribuicao 
         Caption         =   "Ins&erir"
         Height          =   330
         Left            =   -66885
         TabIndex        =   71
         Top             =   960
         Width           =   1080
      End
      Begin Threed.SSPanel PanNotasFiscais 
         Height          =   1695
         Left            =   120
         TabIndex        =   52
         Top             =   1130
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   2990
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox TxtValor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7090
            TabIndex        =   7
            Top             =   675
            Width           =   1800
         End
         Begin VB.TextBox TxtDevolucao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7095
            TabIndex        =   138
            Text            =   "0"
            Top             =   675
            Width           =   1800
         End
         Begin VB.TextBox TxtSerie 
            Height          =   315
            Left            =   8400
            MaxLength       =   3
            TabIndex        =   4
            Top             =   195
            Width           =   495
         End
         Begin MSAdodcLib.Adodc DatFornecedor 
            Height          =   330
            Left            =   2880
            Top             =   195
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "DatFornecedor"
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
         Begin MSAdodcLib.Adodc DatHistorico 
            Height          =   330
            Left            =   2520
            Top             =   1170
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
            Caption         =   "DatHistorico"
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
         Begin MSDataListLib.DataCombo CboFornecedor 
            Bindings        =   "FrmNotasFiscais.frx":0177
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   195
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "focl_tx_razaosocial"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   "CboFornecedor"
         End
         Begin MSMask.MaskEdBox MskNumdocumento 
            Height          =   315
            Left            =   6740
            TabIndex        =   3
            Top             =   200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboHistorico 
            Bindings        =   "FrmNotasFiscais.frx":0193
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Top             =   1165
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_codigo"
            Text            =   "CboHistorico"
         End
         Begin Threed.SSCommand CmdLimparHistorico 
            Height          =   330
            Left            =   5420
            TabIndex        =   88
            Top             =   1150
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            Enabled         =   0   'False
            MouseIcon       =   "FrmNotasFiscais.frx":01AE
            Picture         =   "FrmNotasFiscais.frx":0600
         End
         Begin MSComCtl2.DTPicker DtpEntrada 
            Height          =   315
            Left            =   1320
            TabIndex        =   5
            Top             =   675
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   24641537
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker DtpEmissao 
            Height          =   315
            Left            =   4095
            TabIndex        =   6
            Top             =   690
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   24641537
            CurrentDate     =   37460
         End
         Begin VB.Label LblValor1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6480
            TabIndex        =   56
            Top             =   750
            Width           =   555
         End
         Begin VB.Label LblDevolucao 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6480
            TabIndex        =   139
            Top             =   735
            Width           =   555
         End
         Begin VB.Label LblValorNota2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor a Pagar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   210
            Left            =   5880
            TabIndex        =   77
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label LblValorLiquido1 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   7440
            TabIndex        =   9
            Top             =   1165
            Width           =   1485
         End
         Begin VB.Label LblEntrada 
            Alignment       =   1  'Right Justify
            Caption         =   "Competência:"
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
            Left            =   60
            TabIndex        =   62
            Top             =   705
            Width           =   1185
         End
         Begin VB.Label LblEmissao 
            Alignment       =   1  'Right Justify
            Caption         =   "Emissão:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   61
            Top             =   705
            Width           =   885
         End
         Begin VB.Label LblSerie 
            Alignment       =   1  'Right Justify
            Caption         =   "Série:"
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
            Left            =   7680
            TabIndex        =   57
            Top             =   210
            Width           =   675
         End
         Begin VB.Label LblNumDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº Documento:"
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
            Left            =   5310
            TabIndex        =   55
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label LblHistorico 
            Alignment       =   1  'Right Justify
            Caption         =   "Histórico:"
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
            Left            =   120
            TabIndex        =   54
            Top             =   1185
            Width           =   1125
         End
         Begin VB.Label LblNomeFantasia 
            Alignment       =   1  'Right Justify
            Caption         =   "Fornecedor:"
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
            Left            =   120
            TabIndex        =   53
            Top             =   210
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   645
         Left            =   120
         TabIndex        =   58
         Top             =   430
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1138
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CboStatus 
            Height          =   315
            ItemData        =   "FrmNotasFiscais.frx":0712
            Left            =   1320
            List            =   "FrmNotasFiscais.frx":071F
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   160
            Width           =   2415
         End
         Begin MSAdodcLib.Adodc DatTipoDoc 
            Height          =   330
            Left            =   6360
            Top             =   165
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
            Caption         =   "DatTipoDoc"
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
         Begin MSDataListLib.DataCombo CboTipoDoc 
            Bindings        =   "FrmNotasFiscais.frx":0747
            Height          =   315
            Left            =   5400
            TabIndex        =   1
            Top             =   165
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "tido_tx_descricao"
            BoundColumn     =   "tido_cd_tipodoc"
            Text            =   "CboTipoDoc"
         End
         Begin VB.Label LblTipoDoc 
            Caption         =   "Tipo Documento:"
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
            Left            =   3840
            TabIndex        =   60
            Top             =   195
            Width           =   1515
         End
         Begin VB.Label LblStatus1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   59
            Top             =   195
            Width           =   645
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   1820
         Left            =   120
         TabIndex        =   63
         Top             =   2870
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3228
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Apropriação"
         TabPicture(0)   =   "FrmNotasFiscais.frx":0760
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblPcr"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblConta"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "CboPcr"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "MskPcr"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CboPlanoContas"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Mskconta"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Pagamento"
         TabPicture(1)   =   "FrmNotasFiscais.frx":077C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LblBcForn"
         Tab(1).Control(1)=   "CboBancoForn1"
         Tab(1).Control(2)=   "PanPagto1"
         Tab(1).ControlCount=   3
         Begin Threed.SSPanel PanPagto1 
            Height          =   1080
            Left            =   -74940
            TabIndex        =   102
            Top             =   360
            Width           =   9015
            _Version        =   65536
            _ExtentX        =   15901
            _ExtentY        =   1905
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Begin VB.ComboBox CboTipoPagto1 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":0798
               Left            =   1860
               List            =   "FrmNotasFiscais.frx":07B1
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   367
               Width           =   2655
            End
            Begin VB.TextBox TxtDocPagto1 
               Height          =   315
               Left            =   6555
               MaxLength       =   7
               TabIndex        =   15
               Top             =   360
               Width           =   1635
            End
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "FrmNotasFiscais.frx":080B
               Height          =   330
               Left            =   1860
               TabIndex        =   16
               Top             =   705
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "coco_cd_codigo"
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
            Begin MSComCtl2.DTPicker DtpVencParcelaUnica 
               Height          =   315
               Left            =   1860
               TabIndex        =   115
               Top             =   30
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   24641537
               CurrentDate     =   37460
            End
            Begin MSComCtl2.DTPicker DtpPagtoParcelaUnica 
               Height          =   315
               Left            =   4860
               TabIndex        =   116
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   24641537
               CurrentDate     =   37460
            End
            Begin Threed.SSCommand CmdEntMultiplas 
               Height          =   285
               Left            =   7410
               TabIndex        =   117
               Top             =   30
               Width           =   1470
               _Version        =   65536
               _ExtentX        =   2593
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "&Entradas Múltiplas"
               ForeColor       =   16711680
               Enabled         =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin VB.Label LblCCorrente 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Conta Corrente:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   450
               TabIndex        =   107
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label LblTipoPagto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tipo Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   90
               TabIndex        =   106
               Top             =   360
               Width           =   1725
            End
            Begin VB.Label LblDocPagto 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Doc. Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4665
               TabIndex        =   105
               Top             =   420
               Width           =   1815
            End
            Begin VB.Label LblVencimento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   690
               TabIndex        =   104
               Top             =   15
               Width           =   1125
            End
            Begin VB.Label LblPagto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Pagamento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   3555
               TabIndex        =   103
               Top             =   0
               Width           =   1245
            End
         End
         Begin MSMask.MaskEdBox Mskconta 
            Height          =   315
            Left            =   1920
            TabIndex        =   12
            Top             =   1170
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContas 
            Bindings        =   "FrmNotasFiscais.frx":082A
            Height          =   330
            Left            =   2880
            TabIndex        =   13
            Top             =   1170
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "plco_cd_codreduzido"
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
         Begin MSDataListLib.DataCombo CboBancoForn1 
            Bindings        =   "FrmNotasFiscais.frx":0841
            Height          =   330
            Left            =   -73080
            TabIndex        =   17
            Top             =   1430
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "bafo_cd_codigo"
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
         Begin MSMask.MaskEdBox MskPcr 
            Height          =   330
            Left            =   1920
            TabIndex        =   10
            Top             =   705
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr 
            Bindings        =   "FrmNotasFiscais.frx":0862
            Height          =   330
            Left            =   2880
            TabIndex        =   11
            Top             =   705
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "cere_cd_estruturado"
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
         Begin VB.Label LblConta 
            Alignment       =   1  'Right Justify
            Caption         =   "Conta Contábil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   1215
            Width           =   1605
         End
         Begin VB.Label LblPcr 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro Custo:"
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
            Left            =   240
            TabIndex        =   65
            Top             =   765
            Width           =   1605
         End
         Begin VB.Label LblBcForn 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco Fornecedor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   -74850
            TabIndex        =   64
            Top             =   1440
            Width           =   1725
         End
      End
      Begin Threed.SSPanel pandescontos 
         Height          =   1695
         Left            =   -74280
         TabIndex        =   80
         Top             =   975
         Visible         =   0   'False
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
         _ExtentY        =   2990
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   0
         Enabled         =   0   'False
         Begin VB.CommandButton CmdConfirmaDesconto 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   4080
            TabIndex        =   38
            Top             =   1200
            Width           =   1140
         End
         Begin VB.CommandButton CmdSairDesconto 
            BackColor       =   &H00000000&
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
            Height          =   330
            Left            =   5400
            TabIndex        =   39
            Top             =   1200
            Width           =   1140
         End
         Begin VB.TextBox TxtValorDesconto 
            Height          =   315
            Left            =   4800
            TabIndex        =   37
            Top             =   720
            Width           =   1755
         End
         Begin MSAdodcLib.Adodc DatDescontosNF 
            Height          =   330
            Left            =   2400
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "DatDescontosNF"
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
         Begin MSDataListLib.DataCombo CboDescontos 
            Bindings        =   "FrmNotasFiscais.frx":0877
            Height          =   315
            Left            =   1575
            TabIndex        =   35
            Top             =   240
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desc_tx_descricao"
            BoundColumn     =   "desc_cd_desconto"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DtpVencDesconto 
            Height          =   315
            Left            =   1575
            TabIndex        =   36
            Top             =   705
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   24641537
            CurrentDate     =   37460
         End
         Begin VB.Label LblVlDesc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3645
            TabIndex        =   83
            Top             =   735
            Width           =   1125
         End
         Begin VB.Label LblDesconto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descontos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   360
            TabIndex        =   82
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label LblVencimento3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   360
            TabIndex        =   81
            Top             =   735
            Width           =   1125
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   3135
         Left            =   -74280
         TabIndex        =   84
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
         _LayoutType     =   4
         _RowHeight      =   12
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).ValueItems(0)._DefaultItem=   0
         Columns(0).ValueItems(0).Value=   ""
         Columns(0).ValueItems(0).Value.vt=   8
         Columns(0).ValueItems(0).DisplayValue=   ""
         Columns(0).ValueItems(0).DisplayValue.vt=   8
         Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems.Count=   1
         Columns(0).Caption=   "Descrição"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Vencimento"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "dd/mm/yy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "codigo"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Fatura"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=6112"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6033"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1057296"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=88"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2619"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2540"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=1049105"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=1049362"
         Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(26)=   "Column(2)._MinWidth=87777292"
         Splits(0)._ColumnProps(27)=   "Column(3).Width=2064"
         Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=1048848"
         Splits(0)._ColumnProps(32)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(35)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(36)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(38)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(4)._ColStyle=1048848"
         Splits(0)._ColumnProps(40)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(31)  =   ":id=28,.locked=-1"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
         _StyleDefs(49)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(51)  =   "Named:id=33:Normal"
         _StyleDefs(52)  =   ":id=33,.parent=0"
         _StyleDefs(53)  =   "Named:id=34:Heading"
         _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   ":id=34,.wraptext=-1"
         _StyleDefs(56)  =   "Named:id=35:Footing"
         _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   "Named:id=36:Selected"
         _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(64)  =   "Named:id=39:EvenRow"
         _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(66)  =   "Named:id=40:OddRow"
         _StyleDefs(67)  =   ":id=40,.parent=33"
         _StyleDefs(68)  =   "Named:id=41:RecordSelector"
         _StyleDefs(69)  =   ":id=41,.parent=34"
         _StyleDefs(70)  =   "Named:id=42:FilterBar"
         _StyleDefs(71)  =   ":id=42,.parent=33"
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid2 
         Height          =   3255
         Left            =   -74805
         TabIndex        =   91
         Top             =   945
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   5741
         _LayoutType     =   4
         _RowHeight      =   27
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Vencimento"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "dd/mm/yy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Devolução"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor Líquido"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CodStatus"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Status"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "CodTipoPagto"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Tipo Pag."
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Doc. Pag."
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Data Pag."
         Columns(10).DataField=   ""
         Columns(10).NumberFormat=   "dd/mm/yy"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "CodConta"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Conta Corrente (Empresa)"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "CodBancoforn"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Conta Corrente (Fornecedor)"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Cód. Duplicata"
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   16
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=16"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1057296"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1693"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1614"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1057297"
         Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2275"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2196"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1057554"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(2)._MinWidth=1953261894"
         Splits(0)._ColumnProps(26)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=1057554"
         Splits(0)._ColumnProps(31)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(34)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=1057554"
         Splits(0)._ColumnProps(39)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(40)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(41)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(42)=   "Column(5).Width=2064"
         Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=1984"
         Splits(0)._ColumnProps(45)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=1049362"
         Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(49)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(50)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(51)=   "Column(6).Width=2143"
         Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=2064"
         Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=1057296"
         Splits(0)._ColumnProps(56)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(57)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(59)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(60)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(62)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=1048848"
         Splits(0)._ColumnProps(64)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(65)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(67)=   "Column(8).Width=1984"
         Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=1905"
         Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(8)._ColStyle=1057296"
         Splits(0)._ColumnProps(72)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(73)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(74)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(75)=   "Column(9).Width=1720"
         Splits(0)._ColumnProps(76)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(77)=   "Column(9)._WidthInPix=1640"
         Splits(0)._ColumnProps(78)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(79)=   "Column(9)._ColStyle=1057296"
         Splits(0)._ColumnProps(80)=   "Column(9).WrapText=1"
         Splits(0)._ColumnProps(81)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(82)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(83)=   "Column(10).Width=1535"
         Splits(0)._ColumnProps(84)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(85)=   "Column(10)._WidthInPix=1455"
         Splits(0)._ColumnProps(86)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(87)=   "Column(10)._ColStyle=1057297"
         Splits(0)._ColumnProps(88)=   "Column(10).WrapText=1"
         Splits(0)._ColumnProps(89)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(90)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(91)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(92)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(94)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(95)=   "Column(11)._ColStyle=1048848"
         Splits(0)._ColumnProps(96)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(97)=   "Column(11).WrapText=1"
         Splits(0)._ColumnProps(98)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(99)=   "Column(12).Width=3863"
         Splits(0)._ColumnProps(100)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(101)=   "Column(12)._WidthInPix=3784"
         Splits(0)._ColumnProps(102)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(103)=   "Column(12)._ColStyle=1057296"
         Splits(0)._ColumnProps(104)=   "Column(12).WrapText=1"
         Splits(0)._ColumnProps(105)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(106)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(107)=   "Column(13).Width=1270"
         Splits(0)._ColumnProps(108)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(109)=   "Column(13)._WidthInPix=1191"
         Splits(0)._ColumnProps(110)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(111)=   "Column(13)._ColStyle=1048848"
         Splits(0)._ColumnProps(112)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(113)=   "Column(13).WrapText=1"
         Splits(0)._ColumnProps(114)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(115)=   "Column(14).Width=3916"
         Splits(0)._ColumnProps(116)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(117)=   "Column(14)._WidthInPix=3836"
         Splits(0)._ColumnProps(118)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(119)=   "Column(14)._ColStyle=1057296"
         Splits(0)._ColumnProps(120)=   "Column(14).WrapText=1"
         Splits(0)._ColumnProps(121)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(122)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(123)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(124)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(125)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(126)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(127)=   "Column(15)._ColStyle=1048848"
         Splits(0)._ColumnProps(128)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(129)=   "Column(15).WrapText=1"
         Splits(0)._ColumnProps(130)=   "Column(15).Order=16"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=196,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(31)  =   ":id=28,.locked=-1"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(36)  =   ":id=32,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(41)  =   ":id=58,.locked=-1"
         _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=20,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=44,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=45"
         _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=47"
         _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=16,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=44,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=45"
         _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=47"
         _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=44,.alignment=1,.locked=0"
         _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=45"
         _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=47,.locked=0"
         _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(58)  =   ":id=66,.locked=-1"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=45"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=47"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=70,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=44"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=45"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=47"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=74,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(67)  =   ":id=74,.locked=-1"
         _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=45"
         _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=47"
         _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=78,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(72)  =   ":id=78,.locked=-1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=44,.alignment=2"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=45"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=47"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=82,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(77)  =   ":id=82,.locked=-1"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=44,.alignment=2"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=45"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=47"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=86,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=44"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=45"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=47"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=90,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(86)  =   ":id=90,.locked=-1"
         _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=44,.alignment=2"
         _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=45"
         _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=47"
         _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=94,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=91,.parent=44"
         _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=92,.parent=45"
         _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=93,.parent=47"
         _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=98,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(95)  =   ":id=98,.locked=-1"
         _StyleDefs(96)  =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=44,.alignment=2"
         _StyleDefs(97)  =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=45"
         _StyleDefs(98)  =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=47"
         _StyleDefs(99)  =   "Splits(0).Columns(15).Style:id=102,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(100) =   "Splits(0).Columns(15).HeadingStyle:id=99,.parent=44"
         _StyleDefs(101) =   "Splits(0).Columns(15).FooterStyle:id=100,.parent=45"
         _StyleDefs(102) =   "Splits(0).Columns(15).EditorStyle:id=101,.parent=47"
         _StyleDefs(103) =   "Named:id=33:Normal"
         _StyleDefs(104) =   ":id=33,.parent=0"
         _StyleDefs(105) =   "Named:id=34:Heading"
         _StyleDefs(106) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(107) =   ":id=34,.wraptext=-1"
         _StyleDefs(108) =   "Named:id=35:Footing"
         _StyleDefs(109) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(110) =   "Named:id=36:Selected"
         _StyleDefs(111) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(112) =   "Named:id=37:Caption"
         _StyleDefs(113) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(114) =   "Named:id=38:HighlightRow"
         _StyleDefs(115) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(116) =   "Named:id=39:EvenRow"
         _StyleDefs(117) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(118) =   "Named:id=40:OddRow"
         _StyleDefs(119) =   ":id=40,.parent=33"
         _StyleDefs(120) =   "Named:id=41:RecordSelector"
         _StyleDefs(121) =   ":id=41,.parent=34"
         _StyleDefs(122) =   "Named:id=42:FilterBar"
         _StyleDefs(123) =   ":id=42,.parent=33"
      End
      Begin Threed.SSPanel PnlDados1 
         Height          =   360
         Left            =   -74280
         TabIndex        =   96
         Top             =   480
         Width           =   5880
         _Version        =   65536
         _ExtentX        =   10372
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados1_1 
         Height          =   360
         Left            =   -68280
         TabIndex        =   97
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados2 
         Height          =   360
         Left            =   -74280
         TabIndex        =   98
         Top             =   480
         Width           =   5880
         _Version        =   65536
         _ExtentX        =   10372
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados2_2 
         Height          =   360
         Left            =   -68280
         TabIndex        =   99
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados3 
         Height          =   360
         Left            =   -74400
         TabIndex        =   100
         Top             =   480
         Width           =   6120
         _Version        =   65536
         _ExtentX        =   10795
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados3_3 
         Height          =   360
         Left            =   -68160
         TabIndex        =   101
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCommand CmdDuplMultiplas 
         Height          =   330
         Left            =   -70860
         TabIndex        =   129
         Top             =   4305
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&Parcelas Múltiplas"
         ForeColor       =   16711680
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSPanel PanFornecedor 
         Height          =   1680
         Left            =   120
         TabIndex        =   130
         Top             =   1125
         Visible         =   0   'False
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   2963
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   0
         Enabled         =   0   'False
         Begin VB.CommandButton Command1 
            BackColor       =   &H00000000&
            Caption         =   "&Retornar"
            Height          =   240
            Left            =   1905
            TabIndex        =   132
            Top             =   1335
            Width           =   780
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00000000&
            Caption         =   "&Confirmar"
            Height          =   240
            Left            =   1080
            TabIndex        =   131
            Top             =   1335
            Width           =   780
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   1215
            Left            =   90
            TabIndex        =   133
            Top             =   75
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   2143
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.OptionButton OptProdServ 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produtos e Serviços"
               Height          =   255
               Left            =   210
               TabIndex        =   137
               Top             =   915
               Width           =   1725
            End
            Begin VB.OptionButton OptProduto 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produtos"
               Height          =   255
               Left            =   210
               TabIndex        =   136
               Top             =   630
               Width           =   1725
            End
            Begin VB.OptionButton OptFisica 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pessoa Física"
               Height          =   255
               Left            =   210
               TabIndex        =   135
               Top             =   330
               Width           =   1725
            End
            Begin VB.OptionButton OptJuridica 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pessoa Jurídica"
               Height          =   255
               Left            =   210
               TabIndex        =   134
               Top             =   45
               Value           =   -1  'True
               Width           =   1725
            End
         End
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -70080
         TabIndex        =   94
         Top             =   4365
         Width           =   1170
      End
      Begin VB.Label LblValorTotal 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67965
         TabIndex        =   93
         Top             =   4320
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -69045
         TabIndex        =   92
         Top             =   4365
         Width           =   1170
      End
      Begin VB.Label LblValorTotalDistrib 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -69000
         TabIndex        =   76
         Top             =   4320
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -70200
         TabIndex        =   75
         Top             =   4395
         Width           =   1050
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Nota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -73920
         TabIndex        =   74
         Top             =   4365
         Width           =   1410
      End
      Begin VB.Label LblValorNota 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72600
         TabIndex        =   73
         Top             =   4320
         Width           =   1485
      End
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   4004
      TabIndex        =   49
      Top             =   4950
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5865
      TabIndex        =   50
      Top             =   4950
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DatPcr 
      Height          =   330
      Left            =   8475
      Top             =   4815
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DatPcr"
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
   Begin MSAdodcLib.Adodc DatConta 
      Height          =   330
      Left            =   8475
      Top             =   5055
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DatConta"
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
   Begin MSAdodcLib.Adodc DatContaCorrente 
      Height          =   330
      Left            =   7575
      Top             =   4800
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
      Caption         =   "DatContaCorrente"
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
   Begin MSAdodcLib.Adodc DatBancoFornecedor 
      Height          =   330
      Left            =   7560
      Top             =   5055
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
      Caption         =   "DatBancoFornecedor"
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
   Begin Threed.SSCommand CmdNovo 
      Height          =   330
      Left            =   2130
      TabIndex        =   128
      Top             =   4950
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Novo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FrmNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ValorTotalDuplicata As Double 'Acumula o total de duplicatas até compor o valor total da Nota Fiscal
Dim ValorTotalDistribuicao As Double 'Acumula o total de distribuições até compor o valor total da Nota Fiscal
Dim Xtipomov As String
Dim Xdebcred As String
Dim Xbancoforn As Integer
Dim XConta As Integer
Dim xcont As Integer 'Controle na ditribuicao
Dim XContDesc As Integer 'Controle em descontos
Dim XDistribuido As String 'verifica se o valor total foi distribuido
Dim XPcr As String  'Armazena o pcr digitado na distribuicao
Dim xcc As String   'Armazena a conta contábil digitada na distribuicao
Dim xcod As String 'Armazena o desconto escolhido na alteracao do desconto
Dim xNumFat As Long
Dim Tipo As String 'Verifica o tipo do documento
Dim ParcelasMultiplas As Boolean 'Verifica se foi gerada parcelas multiplas
Dim XTipoFatura As String '1=(1 nota 1 dupl) 2=(1 nota 1 ou + dupl) 3=(2 ou + notas p/ 1 ou + dupl)
Dim XValorBruto As Double  'Armazena o valor bruto da nota
Dim Status As String 'Armazena o Status da Nota Fiscal
Dim XVlDuplicata As Double 'Valor inicial da duplicata que está sendo alterada
Dim NumDocAnt As String 'armazena o nº do documento no caso de uma alteração

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor1 As New XArray
Dim Vetor2 As New XArray
Dim Vetor3 As New XArray
Dim Vetor4 As New XArray 'Descontos/Acrescimos de duplicata
Dim VetAnt As New XArray 'Armazena os valores do vetor antes das alterações
Dim EntrouDesc As Boolean 'Controla a entrada na orelha de desconto
Dim EntrouDupl As Boolean 'Controla a entrada na orelha de duplicata
Dim EntrouDist As Boolean 'Controla a entrada na orelha de distribuicao
Dim XDistribuicao As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDuplicata As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDesconto As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XFT_INSPARCMULTIPLA As String
Dim XFT_INSDUPLMULTIPLA As String
Dim XFT_NOTADUPLICATA As String ' Informar se é Nota ou Duplicata

'Variável utilizada na alteração da duplicata
Dim PosCursor As Integer 'Guarda o Bookmark do grid

Dim errLoop As ADODB.Error
Dim StrError As String
Function ConsistenciaStatus() As Boolean
    Dim ExisteStatus As Boolean
    Dim x As Byte
    'Não é preciso prever a mudança para previsto nem para tipo de documento = 0
    If CboStatus.ListIndex <> 0 And Tipo = 0 Then
        If EntrouDupl Or XInserir = "A" Then
            If Vetor2(0, 0) <> Empty Then 'REVER
            ExisteStatus = False
            If CboStatus.ListIndex = 1 Then
                'Se o status for Compromissado, pelo menos 1 parcela deve ser compromissada
                'If Vetor2(0, 0) <> Empty Then 'REVER
                While x <= Vetor2.UpperBound(1)
                    If Vetor2(x, 5) = "C" Then
                        ExisteStatus = True
                        x = Vetor2.UpperBound(1) + 1
                    ElseIf Vetor2(x, 5) = "R" Then
                        x = x + 1
                    ElseIf Vetor2(x, 5) = "P" Then
                        Vetor2(x, 5) = "C"
                        Vetor2(x, 6) = "Compromissado"
                        ExisteStatus = True
                        EntrouDupl = True
                        DtpVencParcelaUnica.Value = Null
                        x = x + 1
                    End If
                Wend
                'End If
                If ExisteStatus = False Then
                    MsgBox "O Status de uma das Parcelas deve ser 'Compromissado'!", vbCritical, "ATENÇÃO"
                    ConsistenciaStatus = False
                    Exit Function
                Else
                    ConsistenciaStatus = True
                    Exit Function
                End If
            ElseIf CboStatus.ListIndex = 2 Then
                'Se o status for Realizado, todas as parcelas deve ser realizadas
                'If Vetor2(0, 0) <> Empty Then 'REVER
                While x <= Vetor2.UpperBound(1)
                    If Vetor2(x, 5) = "R" Then
                        ExisteStatus = True
                        x = x + 1
                    ElseIf Vetor2(x, 5) = "C" Then
                        ExisteStatus = False
                        x = Vetor2.UpperBound(1) + 1
                    ElseIf Vetor2(x, 5) = "P" Then
                        Vetor2(x, 5) = "R"
                        Vetor2(x, 6) = "Realizado"
                        ExisteStatus = True
                        EntrouDupl = True
                        DtpVencParcelaUnica.Value = Null
                        x = x + 1
                    End If
                Wend
                'End If
                If ExisteStatus = False Then
                    MsgBox "O Status das Parcelas deve ser 'Realizado'!", vbCritical, "ATENÇÃO"
                    ConsistenciaStatus = False
                    Exit Function
                Else
                    ConsistenciaStatus = True
                    Exit Function
                End If
            End If
'        ElseIf XInserir = "A" And Not EntrouDupl Then  'Será acessada qdo clicar na orelha de dupl ou qdo gravar sem ter entrado na orelha
'            TDBGrid2.MoveFirst
'            If Not (TDBGrid2.BOF And TDBGrid2.EOF) Then
'               While Not TDBGrid2.EOF
'                    TDBGrid2.Columns("CodStatus").Text = CboStatus.ListIndex
'                    TDBGrid2.Update
'                    TDBGrid2.MoveNext
'               Wend
'            End If
        End If
        End If
    Else
        If EntrouDesc Then
            If Vetor1(0, 0) <> Empty Then 'Status previsto ou tipo =1 não possui descontos
                Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & NFatura
                LblValorLiquido1.Caption = TxtValor.Text
                ConsistenciaStatus = True
                Exit Function
            End If
        End If
    End If
    ConsistenciaStatus = True
End Function

Function VerificaValores() As Boolean
    Dim SomaDesc As Double
    Dim Somadupl As Double
    Dim Sql As String
    Dim XEntrouDupl As Boolean
    Dim XEntrouDesc As Boolean

    'VerificaValores = False
    XEntrouDupl = EntrouDupl
    XEntrouDesc = EntrouDesc
    If XInserir = "A" Then
        If XTipoFatura <> 4 Then 'não possui fatura
            If Not XEntrouDupl Then
                XEntrouDupl = True
            End If
            If Not XEntrouDesc Then
                XEntrouDesc = True
            End If
        End If
    End If
    
    If XEntrouDesc Then
        If Vetor1(0, 0) <> Empty Or Vetor2(0, 0) <> Empty Then
            SomaDesc = funSomaColuna(Vetor1, 2)
            Somadupl = funSomaColuna(Vetor2, 2)
            If Format(TxtValor.Text, "standard") <> Format((Somadupl + SomaDesc), "standard") Then
            'If CDbl(TxtValor.Text) <> (Somadupl + SomaDesc) Then
                MsgBox "A soma das Parcelas com os Descontos é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                TabNotasFiscais.Tab = 1
                If Format(TxtValor.Text, "standard") > Format((Somadupl + SomaDesc), "standard") Then
                    CmdInsereduplicata_Click
                    TxtValorDuplicata.Text = (CDbl(TxtValor.Text) - (Somadupl + SomaDesc))
                End If
                VerificaValores = True
                Exit Function
            End If
        End If
    ElseIf XEntrouDupl Then
        If Vetor2(0, 0) <> Empty Then
            Somadupl = funSomaColuna(Vetor2, 2)
            If Format(TxtValor.Text, "standard") <> Format(Somadupl, "standard") Then
                MsgBox "A soma das Parcelas é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                TabNotasFiscais.Tab = 1
                If Format(TxtValor.Text, "standard") > Format(Somadupl, "standard") Then
                    CmdInsereduplicata_Click
                    TxtValorDuplicata.Text = Format((CDbl(TxtValor.Text) - Somadupl), "standard")
                End If
                VerificaValores = True
                Exit Function
            End If
        Else 'entrou na orelha, mas não cadasrou nenhuma duplicata
            If CboStatus.ListIndex = "2" And Tipo = "0" Then
                MsgBox "A composição da Duplicata é obrigatória para Notas Fiscais com status 'Realizado'.", vbCritical, "ATENÇÃO"
                VerificaValores = True
                Exit Function
            Else
                EntrouDupl = False
            End If
        End If
    ElseIf Not XEntrouDupl And CboStatus.ListIndex = "2" And Tipo = "0" Then 'Não entrou na orelha de duplicata é realizada e não compoe duplicata automat.
        MsgBox "A composição da Duplicata é obrigatória para Notas Fiscais com status 'Realizado'.", vbCritical, "ATENÇÃO"
        VerificaValores = True
        Exit Function
    End If
End Function
Sub InsereRegistroVetorDupl(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)  'Carrega o vetores
    Dim Cont As Integer
    Dim x As Integer
    
    'Insere um regitro no vetor
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
        Vetor.Insert 1, Vetor.UpperBound(1) + 1
    End If
    Cont = Vetor.UpperBound(1)
    For x = 0 To Grid.Columns.Count - 2
        subSetaVal Vetor, Cont, x, FunNulo(VetorCampos(x))
    Next
    With Grid
        .ReBind
    End With
End Sub

Function ExisteDistribuicao(Numchaves As Byte, coluna As Variant, Valor As Variant, Vetor As XArray, Grid As TDBGrid) As Boolean
    Dim Cont As Integer, Achou As Boolean
    Dim x As Integer, y As Integer
    
    'Verifica se o registro está cadastrado no vetor
    
    ExisteDistribuicao = False
    
    While x <= Vetor.UpperBound(1)
        If Vetor(x, coluna(0)) = CStr(Valor(0)) Then
            If Numchaves = 1 Then
                ExisteDistribuicao = True
                x = Vetor.UpperBound(1)
            Else
                Achou = True
                y = 1
                While Achou And y <= Numchaves - 1
                    'If coluna(Y) <> Valor(Y) Then
                    If Vetor(x, coluna(y)) <> CStr(Valor(y)) Then
                        Achou = False
                    End If
                    y = y + 1
                Wend
                If Achou Then
                    ExisteDistribuicao = True
                    x = Vetor.UpperBound(1)
                End If
            End If
        End If
        x = x + 1
    Wend
End Function

Sub HabilitaControles()
    If CboStatus.ListIndex = 0 Then 'Previsto
        PanNotasFiscais.Enabled = True
        CboFornecedor.Enabled = True
        'MskNumdocumento.Enabled = True
        'TxtSerie.Enabled = False
        DtpEntrada.Enabled = True
        DtpEntrada.Value = Format(Now, "dd/MM/yy")
        DtpEmissao.Value = Format(Now, "dd/MM/yy")
        DtpEmissao.Enabled = False
        TxtValor.Enabled = True
        CboHistorico.Enabled = True
        CmdLimparHistorico.Enabled = True
        LblValorLiquido1.Enabled = True
        
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
        TabNotasFiscais.TabEnabled(3) = False
        TabAuxiliar.Enabled = True
        CboPcr.Enabled = True
        Mskconta.Enabled = True
        CboPlanoContas.Enabled = True
        TabAuxiliar.TabEnabled(0) = True
        TabAuxiliar.TabEnabled(1) = True
        If DtpPagtoParcelaUnica.Enabled Then DtpPagtoParcelaUnica.Value = ""
        DtpPagtoParcelaUnica.Enabled = False
        If TxtDocPagto1.Enabled Then TxtDocPagto1.Text = ""
        TxtDocPagto1.Enabled = False
        ChkDistribuicao.Value = 0
        ChkDistribuicao.Enabled = False
        If XInserir = "I" Then
            CmdEntMultiplas.Enabled = True
        End If
    ElseIf CboStatus.ListIndex > 0 Then
        CmdEntMultiplas.Enabled = False
        PanNotasFiscais.Enabled = True
        CboFornecedor.Enabled = True
        'MskNumdocumento.Enabled = True
        DtpEntrada.Enabled = True
        DtpEmissao.Enabled = True
        TxtValor.Enabled = True
        CboHistorico.Enabled = True
        CmdLimparHistorico.Enabled = True
        LblValorLiquido1.Enabled = True
        If XInserir = "I" Then
            ChkDistribuicao.Value = 0
        End If
        ChkDistribuicao.Enabled = True

        If ChkDistribuicao.Value = "1" Then
            TabAuxiliar.TabEnabled(0) = False
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            Mskconta.Enabled = False
            CboPlanoContas.Enabled = False
        Else
            TabAuxiliar.TabEnabled(0) = True
            MskPcr.Enabled = True
            CboPcr.Enabled = True
            Mskconta.Enabled = True
            CboPlanoContas.Enabled = True
        End If

        If CboStatus.ListIndex = 1 Then 'Compromissado
            'TxtSerie.Enabled = False
            If Tipo = "1" Then 'gera duplicata automaticamente
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = True
                TabNotasFiscais.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(2) = False
                If DtpPagtoParcelaUnica.Enabled Then DtpPagtoParcelaUnica.Value = ""
                DtpPagtoParcelaUnica.Enabled = False
                If TxtDocPagto1.Enabled Then TxtDocPagto1.Text = ""
                TxtDocPagto1.Enabled = False
            Else     ' Não gera duplicata automaticamente
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(1) = True
                DtpPagamento.Enabled = False
                TxtDocPagto1.Enabled = False
                'If EntSaida = "E" Then TabNotasFiscais.TabEnabled(2) = True
                'TESTE
                TabNotasFiscais.TabEnabled(2) = True
            End If
            ElseIf CboStatus.ListIndex = 2 Then 'Realizado
            'TxtSerie.Enabled = True
            If Tipo = "1" Then  'gera duplicata automaticamente
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = True
                DtpPagtoParcelaUnica.Enabled = True
                TxtDocPagto1.Enabled = True
                TabNotasFiscais.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(2) = False
            Else     ' Não gera duplicata automaticamente
                DtpPagamento.Enabled = True
                TxtDocPagto2.Enabled = True
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(1) = True
                'If EntSaida = "E" Then TabNotasFiscais.TabEnabled(2) = True
                'TESTE
                TabNotasFiscais.TabEnabled(2) = True
            End If
        End If
    End If
    If XTipoFatura = "3" Then
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
    End If
End Sub

Sub DesabilitaRealizado(XLT_DOCUMENTO As String, XLB_TRUEFALSE As Boolean)
    If XLT_DOCUMENTO = "Nota" Then
        PanNotasFiscais.Enabled = True
        CboStatus.Enabled = False
        CboTipoDoc.Enabled = False
        CboFornecedor.Enabled = False
        MskNumdocumento.Enabled = False
        TxtSerie.Enabled = False
        DtpEntrada.Enabled = False
        DtpEmissao.Enabled = False
        TxtValor.Enabled = False
        CboHistorico.Enabled = False
        CmdLimparHistorico.Enabled = False
        TabAuxiliar.Enabled = True
        
        ChkDistribuicao.Enabled = True
        If ChkDistribuicao.Value = "1" Then
            TabAuxiliar.TabEnabled(0) = False
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            Mskconta.Enabled = False
            CboPlanoContas.Enabled = False
        Else
            TabAuxiliar.TabEnabled(0) = True
            MskPcr.Enabled = True
            CboPcr.Enabled = True
            Mskconta.Enabled = True
            CboPlanoContas.Enabled = True
        End If

        If Tipo = "1" Then  'gera duplicata automaticamente
            TabAuxiliar.TabEnabled(1) = True
            DtpVencParcelaUnica.Enabled = False
            DtpPagtoParcelaUnica.Enabled = False
            CboTipoPagto1.Enabled = False
            TxtDocPagto1.Enabled = False
            CboCCorrente.Enabled = False
            CboBancoForn1.Enabled = False
            TabNotasFiscais.TabEnabled(1) = False
            TabNotasFiscais.TabEnabled(2) = False
        Else     ' Não gera duplicata automaticamente
            TabAuxiliar.TabEnabled(1) = False
            TabNotasFiscais.TabEnabled(1) = True
            TabNotasFiscais.TabEnabled(2) = True
            CmdInsereduplicata.Enabled = False
            CmdAlteraduplicata.Enabled = False
            CmdExcluiduplicata.Enabled = False
            CmdInsereDesconto.Enabled = False
            CmdExcluiDesconto.Enabled = False
        End If
    Else
        TxtDuplicata.Enabled = XLB_TRUEFALSE
        TxtValorDuplicata.Enabled = XLB_TRUEFALSE
        CboStatusduplicata.Enabled = XLB_TRUEFALSE
        CboTipoPagto2.Enabled = XLB_TRUEFALSE
        CboCCorrente2.Enabled = XLB_TRUEFALSE
        TxtDocPagto2.Enabled = XLB_TRUEFALSE
        CboBancoForn2.Enabled = XLB_TRUEFALSE
        DtpVencimento.Enabled = XLB_TRUEFALSE
        DtpPagamento.Enabled = XLB_TRUEFALSE
        CmdConfirmaduplicata.Enabled = XLB_TRUEFALSE
    End If
    
    
End Sub

Sub DesabilitaBotoes()
    If TabNotasFiscais.Tab = 1 Then
        CmdInsereduplicata.Enabled = False
        CmdAlteraduplicata.Enabled = False
        CmdExcluiduplicata.Enabled = False
    ElseIf TabNotasFiscais.Tab = 2 Then
        CmdInsereDesconto.Enabled = False
        CmdExcluiDesconto.Enabled = False
    ElseIf TabNotasFiscais.Tab = 3 Then
        CmdInsereDistribuicao.Enabled = False
        CmdRemoverDistribuicao.Enabled = False
        CmdAlteraDistribuicao.Enabled = False
    End If
End Sub
                                                        
Sub HabilitaBotoes()
    If TabNotasFiscais.Tab = 1 Then
        CmdInsereduplicata.Enabled = True
        CmdAlteraduplicata.Enabled = True
        CmdExcluiduplicata.Enabled = True
    ElseIf TabNotasFiscais.Tab = 2 Then
        CmdInsereDesconto.Enabled = True
        CmdExcluiDesconto.Enabled = True
    ElseIf TabNotasFiscais.Tab = 3 Then
        CmdInsereDistribuicao.Enabled = True
        CmdRemoverDistribuicao.Enabled = True
        CmdAlteraDistribuicao.Enabled = True
    End If
End Sub
Sub PreencheCampos(XForm As String)
    Dim Sql As String
    Dim ResDoc As Object
    Dim ResDist As Object
    Dim ResDupl As Object
    Dim ResNota As Object
    Dim x As Integer
        
    'Tab 0
   
    TabNotasFiscais.TabEnabled(1) = True
    MskNumdocumento.Enabled = True
    SubQOpenRecordset ResNota, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal =" & Chave, "Estatico"
    If Not (ResNota.BOF And ResNota.EOF) Then
        XTipoFatura = ResNota!nofi_tx_tipofatura
        XValorBruto = ResNota!nofi_vl_valor
        Status = ResNota!nofi_tx_status
        If Status = "P" Then 'Previsto
           Status = "0"
        ElseIf Status = "C" Then 'Compromissado
               Status = "1"
        ElseIf Status = "R" Then 'Realizado
               Status = "2"
        End If
        CboStatus.ListIndex = Status
        CboTipoDoc.BoundText = FunNulo(ResNota!tido_cd_tipodoc)
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "TB" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "PC" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "DEV" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "REST" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "RA" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "NP" Then
            CboTipoDoc.Enabled = False
        End If
        'If Not IsNull(ResNota!nofi_tx_controle) And ResNota!nofi_tx_controle <> "" And EntSaida = "E" Then
        If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "DEV" Then
            'subConectarControleDados DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'F' ORDER BY focl_tx_fantasia ", Estatico
            'subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and (empr_cd_empresa = " & PCodEmpresa & " or isnull(empr_cd_empresa)) ORDER BY focl_tx_fantasia ", Estatico
            subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and (empr_cd_empresa = " & PCodEmpresa & " or empr_cd_empresa is null) ORDER BY focl_tx_razaosocial ", Estatico
        End If
        If Not IsNull(ResNota!focl_cd_forncli) Then
            CboFornecedor.BoundText = ResNota!focl_cd_forncli
        End If
        If EntSaida = "S" Then
            subConectarControleDados DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
            'Verifica se existe Devolução
            If Not IsNull(ResNota!nofi_cd_notaGlosaDev) Then
                LblDevolucao.Caption = "Devolução: "
                LblEmissao.Left = 2550
                DtpEmissao.Left = 3465
                LblValor1.Left = 4965
                TxtValor.Left = 5535
                TxtValor.Width = 1305
                LblDevolucao.Left = 6930
                LblDevolucao.Width = 975
                TxtDevolucao.Left = 7920
                TxtDevolucao.Width = 1050
                
                TDBGrid2.Columns("Valor Líquido").Visible = True
                TDBGrid2.Columns(3).Caption = "Devolução"
                TDBGrid2.Columns(3).Visible = True
                
                TDBGrid3.Columns("Valor Líquido").Visible = True
                TDBGrid3.Columns(5).Caption = "Devolução"
                TDBGrid3.Columns(5).Visible = True
            End If
        Else
            'Verifica se existe Glosa
            If Not IsNull(ResNota!nofi_cd_notaGlosaDev) Then
                LblDevolucao.Caption = "Glosa: "
                LblEmissao.Left = 2700
                DtpEmissao.Left = 3615
                LblValor1.Left = 5215
                TxtValor.Left = 5825
                TxtValor.Width = 1305
                LblDevolucao.Left = 6880
                LblDevolucao.Width = 975
                TxtDevolucao.Left = 7870
                TxtDevolucao.Width = 1050
                
                TDBGrid2.Columns("Valor Líquido").Visible = True
                TDBGrid2.Columns(3).Caption = "Glosa"
                TDBGrid2.Columns(3).Visible = True
                
                TDBGrid3.Columns("Valor Líquido").Visible = True
                TDBGrid3.Columns(5).Caption = "Glosa"
                TDBGrid3.Columns(5).Visible = True
            End If
        End If
        If ResNota!nofi_tx_distribuir = "S" Then
            ChkDistribuicao.Value = 1
            XDistribuido = "S"
        Else
            XDistribuido = "N"
            TabNotasFiscais.TabEnabled(3) = False
        End If
'        CboTipoDoc.BoundText = FunNulo(ResNota!tido_cd_tipodoc)
'        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
'        If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "TB" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "PC" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "DEV" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "REST" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "RA" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "NP" Then
'            CboTipoDoc.Enabled = False
'        End If
        MskNumdocumento.Text = FunNulo(ResNota!nofi_nr_documento)
        NumDocAnt = MskNumdocumento.Text
        DtpEntrada.Value = Format$(ResNota!nofi_dt_entrada, "DD/MM/YY")
        If Status <> "0" Then
            If IsDate(ResNota!nofi_dt_emissao) Then
                DtpEmissao.Value = Format$(ResNota!nofi_dt_emissao, "DD/MM/YY")
            End If
            ChkDistribuicao.Enabled = True
        End If
        TxtSerie.Text = FunNulo(ResNota!nofi_tx_serie)
        CboHistorico.Text = FunNulo(ResNota!nofi_tx_historico)
        If Not IsNull(ResNota!nofi_vl_valor) Then
            TxtValor.Text = Format$(ResNota!nofi_vl_valor, "standard")
        End If
        
    End If
    ResNota.Close
    
    Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
    'subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2 'ACCESS (testar no SQL Server)
    'subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
    subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_vl_GlosaDevolucao", "Valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
    TxtDevolucao.Text = Format(funSomaColuna(Vetor2, 3), "standard")
    
    Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
    subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura"), Vetor1, TDBGrid1
    
    'Sql = "SELECT Distribuicao.*,Notasfiscais.nofi_vl_valor,(Distribuicao.dist_vl_valor / NotasFiscais.nofi_vl_valor) AS Percentual,CentroResultados.cere_cd_estruturado, PlanoContas.plco_cd_codreduzido " & _
    "FROM NotasFiscais,Distribuicao,CentroResultados, PlanoContas WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " and Distribuicao.nofi_cd_notafiscal = " & Chave & " and Distribuicao.cere_cd_pcr =CentroResultados.cere_cd_pcr and Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta"
    'subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
    Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " and nofi_cd_notafiscal = " & Chave & ""
    subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "dist_vl_GlosaDevolucao", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
        
    Sql = "SELECT DuplDescAcresc.* from DuplDescAcresc LEFT JOIN Duplicatas ON Duplicatas.dupl_cd_duplicata = DuplDescAcresc.dupl_cd_duplicata Where Duplicatas.dupl_nr_fatura = " & NFatura
    CarregaVetorDescAcresc Sql, Array("dupl_cd_duplicata", "deac_cd_descacresc", "dude_vl_descacresc"), Vetor4, 3
         
    If Vetor1(0, 0) <> Empty Then
       LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
    Else
        LblValorLiquido1.Caption = Format$(CDbl(TxtValor.Text) - CDbl(TxtDevolucao), "##,##0.00")
    End If
    If Status = "0" Or XDistribuido = "N" Then 'É previsto ou não foi realizada distribuição e sim apropriação
        SubQOpenRecordset ResDist, "SELECT Distribuicao.*,PlanoContas.plco_cd_codreduzido, PlanoContas.plco_tx_desccontabil,CentroResultados.cere_cd_estruturado ,CentroResultados.cere_tx_descricao FROM ((Distribuicao LEFT JOIN PlanoContas ON Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta) LEFT JOIN CentroResultados ON Distribuicao.cere_cd_pcr = CentroResultados.cere_cd_pcr) where nofi_cd_notafiscal =" & Chave, "Estatico"
        If Not (ResDist.BOF And ResDist.EOF) Then
           MskPcr.Text = ResDist!cere_cd_estruturado
           CboPcr.BoundText = MskPcr.Text
           Mskconta.Text = ResDist!plco_cd_codreduzido
           CboPlanoContas.BoundText = Mskconta.Text
           'MskConta_LostFocus
        End If
        ResDist.Close
    End If
    
    SubQOpenRecordset ResDoc, "SELECT tipodocumentos.tido_cd_tipodoc,tipodocumentos.tido_tx_duplicata FROM TipoDocumentos where TipoDocumentos.tido_cd_tipodoc = " & CboTipoDoc.BoundText, "Estatico"
    If Not ResDoc.BOF And Not ResDoc.BOF Then
       Tipo = ResDoc!tido_tx_duplicata
       ResDoc.Close
    End If

    If Status = "0" Or Tipo = "1" Then 'É previsto ou gera duplicata automaticamente
        'subQOpenRecordset ResDupl, "SELECT Duplicatas.*,BancoFornecedor.bafo_cd_codigo,ContasCorrente.coco_tx_conta,ContasCorrente.coco_cd_agencia FROM ((Duplicatas LEFT JOIN BancoFornecedor ON Duplicatas.bafo_cd_codigo = BancoFornecedor.bafo_cd_codigo) LEFT JOIN ContasCorrente ON Duplicatas.coco_cd_codigo = ContasCorrente.coco_cd_codigo) where dupl_nr_fatura =" & NFatura, "Estatico"
        SubQOpenRecordset ResDupl, "SELECT Duplicatas.*,ContasCorrente.coco_tx_conta,ContasCorrente.coco_cd_agencia FROM ((Duplicatas LEFT JOIN BancoFornecedor ON Duplicatas.bafo_cd_codigo = BancoFornecedor.bafo_cd_codigo) LEFT JOIN ContasCorrente ON Duplicatas.coco_cd_codigo = ContasCorrente.coco_cd_codigo) where dupl_nr_fatura =" & NFatura, "Estatico"
        If Not (ResDupl.BOF And ResDupl.EOF) Then
             DtpVencParcelaUnica.Value = Format$(ResDupl!dupl_dt_vencimento, "DD/MM/YY")
             If Status = "2" Then
                DtpPagtoParcelaUnica.Value = Format$(ResDupl!dupl_dt_pagrec, "DD/MM/YY")
             End If
             If Not IsNull(ResDupl!dupl_tx_TipoPagRec) Then
                For x = 0 To CboTipoPagto1.ListCount - 1
                    If CboTipoPagto1.ItemData(x) = CInt(ResDupl!dupl_tx_TipoPagRec) Then
                        CboTipoPagto1.ListIndex = x
                    End If
                Next
             End If
             TxtDocPagto1.Text = FunNulo(ResDupl!dupl_tx_DocPagRec)
             If Not IsNull(ResDupl!coco_cd_codigo) Then
                CboCCorrente.BoundText = ResDupl!coco_cd_codigo
             End If
             If Not IsNull(ResDupl!bafo_cd_codigo) Then
                CboBancoForn1.BoundText = ResDupl!bafo_cd_codigo
             End If
       End If
       ResDupl.Close
    End If
End Sub
Sub CarregaVetorDescAcresc(Sql As String, VetorCampos As Variant, Vetor As XArray, Tam As Integer)    'Carrega o vetores
    Dim ResAvaliacao As Object
    Dim Cont As Integer
    Dim x As Integer
    
    'Carrega os valores no vetor
    
    'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
    'VetorCampos - vetor simples com os campos na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    SubQOpenRecordset ResAvaliacao, Sql, Estatico
    Vetor.ReDim 0, 0, 0, Tam - 1
    Vetor.Clear
    If Not (ResAvaliacao.EOF And ResAvaliacao.BOF) Then
        ResAvaliacao.MoveFirst
        Do While Not ResAvaliacao.EOF
            If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
                Vetor.Insert 1, Vetor.UpperBound(1) + 1
            End If
            Cont = Vetor.UpperBound(1)
            For x = 0 To Tam - 1
                subSetaVal Vetor, Cont, x, FunNulo(ResAvaliacao(VetorCampos(x)))
            Next
            ResAvaliacao.MoveNext
        Loop
        
    End If
    ResAvaliacao.Close
End Sub



Function ChecarCamposOrelhaPrincipal() As Boolean
    ChecarCamposOrelhaPrincipal = True
    If TabNotasFiscais.Tab = 0 Then
        If FunObrigatorioCBOX(CboStatus, "Status é Obrigatório") Then Exit Function
        If FunObrigatorioCBO(CboFornecedor, "Fornecedor é Obrigatório") Then Exit Function
        If CboStatus.ListIndex <> 0 Then
            If IsNull(DtpEmissao.Value) Then
                MsgBox "Data de Emissão é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
        End If
        If XDistribuido = "N" And ChkDistribuicao.Value = 0 Then
            If FunObrigatorioCBO(CboPcr, "Centro Custo é Obrigatório") Then Exit Function
            If FunObrigatorioMSK(Mskconta, "Conta Contábil é Obrigatória") Then Exit Function
        End If
        If Tipo = "1" Or CboStatus.ListIndex = 0 Then 'Gera duplicata automaticamente ou é previsto
            If IsNull(DtpVencParcelaUnica.Value) Then
                MsgBox "O preenchimento do campo Vencimento é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
            If CboStatus.ListIndex = 2 Then
                If IsNull(DtpPagtoParcelaUnica.Value) Then
                    MsgBox "Data do Pagamento é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
                    Exit Function
                End If
            End If
            If EntSaida = "E" Then
                If FunObrigatorioCBOX(CboTipoPagto1, "O preenchimento do campo Tipo Recebimento é obrigatório.") Then Exit Function
            Else
                If FunObrigatorioCBOX(CboTipoPagto1, "O preenchimento do campo Tipo Pagamento é obrigatório.") Then Exit Function
                'If CboTipoPagto1.ListIndex = 2 Or CboTipoPagto1.ListIndex = 3 Or CboTipoPagto1.ListIndex = 4 Then
                If CboTipoPagto1.ListIndex = 2 Then
                    If FunObrigatorioCBO(CboBancoForn1, "O preenchimento do Campo Banco do Fornecedor é obrigatório.") Then Exit Function
                End If
            End If
            DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
            If DatTipoDoc.Recordset.Fields("tido_tx_bancario") <> "0" Then
                'A nota não entra no saldo do banco, logo não possui conta
                If FunObrigatorioCBO(CboCCorrente, "O preenchimento do campo Conta Corrente é obrigatório.") Then Exit Function
            End If
        End If
        If FunObrigatorioTXT(TxtValor, "Valor é Obrigatório") Then
            Exit Function
        Else
            If TxtValor = "0,00" Then
                MsgBox "Não é permitido cadastrar um valor 'zero' para um documento.", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
        End If
        
    Else
        If FunObrigatorioCBO(CboFornecedor, "Fornecedor é Obrigatório") Then Exit Function
        If IsNull(DtpEmissao.Value) Then
            MsgBox "Data de Emissão é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
            Exit Function
        End If
        If FunObrigatorioTXT(TxtValor, "Valor é Obrigatório") Then Exit Function
    End If
    ChecarCamposOrelhaPrincipal = False
End Function

Function ChecarCampos() As Boolean
    ChecarCampos = True
    
    If TabNotasFiscais.Tab = 1 Then
        If FunObrigatorioTXT(TxtValorDuplicata, "O preenchimento do campo Valor é obrigatório.") Then Exit Function
        If EntSaida = "E" Then
            If FunObrigatorioCBOX(CboTipoPagto2, "O preenchimento do campo Tipo Recebimento é obrigatório.") Then Exit Function
        ElseIf FunObrigatorioCBOX(CboTipoPagto2, "O preenchimento do campo Tipo Pagamento é obrigatório.") Then Exit Function
        End If
        If FunObrigatorioCBO(CboCCorrente2, "O preenchimento do campo Conta Corrente é obrigatório.") Then Exit Function
        If CboStatusduplicata.ListIndex = 1 Then
            If IsNull(DtpPagamento.Value) Then
                MsgBox "Data do Pagamento é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
        End If
        If EntSaida = "S" Then
            'If CboTipoPagto2.ListIndex = 2 Or CboTipoPagto2.ListIndex = 3 Or CboTipoPagto2.ListIndex = 4 Then
            If CboTipoPagto2.ListIndex = 2 Then
                If FunObrigatorioCBO(CboBancoForn2, "O preenchimento do Campo Banco do Fornecedor é obrigatório.") Then Exit Function
            End If
        End If
    ElseIf TabNotasFiscais.Tab = 3 Then
        If FunObrigatorioTXT(TxtVldistribuicao, "O preenchimento do campo Valor é obrigatório.") Then Exit Function
        If FunObrigatorioCBO(CboPcr2, "O preenchimento do campo Centro Custo é obrigatório.") Then Exit Function
        If FunObrigatorioMSK(MskConta2, "O preenchimento do campo Centro Contábil é obrigatório.") Then Exit Function
    End If
    
    ChecarCampos = False
End Function

Sub PrepararInsercao()
    If TabNotasFiscais.Tab = 0 Then
        CboTipoDoc.Text = ""
        CboFornecedor.Text = ""
        MskNumdocumento.Text = ""
        TxtSerie.Text = ""
        DtpEntrada.Value = Format$(Now, "dd/mm/yy")
        DtpEmissao.Value = Format$(Now, "dd/mm/yy")
        TxtValor.Text = ""
        CboHistorico.Text = ""
        LblValorLiquido1.Caption = ""
        
        'TabAuxiliar (0)
        MskPcr.Mask = ""
        MskPcr.Text = ""
        MskPcr.Mask = pPCRmascara
        CboPcr.Text = ""
        Mskconta.Mask = ""
        Mskconta.Text = ""
        CboPlanoContas.Text = ""
        
        'TabAuxiliar(1)
        DtpVencParcelaUnica.Value = Null
        DtpPagtoParcelaUnica.Value = Null
        If Not TelaNotasFiscais.Adodc1.Recordset.EOF Then
            If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_TipoPag")) Then
                CboTipoPagto1.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_TipoPag")
            Else
                CboTipoPagto1.ListIndex = 0
            End If
        End If
        TxtDocPagto1.Text = ""
        CboBancoForn1.Text = ""
        CboCCorrente.Text = ""
                
   ElseIf TabNotasFiscais.Tab = 2 Then
        DtpVencDesconto.Value = Format(Now, "dd/mm/yy")
        TxtValorDesconto.Text = ""
        CboDescontos.Text = ""
        
   ElseIf TabNotasFiscais.Tab = 1 Then
        If Vetor2.UpperBound(1) > 0 Then
            TxtDuplicata.Text = Chr(66 + Vetor2.UpperBound(1))
        Else
            If Vetor2(0, 0) <> Empty Then
                TxtDuplicata.Text = "B"
            Else
                TxtDuplicata.Text = "A"
            End If
        End If
        TxtDuplicata.Enabled = False
        'TxtDuplicata.Text = ""
        If LblValorLiquido1.Caption <> "" And LblValorTotal.Caption <> "" Then
            TxtValorDuplicata.Text = Format(((LblValorLiquido1.Caption + CDbl(TxtDevolucao.Text)) - LblValorTotal.Caption), "standard")
        Else
            TxtValorDuplicata.Text = ""
        End If
        If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_TipoPag")) Then
            CboTipoPagto2.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_TipoPag")
        Else
            CboTipoPagto2.ListIndex = 0
        End If
        TxtDocPagto2.Text = ""
        CboCCorrente2.Text = ""
        CboStatusduplicata.ListIndex = 0
        CboBancoForn2.Text = ""
        DtpVencimento.Value = Format(Now, "dd/mm/yy")
        DtpPagamento.Value = Null
        
    ElseIf TabNotasFiscais.Tab = 3 Then
        TxtPercDistribuicao.Text = ""
        If LblValorNota.Caption <> "" Then
            If LblValorTotalDistrib.Caption = "" Then
                LblValorTotalDistrib.Caption = "0"
            End If
            TxtVldistribuicao.Text = Format(CStr(CDbl(LblValorNota.Caption) - CDbl(LblValorTotalDistrib.Caption)), "standard")
            Call TxtVldistribuicao_LostFocus
        End If
        MskPcr2.Mask = ""
        MskPcr2.Text = ""
        MskPcr2.Mask = pPCRmascara
        CboPcr2.Text = ""
        MskConta2.Mask = ""
        MskConta2.Text = ""
        CboPlanoContas2.Text = ""
    End If
End Sub

Private Sub CboBancoForn1_Change()
    If CboBancoForn1.Text <> "" Then
        DatBancoFornecedor.Recordset.Bookmark = CboBancoForn1.SelectedItem
    End If
End Sub

Private Sub CboBancoForn2_Change()
    If CboBancoForn2.Text <> "" Then
        DatBancoFornecedor.Recordset.Bookmark = CboBancoForn2.SelectedItem
    End If
End Sub

Private Sub CboCCorrente_Change()
    If CboCCorrente.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    End If
End Sub

Private Sub CboCCorrente2_Change()
    If CboCCorrente2.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente2.SelectedItem
    End If
End Sub

Private Sub CboDescontos_Change()
    'If CboDescontos.Text <> "" And EntSaida = "S" Then
    'TESTE
    If CboDescontos.Text <> "" And EntSaida = "S" And EntrouDupl Then
        DatDescontosNF.Recordset.Bookmark = CboDescontos.SelectedItem
        If DatDescontosNF.Recordset.Fields("desc_tx_tipobase") = "2" Then
            DtpVencDesconto.Value = Format$(Vetor2(0, 1), "DD/MM/YY")
        ElseIf Not IsNull(DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")) Then
                Dim Ano As Integer
                Dim Mes As Integer
                
                Ano = Year(CDate(Vetor2(0, 1)))
                Mes = Month(CDate(Vetor2(0, 1)))
                If Mes = 12 Then
                   Mes = 1
                Else
                   Mes = Mes + 1
                End If
                DtpVencDesconto.Value = Format$(DateSerial(Ano, Mes, DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")), "dd/mm/yy")
            Else
                DtpVencDesconto.Value = Format(Now, "dd/mm/yy")
        End If
    End If
End Sub

Private Sub CboFornecedor_Change()
    If CboFornecedor.Text <> "" Then
        CboBancoForn1.Text = ""
        CboTipoPagto1.ListIndex = 0
        DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
        If Not IsNull(DatFornecedor.Recordset.Fields("cere_cd_estruturado")) Then
            'DatPcr.Recordset.Find "cere_cd_pcr =" & DatFornecedor.Recordset.Fields("cere_cd_pcr")
            'CboPcr.BoundText = DatPcr.Recordset.Fields("cere_cd_estruturado")
            CboPcr.BoundText = DatFornecedor.Recordset.Fields("cere_cd_estruturado")
        End If
        If Not IsNull(DatFornecedor.Recordset.Fields("CodRedRecDesp")) Then
            CboPlanoContas.BoundText = DatFornecedor.Recordset.Fields("CodRedRecDesp")
        End If
        If Tipo = "1" And DatFornecedor.Recordset.Fields("focl_tx_classe") = "F" Then
            CboTipoPagto1.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_tipopag")
        ElseIf Tipo = "0" And DatFornecedor.Recordset.Fields("focl_tx_classe") = "F" Then
            CboTipoPagto2.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_tipopag")
        End If
    End If
End Sub

Private Sub CboFornecedor_Click(Area As Integer)
    If CboFornecedor.Text <> "" And EntSaida = "S" Then
        subConectarControleDados DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
    End If
End Sub

Private Sub CboFornecedor_DblClick(Area As Integer)
    Dim ChaveTemp As Variant
    
    ChaveTemp = Chave
        
    If EntSaida = "E" Then
        FornClie = "C" 'Cliente
    Else
        FornClie = "F"  'Fornecedor
    End If
    Pessoa = "J" 'Pessoa Jurídica
    
    XFormForn = "FrmNotasFiscais"
    FrmFornecedores.Show 1
    Chave = ChaveTemp
    Set Formulario = FrmNotasFiscais
    FrmNotasFiscais.DatFornecedor.Refresh
    FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
    
'    PanFornecedor.Visible = False
'    PanFornecedor.Enabled = False

'    PanFornecedor.Visible = True
'    PanFornecedor.Enabled = True
'    If EntSaida = "E" Then
'        OptProduto.Visible = False
'        OptProdServ.Visible = False
'        OptJuridica.Top = 285
'        OptFisica.Top = 660
'    End If
End Sub



Private Sub CboFornecedor_LostFocus()
    If CboFornecedor.Text <> "" Then
        MskNumdocumento.Enabled = True
        MskNumdocumento.SetFocus
    End If
End Sub

Private Sub CboPcr_Change()
    If CboPcr.Text <> "" Then
        DatPcr.Recordset.Bookmark = CboPcr.SelectedItem
        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            Exit Sub
        End If
        MskPcr.Text = CboPcr.BoundText
    End If
End Sub

Private Sub CboPcr2_Change()
    If CboPcr2.Text <> "" Then
        DatPcr.Recordset.Bookmark = CboPcr2.SelectedItem
        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr2.BoundText = MskPcr2.Text
            CboPcr2.SetFocus
            Exit Sub
        End If
        MskPcr2.Text = CboPcr2.BoundText
    End If
End Sub
Private Sub CboPlanoContas_Change()
    If CboPlanoContas.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = Mskconta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        Mskconta.Text = CboPlanoContas.BoundText
    End If
End Sub

Private Sub CboPlanoContas2_Change()
    If CboPlanoContas2.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas2.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas2.BoundText = MskConta2.Text
            CboPlanoContas2.SetFocus
            Exit Sub
        End If
        MskConta2.Text = CboPlanoContas2.BoundText
    End If
End Sub

Private Sub CboStatus_Change()
    If CboStatus.Text <> "" Then
        HabilitaControles
        ConsistenciaStatus
    End If
End Sub

Private Sub CboStatus_Click()
    If CboTipoDoc.Enabled = False Then
        CboTipoDoc.Enabled = True
    Else
        HabilitaControles
    End If
End Sub

Private Sub CboStatusduplicata_click()
    If CboStatusduplicata.ListIndex = 1 Then
        TxtDocPagto2.Enabled = True
        DtpPagamento.Enabled = True
    Else
        TxtDocPagto2.Text = ""
        TxtDocPagto2.Enabled = False
        DtpPagamento.Value = Null
        DtpPagamento.Enabled = False
    End If
End Sub

Private Sub CboTipoDoc_LostFocus()
    Dim ResDoc As Object
    If Not IsNull(CboTipoDoc.SelectedItem) Then
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        If Not IsNull(DatTipoDoc.Recordset.Fields("tido_tx_duplicata")) Then
            Tipo = DatTipoDoc.Recordset.Fields("tido_tx_duplicata")
        End If
        HabilitaControles
        CboFornecedor.SetFocus
    End If
End Sub

Private Sub CboTipoPagto1_CLICK()
    Dim Achou As Boolean
    
    Achou = False
    If EntSaida = "S" Then
        'If CboTipoPagto1.ListIndex = 2 Or CboTipoPagto1.ListIndex = 3 Or CboTipoPagto1.ListIndex = 4 Then
        If CboTipoPagto1.ListIndex = 2 Then
            'trazer o banco padrão
            If Not (DatBancoFornecedor.Recordset.BOF And DatBancoFornecedor.Recordset.EOF) Then
                DatBancoFornecedor.Recordset.MoveFirst
                While Not DatBancoFornecedor.Recordset.EOF And Achou = False
                    If DatBancoFornecedor.Recordset.Fields("bafo_tx_padrao") = "N" Then
                        DatBancoFornecedor.Recordset.MoveNext
                    ElseIf DatBancoFornecedor.Recordset.Fields("bafo_tx_padrao") = "S" Then
                        CboBancoForn1.BoundText = DatBancoFornecedor.Recordset.Fields("bafo_cd_codigo")
                        Achou = True
                    End If
                Wend
            End If
            CboBancoForn1.Enabled = True
        Else
            CboBancoForn1.Text = ""
            CboBancoForn1.Enabled = False
        End If
    End If
End Sub

Private Sub CboTipoPagto2_click()
    Dim Achou As Boolean
    
    Achou = False
    If EntSaida = "S" Then
        'If CboTipoPagto2.ListIndex = 2 Or CboTipoPagto2.ListIndex = 3 Or CboTipoPagto2.ListIndex = 4 Then
        If CboTipoPagto2.ListIndex = 2 Then
            'trazer o banco padrão
            If Not (DatBancoFornecedor.Recordset.BOF And DatBancoFornecedor.Recordset.EOF) Then
                DatBancoFornecedor.Recordset.MoveFirst
                While Not DatBancoFornecedor.Recordset.EOF And Achou = False
                    If DatBancoFornecedor.Recordset.Fields("bafo_tx_padrao") = "N" Then
                        DatBancoFornecedor.Recordset.MoveNext
                    ElseIf DatBancoFornecedor.Recordset.Fields("bafo_tx_padrao") = "S" Then
                        CboBancoForn2.BoundText = DatBancoFornecedor.Recordset.Fields("bafo_cd_codigo")
                        Achou = True
                    End If
                Wend
                'CboBancoForn2.BoundText = DatBancoFornecedor.Recordset.Fields("bafo_cd_codigo")
            End If
            CboBancoForn2.Enabled = True
        Else
            CboBancoForn2.Text = ""
            CboBancoForn2.Enabled = False
        End If
    End If
End Sub

Private Sub ChkDistribuicao_Click()
    'If XInserir = "I" Then
        If ChkDistribuicao.Value = 1 Then
           MskPcr.Mask = ""
           MskPcr.Text = ""
           MskPcr.Mask = pPCRmascara
           Mskconta.Text = ""
           CboPlanoContas.Text = ""
           CboPcr.Text = ""
           TabAuxiliar.TabEnabled(0) = False
           MskPcr.Enabled = False
           CboPcr.Enabled = False
           Mskconta.Enabled = False
           CboPlanoContas.Enabled = False
           TabNotasFiscais.TabEnabled(3) = True
        Else
           Mskconta.Text = ""
           CboPlanoContas.Text = ""
           CboPcr.Text = ""
           TabAuxiliar.TabEnabled(0) = True
           MskPcr.Enabled = True
           CboPcr.Enabled = True
           Mskconta.Enabled = True
           CboPlanoContas.Enabled = True
           TabNotasFiscais.TabEnabled(3) = False
           TabNotasFiscais.Tab = 0
        End If
        XDistribuido = "N"
    'End If
End Sub


Private Sub CmdAlteraDesconto_Click()
    Dim Xvalor As Double
                    
    FrmNotasFiscais.MousePointer = vbHourglass
    'If ExisteChave(1, Array("4"), Array(Fatura), Vetor1, TDBGrid1) Then Exit Sub
    ChaveD = -1
    XDesconto = "A"
    
    PrepararInsercao
    CboDescontos.BoundText = TDBGrid1.Columns("codigo").Text
    xcod = CboDescontos.BoundText
    If TDBGrid1.Columns("Vencimento") <> Null Or TDBGrid1.Columns("Vencimento") <> "" Then
        DtpVencDesconto.Value = Format(TDBGrid1.Columns("Vencimento").Text, "dd/mm/yy")
    End If
    
    TxtValorDesconto.Text = Format(TDBGrid1.Columns("Valor").Text, "standard")
    
    pandescontos.Visible = True
    pandescontos.Enabled = True
    CboDescontos.SetFocus
    DesabilitaBotoes
    FrmNotasFiscais.MousePointer = vbDefault
    
End Sub

Private Sub CmdAlteraDistribuicao_Click()
    
    'On Error GoTo TrataErro
    
    DesabilitaBotoes
    TDBGrid3.Enabled = False
    
    PanDistribuicao.Visible = True
    PanDistribuicao.Enabled = True
    XDistribuicao = "A"
    
    MskPcr2.Text = TDBGrid3.Columns(1).Text
    CboPcr2.BoundText = MskPcr2.Text
    XPcr = CboPcr2.BoundText
    MskConta2.Text = TDBGrid3.Columns(3).Text
    CboPlanoContas2.BoundText = MskConta2.Text
    xcc = MskConta2.Text
    TxtVldistribuicao.Text = Format(CDbl(TDBGrid3.Columns(4).Text), "standard")
    TxtPercDistribuicao.Text = Mid(TDBGrid3.Columns(7).Text, 1, (Len(TDBGrid3.Columns(7).Text) - 1))
    DesabilitaBotoes
    MskPcr2.SetFocus
    
Exit Sub

TrataErro:

  funTrataErros (ComMensagem)

End Sub

Private Sub CmdAlteraduplicata_Click()
    Dim Status As String
    Dim x As Integer
    
    DesabilitaBotoes
    TDBGrid1.Enabled = False
    
    PnlCompDuplicata.Visible = True
    PnlCompDuplicata.Enabled = True
    
    XDuplicata = "A"
    
    TxtDuplicata.Text = TDBGrid2.Columns("Código").Text
    DtpVencimento.Value = Format$(TDBGrid2.Columns("Vencimento").Text, "DD/MM/YY")
    TxtValorDuplicata.Text = Format$(TDBGrid2.Columns("Valor"), "standard")
    If TDBGrid2.Columns("CodStatus") = "C" Then 'Compromissado
           Status = "0"
    ElseIf TDBGrid2.Columns("CodStatus") = "R" Then 'Realizado
           Status = "1"
    ElseIf TDBGrid2.Columns("CodStatus") = "P" Then 'Previsto(qdo altera o status de uma nota Prevista)
        If CboStatus.Text = "Compromissado" Then
           Status = "0"
        ElseIf CboStatus.Text = "Realizado" Then
             Status = "1"
        End If
    End If
    CboStatusduplicata.ListIndex = Status
    If TDBGrid2.Columns("CodTipoPagto") <> "" Then
       For x = 0 To CboTipoPagto2.ListCount - 1
           If CboTipoPagto2.ItemData(x) = CInt(TDBGrid2.Columns("CodTipoPagto")) Then
               CboTipoPagto2.ListIndex = x
           End If
       Next
    Else
       CboTipoPagto2.ListIndex = ""
    End If
    If EntSaida = "S" Then
        TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Pag."))
        If TDBGrid2.Columns("Data Pag.") <> Null Or TDBGrid2.Columns("Data Pag.") <> "" Then
            DtpPagamento.Enabled = True
            DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Pag."), "DD/MM/YY")
        End If
    Else
        TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Rec."))
        If TDBGrid2.Columns("Data Rec.") <> Null Or TDBGrid2.Columns("Data Rec.") <> "" Then
            DtpPagamento.Enabled = True
            DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Rec."), "DD/MM/YY")
        End If
    End If
    CboCCorrente2.BoundText = TDBGrid2.Columns("CodConta")
    CboBancoForn2.BoundText = TDBGrid2.Columns("CodBancoForn")
    XVlDuplicata = CDbl(TxtValorDuplicata.Text)
    
    If CboStatus.Text = "Realizado" Then
        CboStatusduplicata.Enabled = False
        CboStatusduplicata.ListIndex = 1
    Else
        CboStatusduplicata.Enabled = True
    End If

    If TDBGrid2.Columns("CodStatus") = "R" Then
        DesabilitaRealizado "Duplicata", False
    Else
        DesabilitaRealizado "Duplicata", True
    End If
End Sub

Private Sub CmdConfEntMult_Click()
    If Not IsNumeric(TxtNParcelas.Text) Then
        MsgBox "O número de parcelas tem que ser definido.", vbCritical, "ATENÇÃO"
        TxtNParcelas.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TxtPeriodicidade.Text) Then
        MsgBox "A periodicidade tem que ser definida.", vbCritical, "ATENÇÃO"
        TxtPeriodicidade.SetFocus
        Exit Sub
    End If
    
    If XFT_NOTADUPLICATA = "N" Then
        ParcelasMultiplas = True
        XFT_INSPARCMULTIPLA = "A"
    Else
        Dim XLI_X As Integer
        Dim XLI_Y As Integer
        Dim XLD_VENCIMENTO As Date
        Dim XLT_PARCELA As String
    
        XFT_INSDUPLMULTIPLA = "A"
            
        If XDuplicata = "A" Then
           ValorTotalDuplicata = ValorTotalDuplicata - XVlDuplicata
        End If
    
        If CDbl(LblValorLiquido1.Caption) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
            MsgBox "A soma das duplicatas excedeu o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
            TxtValorDuplicata.SetFocus
            CmdInsereduplicata_Click
        Else
            For XLI_Y = 1 To Val(TxtNParcelas.Text)    'Grava Duplicatas múltiplas
                If Vetor2.UpperBound(1) > 0 Then
                    XLT_PARCELA = Chr(66 + Vetor2.UpperBound(1))
                Else
                    If Vetor2(0, 0) <> Empty Then
                        XLT_PARCELA = "B"
                    Else
                        XLT_PARCELA = "A"
                    End If
                End If
                XLI_X = Vetor2.UpperBound(1)
                If OptMes.Value Then
                    XLD_VENCIMENTO = DateAdd("m", CInt(TxtPeriodicidade.Text), Vetor2((XLI_Y - 1), 1))
                ElseIf OptDia.Value Then
                    XLD_VENCIMENTO = DateAdd("d", CInt(TxtPeriodicidade.Text), Vetor2((XLI_Y - 1), 1))
                End If
                InsereRegistroVetorDupl Array(XLT_PARCELA, Format(XLD_VENCIMENTO, "dd/mm/yyyy"), Vetor2(XLI_X, 2), 0, Vetor2(XLI_X, 2), "C", "Compromissado", Vetor2(XLI_X, 7), _
                Vetor2(XLI_X, 8), "", "", Vetor2(XLI_X, 11), Vetor2(XLI_X, 12), Vetor2(XLI_X, 13), Vetor2(XLI_X, 14), Vetor2(XLI_X, 15)), Vetor2, TDBGrid2
                        
                If Vetor2(0, 0) <> Empty Then
                    HabilitaBotoes
                    LblValorTotal.Caption = funSomaColuna(Vetor2, 2)
                    ValorTotalDuplicata = CDbl(LblValorTotal.Caption)
                Else
                    CmdInsereduplicata.Enabled = True
                End If
            HabilitaBotoes
            Next
            TDBGrid2.Enabled = True
        End If
        PanEntMultiplas.Visible = False
        TabNotasFiscais.TabEnabled(2) = True
        TabNotasFiscais.Enabled = True
        CmdDuplMultiplas.Enabled = False
    
ResumoDoErro:
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
'    Conexao.RollbackTrans
    Resume ResumoDoErro
    End If
    PanEntMultiplas.Enabled = False
    PanEntMultiplas.Visible = False
    TabNotasFiscais.Enabled = True
    CmdGravar.Enabled = True
    CmdDesistir.Enabled = True
    
End Sub

Private Sub CmdConfirmaDesconto_Click()
    If FunObrigatorioCBO(CboDescontos, "É necessário selecionar um desconto!") Then
        CboDescontos.SetFocus
        Exit Sub
    End If
    
    If FunObrigatorioTXT(TxtValorDesconto, "Valor do desconto é obrigatório!") Then
        TxtValorDesconto.SetFocus
        Exit Sub
    End If
    
    If (xcod = CboDescontos.BoundText) Then
        If XDesconto = "I" Then
           subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, ""), Vetor1, TDBGrid1
        ElseIf XDesconto = "A" Then
            TDBGrid1.Columns("Descrição").Text = CboDescontos.Text
            TDBGrid1.Columns("Vencimento").Text = Format(DtpVencDesconto.Value, "dd/mm/yyyy")
            TDBGrid1.Columns("Valor").Text = TxtValorDesconto.Text
            TDBGrid1.Columns("codigo").Text = CboDescontos.BoundText
            TDBGrid1.Update
        End If
    Else
        DatDescontosNF.Recordset.Bookmark = CboDescontos.SelectedItem
        If funExisteChave(1, Array(3), Array(CboDescontos.BoundText), Vetor1, TDBGrid1) Then
            MsgBox "Esse imposto já foi cadastrado", vbCritical, "ATENÇÃO!"
            If XDesconto = "I" Then
               CboDescontos.SetFocus
               CmdInsereDesconto_Click
            Else
                CboDescontos.SetFocus
                Exit Sub
            End If
        Else
            If XDesconto = "I" Then
                subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, ""), Vetor1, TDBGrid1
            Else
                TDBGrid1.Columns("Descrição").Text = CboDescontos.Text
                TDBGrid1.Columns("Vencimento").Text = Format(DtpVencDesconto.Value, "dd/mm/yyyy")
                TDBGrid1.Columns("Valor").Text = TxtValorDesconto.Text
                TDBGrid1.Columns("codigo").Text = CboDescontos.BoundText
                TDBGrid1.Update
            End If
        End If
    End If
  
    pandescontos.Visible = False
    pandescontos.Enabled = False
    
    If Vetor1(0, 0) <> Empty Then
       HabilitaBotoes
    Else
        CmdInsereDesconto.Enabled = True
    End If
    
    TDBGrid1.Enabled = True
   
    LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")

End Sub

Private Sub botao_conf_Click()
    Dim XLF_VALORNOTA As Double, XLF_VALORDISTRIB As Double
    Dim Xvalor

    'On Error GoTo TrataErro

    'Função para verificar preenchimento dos campos obrigatórios
    If ChecarCampos Then Exit Sub
    
    If XDistribuicao = "A" Then
       ValorTotalDistribuicao = ValorTotalDistribuicao - CDbl(TxtVldistribuicao.Text)
    End If
    XLF_VALORNOTA = CDbl(LblValorNota.Caption)
    XLF_VALORDISTRIB = (ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text))
    If CDbl(Format(((CDbl(LblValorNota.Caption) * 100) / 100), "##,##0.00")) < CDbl(Format((((ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text)) * 100) / 100), "##,##0.00")) Then
    'If ((CDbl(LblValorNota.Caption) * 100) / 100) < (((ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text)) * 100) / 100) Then
    'If Format(XLF_VALORNOTA, "STANDARD") < Format(XLF_VALORDISTRIB, "STANDARD") Then
        MsgBox "As distribuições excederam o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        TxtVldistribuicao.SetFocus
        Exit Sub
    Else
        If (XPcr = MskPcr2.Text) And (xcc = MskConta2.Text) Then
            If XDistribuicao = "I" Then
                'InsereRegistroVetor Array(DatPcr.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, MskConta2.Text, TxtVldistribuicao.Text, (TxtPercDistribuicao.Text / 100)), Vetor3, TDBGrid3
                subInsereRegistroVetor Array(DatPcr.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, DatConta.Recordset.Fields("plco_cd_conta"), MskConta2.Text, TxtVldistribuicao.Text, 0, TxtVldistribuicao.Text, (TxtPercDistribuicao.Text / 100)), Vetor3, TDBGrid3
            ElseIf XDistribuicao = "A" Then
                    TDBGrid3.Columns(0).Text = DatPcr.Recordset.Fields("cere_cd_pcr")
                    TDBGrid3.Columns(1).Text = MskPcr2.Text
                    TDBGrid3.Columns(2).Text = DatConta.Recordset.Fields("plco_cd_conta")
                    TDBGrid3.Columns(3).Text = MskConta2.Text
                    TDBGrid3.Columns(4).Text = TxtVldistribuicao.Text
                    TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                    TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                    TDBGrid3.Update
            End If
        Else
            If ExisteDistribuicao(2, Array(0, 2), Array(DatPcr.Recordset.Fields("cere_cd_pcr"), DatConta.Recordset.Fields("plco_cd_conta")), Vetor3, TDBGrid3) Then
               MsgBox "Distribuição já Cadastrada", vbCritical, "ATENÇÃO"
               If XDistribuicao = "I" Then
                  TxtVldistribuicao.SetFocus
                  CmdInsereDistribuicao_Click
               ElseIf XDistribuicao = "A" Then
                    ValorTotalDistribuicao = ValorTotalDistribuicao + CDbl(TxtVldistribuicao)
                    TabNotasFiscais.Enabled = True
                    Exit Sub
               End If
            Else
                If XDistribuicao = "I" Then
                   subInsereRegistroVetor Array(DatPcr.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, DatConta.Recordset.Fields("plco_cd_conta"), MskConta2.Text, TxtVldistribuicao.Text, 0, TxtVldistribuicao.Text, (TxtPercDistribuicao.Text / 100)), Vetor3, TDBGrid3
                ElseIf XDistribuicao = "A" Then
                        TDBGrid3.Columns(0).Text = DatPcr.Recordset.Fields("cere_cd_pcr")
                        TDBGrid3.Columns(1).Text = MskPcr2.Text
                        TDBGrid3.Columns(2).Text = DatConta.Recordset.Fields("plco_cd_conta")
                        TDBGrid3.Columns(3).Text = MskConta2.Text
                        TDBGrid3.Columns(4).Text = TxtVldistribuicao.Text
                        TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                        TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                        TDBGrid3.Update
                End If
            End If
        End If
    End If
    PanDistribuicao.Enabled = False
    PanDistribuicao.Visible = False
    If Vetor3(0, 0) <> Empty Then
       HabilitaBotoes
       LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
       ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    Else
        LblValorTotalDistrib.Caption = ""
        ValorTotalDistribuicao = 0
        CmdInsereDistribuicao.Enabled = True
    End If
    
    TDBGrid3.Enabled = True
    
    If XInserir = "I" Or XDistribuido = "N" Then
        If LblValorNota.Caption = CDbl(LblValorTotalDistrib.Caption) Then
        'If LblValorNota.Caption = CDbl(LblValorTotalDistrib.Caption) Then
           XDistribuido = "S"
        Else
           XDistribuido = "N"
        End If
    End If
End Sub

Private Sub CmdConfirmaduplicata_Click()
    Dim XStatus As String
    Dim Xvalor As Double
    Dim xcod As String
    Dim XLI_ANO As Integer, XLI_MES As Integer, XLI_X As Integer
    Dim XLO_DESC As Object

    'On Error GoTo TrataErro

    'Função para verificar preenchimento dos campos obrigatórios
    If ChecarCampos Then Exit Sub
    If EntSaida = "E" Then
        Xdebcred = "C"
    ElseIf EntSaida = "S" Then
        Xdebcred = "D"
    End If
        
    If XDuplicata = "A" Then
       ValorTotalDuplicata = ValorTotalDuplicata - XVlDuplicata
    End If

    If (CDbl(LblValorLiquido1.Caption) + CDbl(TxtDevolucao.Text)) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
        MsgBox "A soma das duplicatas excedeu o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        TxtValorDuplicata.SetFocus
        CmdInsereduplicata_Click
    Else
        If TxtDuplicata.Text = "" Then
            xcod = "U"
        Else
            xcod = TxtDuplicata.Text
        End If
    
        If CboStatusduplicata.ListIndex = 0 Then 'Compromissado
           XStatus = "C"
        ElseIf CboStatusduplicata.ListIndex = 1 Then 'Realizado
           XStatus = "R"
        End If

        'XStatus = VerificaStatus(CboStatusduplicata.ListIndex)
 
        If XDuplicata = "I" Then
            InsereRegistroVetorDupl Array(xcod, Format(DtpVencimento.Value, "dd/mm/yyyy"), TxtValorDuplicata.Text, 0, TxtValorDuplicata.Text, XStatus, CboStatusduplicata.Text, _
            CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex), CboTipoPagto2.Text, TxtDocPagto2.Text, Format(DtpPagamento.Value, "dd/mm/yyyy"), CboCCorrente2.BoundText, CboCCorrente2.Text, FunNuloValor(CboBancoForn2), CboBancoForn2.Text), Vetor2, TDBGrid2
            CmdDuplMultiplas.Enabled = True
        ElseIf XDuplicata = "A" Then
            TDBGrid2.Columns("Código").Text = TxtDuplicata.Text
            TDBGrid2.Columns("Vencimento").Text = Format(DtpVencimento.Value, "dd/mm/yyyy")
            TDBGrid2.Columns("Valor").Text = TxtValorDuplicata.Text
            TDBGrid2.Columns("CodStatus").Text = XStatus
            TDBGrid2.Columns("Status").Text = CboStatusduplicata.Text
            TDBGrid2.Columns("CodTipoPagto").Text = CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex)
            If EntSaida = "S" Then
                TDBGrid2.Columns("Tipo Pag.").Text = CboTipoPagto2.Text
                TDBGrid2.Columns("Doc. Pag.").Text = TxtDocPagto2.Text
                TDBGrid2.Columns("Data Pag.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
                TDBGrid2.Columns("CodBancoForn").Text = FunNuloValor(CboBancoForn2)
                TDBGrid2.Columns("Conta Corrente (Fornecedor)").Text = CboBancoForn2.Text
                TDBGrid2.Columns("Valor Líquido").Text = TxtValorDuplicata.Text - TDBGrid2.Columns(3).Text
            Else
                TDBGrid2.Columns("Tipo Rec.").Text = CboTipoPagto2.Text
                TDBGrid2.Columns("Doc. Rec.").Text = TxtDocPagto2.Text
                TDBGrid2.Columns("Data Rec.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
                TDBGrid2.Columns("Valor Líquido").Text = TxtValorDuplicata.Text - TDBGrid2.Columns(3).Text
            End If
            TDBGrid2.Columns("CodConta").Text = CboCCorrente2.BoundText
            TDBGrid2.Columns("Conta Corrente (Empresa)").Text = CboCCorrente2.Text
            TDBGrid2.Update
        End If
        PnlCompDuplicata.Visible = False
        PnlCompDuplicata.Enabled = False
            
        If Vetor2(0, 0) <> Empty Then
            HabilitaBotoes
            LblValorTotal.Caption = funSomaColuna(Vetor2, 2)
            ValorTotalDuplicata = CDbl(LblValorTotal.Caption)
        Else
            CmdInsereduplicata.Enabled = True
        End If
        TDBGrid2.Enabled = True
    End If
    TabNotasFiscais.TabEnabled(2) = True
    
    If Vetor2.UpperBound(1) = 0 And EntrouDesc And XDuplicata = "I" And EntSaida = "S" Then
         TDBGrid1.MoveFirst
         For XLI_X = 0 To Vetor1.UpperBound(1)
            If Vetor1(XLI_X, 1) = Empty Then
               SubQOpenRecordset XLO_DESC, "SELECT * FROM DescontosImpostos WHERE desc_cd_desconto=" & Vetor1(XLI_X, 3), "Estatico"
               If XLO_DESC!desc_tx_tipobase = "1" Or XLO_DESC!desc_tx_tipobase = "3" Then
                   XLI_ANO = Year(CDate(Vetor2(0, 1)))
                   XLI_MES = Month(CDate(Vetor2(0, 1)))
                   If XLI_MES = 12 Then
                       XLI_MES = 1
                   Else
                       XLI_MES = XLI_MES + 1
                   End If
                   'Vetor1(XLI_X, 2) = DateSerial(XLI_ANO, XLI_MES, XLO_DESC!desc_tx_DiaRecolhimento)
                   TDBGrid1.Columns("Vencimento").Text = DateSerial(XLI_ANO, XLI_MES, XLO_DESC!desc_tx_DiaRecolhimento)
               Else
                   'Vetor1(XLI_X, 2) = Vetor2(0, 1)
                   TDBGrid1.Columns("Vencimento").Text = Vetor2(0, 1)
               End If
               XLO_DESC.Close
            End If
            TDBGrid1.MoveNext
        Next
    End If
    
ResumoDoErro:
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
'    Conexao.RollbackTrans
    Resume ResumoDoErro
End Sub

Private Sub CmdDesiste_Click()
    ParcelasMultiplas = False
    PanEntMultiplas.Enabled = False
    PanEntMultiplas.Visible = False
    TabNotasFiscais.Enabled = True
    CmdGravar.Enabled = True
    CmdDesistir.Enabled = True
    If XFT_NOTADUPLICATA = "N" Then
        XFT_INSPARCMULTIPLA = "I"
    Else
        XFT_INSDUPLMULTIPLA = "I"
    End If
End Sub

Private Sub CmdDesistir_Click()
    If XFormulario = "TelaNotasFiscais" Then
        If Not FunTabelaVazia(TelaNotasFiscais.Adodc1) Then
          'TelaNotasFiscais.Adodc1.Recordset.CancelUpdate
          TelaNotasFiscais.CmdAlterar.Enabled = True
          TelaNotasFiscais.CmdExcluir.Enabled = True
          TelaNotasFiscais.CmdImprimir.Enabled = True
          TelaNotasFiscais.CmdOrdem.Enabled = True
          TelaNotasFiscais.CmdDuplicata.Enabled = True
          'TelaNotasFiscais.CmdProcPagto.Enabled = True
          TelaNotasFiscais.CmdRecibo.Enabled = True
          TelaNotasFiscais.CmdSair.Enabled = True
        End If
    ElseIf XFormulario = "TelaDuplicatas" Then
        TelaDuplicatas.Adodc1.Recordset.CancelUpdate
    ElseIf XFormulario = "TelaFatura" Then
        TelaFatura.Adodc1.Recordset.CancelUpdate
    End If
    Unload Me
End Sub

Private Sub CmdDuplMultiplas_Click()
    PanEntMultiplas.Visible = True
    PanEntMultiplas.Enabled = True
    If XFT_INSDUPLMULTIPLA = "I" Then
        TxtNParcelas.Text = ""
        TxtNParcelas.SetFocus
        TxtPeriodicidade.Text = ""
        CmdDesiste.Visible = False
    Else
        CmdDesiste.Visible = True
    End If
    TabNotasFiscais.Enabled = False
    'CmdGravar.Enabled = False
    'CmdDesistir.Enabled = False
    XFT_NOTADUPLICATA = "D" ' Gerar várias duplicatas
End Sub

Private Sub CmdEntMultiplas_Click()
    If ChecarCamposOrelhaPrincipal Then Exit Sub
    PanEntMultiplas.Visible = True
    PanEntMultiplas.Enabled = True
    'CboEntMultiplas.SetFocus
    If XFT_INSPARCMULTIPLA = "I" Then
        TxtNParcelas.Text = ""
        TxtNParcelas.SetFocus
        TxtPeriodicidade.Text = ""
        CmdDesiste.Visible = False
    Else
        CmdDesiste.Visible = True
    End If
    TabNotasFiscais.Enabled = False
    CmdGravar.Enabled = False
    CmdDesistir.Enabled = False
    XFT_NOTADUPLICATA = "N" ' Gerar várias notas fiscais
    
End Sub

Private Sub CmdExcluiDesconto_Click()
   Dim Houve_Erro As Integer

   Houve_Erro = False
       
   DesabilitaBotoes
   TDBGrid1.Enabled = False
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
      subRemoveRegistroVetor Vetor1, TDBGrid1
      If Vetor1(0, 0) <> Empty Then
        HabilitaBotoes
        LblValorLiquido1.Caption = (TxtValor.Text - (funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)))
      Else
        CmdInsereDesconto.Enabled = True
        LblValorLiquido1.Caption = (TxtValor.Text - CDbl(TxtDevolucao.Text))
      End If
      TDBGrid1.Enabled = True
   Else
      HabilitaBotoes
   End If
   Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub CmdExcluiduplicata_Click()
   Dim Houve_Erro As Integer
   Dim ValorExcluido As Double

   Houve_Erro = False
    
   DesabilitaBotoes
   
   TDBGrid2.Enabled = False
   
   If TDBGrid2.Columns("Status") = "Realizado" Then
        MsgBox "A duplicata não pode ser apagada porque já foi realizada.", vbCritical + vbOKOnly, "ATENÇÃO"
        HabilitaBotoes
        Exit Sub
   End If
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        ValorExcluido = TDBGrid2.Columns("Valor").Value
        subRemoveRegistroVetor Vetor2, TDBGrid2
   End If
   
   If Vetor2(0, 0) <> Empty Then
       HabilitaBotoes
   Else
       CmdInsereduplicata.Enabled = True
   End If
   
   LblValorTotal.Caption = Format(ValorTotalDuplicata - ValorExcluido, "standard")
   ValorTotalDuplicata = LblValorTotal.Caption
   TDBGrid2.Enabled = True
   
   Exit Sub

RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub CmdGravar_Click()
    Dim Sql As String
    Dim XStatus As String
    Dim XSerie As String
    Dim ResNF As Object
    Dim x As Byte
    Dim XNumDoc As String
    Dim XCodDupl As String
    Dim ResFat As Object
    Dim ResFormulario As Object
    Dim DataVencimento As Date

    'On Error GoTo TrataErro
    
    'If CmdGravar.Caption = "&Gravar" Then
        TabNotasFiscais.Tab = 0
        If ChecarCamposOrelhaPrincipal Then Exit Sub
        
        If XDistribuido = "N" And (MskPcr.Text = "" And Mskconta.Text = "") Then
            MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf ChkDistribuicao.Value = 1 Then
            If TxtValor.Text <> funSomaColuna(Vetor3, 4) Then
                MsgBox "A soma das distribuições é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                Exit Sub
            End If
        End If
        
        If CboStatus.ListIndex <> 0 And Tipo = 0 And XTipoFatura <> "3" Then 'Não é previsto e não gera duplicata automaticamente.
            'Verifica se o total da nota é igual a soma das duplicatas com os descontos ou se alterou o valor da nota
            If VerificaValores Then
                Exit Sub
            End If
        End If
        
        'If ChecarCamposOrelhaPrincipal Then Exit Sub
        XStatus = funVerificaStatus(CboStatus.ListIndex)
        
        If Not ConsistenciaStatus Then Exit Sub
          
        If EntSaida = "E" Then
            Xtipomov = "E"
            Xdebcred = "C"
        ElseIf EntSaida = "S" Then
            Xtipomov = "S"
            Xdebcred = "D"
        End If
        
        If TxtSerie.Text = "" Then
            XSerie = "U"
        Else
            XSerie = TxtSerie.Text
        End If
        If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente
            XCodDupl = "U"
        End If
        
'-------------------------------------------------------------------
        'Definição de XTipoFatura
        If XTipoFatura <> "3" Then
            If IsNull(DtpVencParcelaUnica.Value) Then
                If EntrouDupl Then
                    If Vetor2.UpperBound(1) > 0 Then
                        XTipoFatura = 2    '2 - Uma nota para várias duplicatas
                    Else
                        If Vetor2(0, 0) = Empty Then
                            XTipoFatura = 4        '4 - Falta compor a fatura
                        Else
                            XTipoFatura = 1    '1 - Uma nota para uma duplicata
                        End If
                    End If
                Else
                    If XInserir = "I" Then
                       XTipoFatura = 4
                    ElseIf Vetor2(0, 0) = Empty Then 'Só entra se for Alteração e o vetor de Duplic.estiver vazio
                        XTipoFatura = 4
                    End If
                End If
            Else
                XTipoFatura = 1
            End If
        End If
        'O XTipoFatura = 3 será atribuído pela rotina de composição de fatura
        'que significa:  Várias notas para várias duplicatas
'-------------------------------------------------------------------
        'Definir o código da duplicata como "U" para caso de unica e
        '"A" para caso de ser a primeira de uma série
        'If dtpvencparcelaunica.value = "  /  /  " Then (Não funciona se não entrar na orelha)
        If EntrouDupl Then
            If Vetor2.UpperBound(1) > 0 Then
                If Vetor2(0, 0) <> Empty Then
                    Vetor2(0, 0) = "A"
                End If
            Else
                If Vetor2(0, 0) <> Empty Then
                    Vetor2(0, 0) = "U"
                End If
            End If
        End If

        Conexao.BeginTrans
        If XInserir = "I" Then
            'Abre o banco de Dados
            SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
            ResFormulario.AddNew
        Else    'Alteração
            SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal = " & Chave, Dinamico
            xNumFat = ResFormulario!nofi_nr_fatura
        End If
        
        If ParcelasMultiplas = True Then
            DataVencimento = CDate(DtpVencParcelaUnica.Value)
            For x = 1 To Val(TxtNParcelas.Text)    'Grava Notas das parcelas múltiplas
                'Gera novo Nº de documento
                If XTipoFatura <> 4 Then
                    xNumFat = FunPegaGuardaUltimo("F")
                Else
                    xNumFat = 0
                End If
                If x > 1 Then
                    ResFormulario.AddNew
                End If
                If MskNumdocumento.Text = "" Or x > 1 Then
                    XNumDoc = Format(FunPegaGuardaUltimo("D"), "A000000000")
                Else
                    XNumDoc = MskNumdocumento.Text
                End If
                
                ResFormulario!nofi_tx_tipomov = Xtipomov
                ResFormulario!nofi_tx_status = XStatus
                ResFormulario!nofi_nr_documento = XNumDoc
                ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
                ResFormulario!nofi_tx_serie = XSerie
                ResFormulario!tido_cd_tipodoc = CboTipoDoc.BoundText
                ResFormulario!nofi_dt_entrada = DtpEntrada.Value
                ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
                ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
                ResFormulario!nofi_tx_distribuir = XDistribuido
                ResFormulario!nofi_nr_fatura = xNumFat
                ResFormulario!nofi_tx_tipofatura = XTipoFatura
                ResFormulario!empr_cd_empresa = PCodEmpresa
                
                ResFormulario.Update
    
                Chave = ResFormulario!nofi_cd_notafiscal
                    
                'Grava distribuição p/ parcelas múltiplas
                'Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal)" & _
                " VALUES ('" & DatPcr.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat(TxtValor.Text)) & "," & Chave & ")"
                Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal)" & _
                " VALUES ('" & DatPcr.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ")"
                Conexao.Execute Sql
                
                'Grava Duplicata da Nota de Parcelas Múltiplas
                'sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                      " VALUES ('U'," & funNuloData(DataVencimento, NomeSgbd) & "," & funNuloVal(funTrataFloat(LblValorLiquido1.Caption)) & ",'" & VerificaStatus(CboEntMultiplas.ListIndex) & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & Nulo(TxtDocPagto1.Text) & "'," & cboCCorrente.BoundText & "," & NuloValor(CboBancoForn1) & ",'" & xdebcred & "'," & xNumFat & "," & funNuloData(dtppagtoparcelaunica.value, NomeSgbd) & ")"
                Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                      " VALUES ('U'," & FunNuloData(DataVencimento, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(TxtDocPagto1.Text) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & xNumFat & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                Conexao.Execute Sql
                If OptMes.Value Then
                    DataVencimento = DateAdd("m", CInt(TxtPeriodicidade.Text), DataVencimento)
                ElseIf OptDia.Value Then
                    DataVencimento = DateAdd("d", CInt(TxtPeriodicidade.Text), DataVencimento)
                End If
                
            Next
        Else
            If XInserir = "I" Then
                If XTipoFatura <> 4 Then
                    xNumFat = FunPegaGuardaUltimo("F")
                Else
                    xNumFat = 0
                End If
            ElseIf XInserir = "A" Then 'Não é preciso incrementar o nº da fatura
                xNumFat = ResFormulario!nofi_nr_fatura
                If Tipo <> "1" Then
                    If xNumFat = 0 And Vetor2(0, 0) <> Empty Then
                        'A nota alterada não foi cadastrada com composição de duplicatas, logo não possuia NumFat
                        xNumFat = FunPegaGuardaUltimo("F")
                    End If
                End If
            End If
            'Gera Novo nº documento
            If MskNumdocumento.Text = "" Then
                XNumDoc = Format(FunPegaGuardaUltimo("D"), "A000000000")
            Else
                XNumDoc = MskNumdocumento.Text
            End If
            If XStatus = "P" Then
              ResFormulario!nofi_dt_emissao = Null
            Else
              ResFormulario!nofi_dt_emissao = Format$(DtpEmissao.Value, "DD/MM/YY")
            End If
            ResFormulario!nofi_tx_tipomov = Xtipomov
            ResFormulario!nofi_tx_status = XStatus
            ResFormulario!nofi_nr_documento = XNumDoc
            ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
            ResFormulario!nofi_tx_serie = XSerie
            ResFormulario!tido_cd_tipodoc = CboTipoDoc.BoundText
            ResFormulario!nofi_dt_entrada = DtpEntrada.Value
            ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
            ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
            ResFormulario!nofi_tx_distribuir = XDistribuido
            ResFormulario!nofi_nr_fatura = xNumFat
            ResFormulario!nofi_tx_tipofatura = XTipoFatura
            ResFormulario!empr_cd_empresa = PCodEmpresa
            
            ResFormulario.Update

            If XInserir = "I" Then
                Chave = ResFormulario!nofi_cd_notafiscal
            End If
        End If
        ResFormulario.Close
                  
        'Descontos
        If EntrouDesc And XStatus <> "P" Then
            If XInserir = "A" Then
                  Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & xNumFat
            End If
            If Vetor1(0, 0) <> Empty Then
                subGravaVetor "DescontosNF", "deen_dt_prevrecolhimento,deen_vl_valor,desc_cd_desconto, deen_nr_fatura", Array("D", "N", "N", "N"), 3, Array(1, 2, 3), "" & xNumFat & "", Vetor1, TDBGrid1
            End If
        End If
        
        'Duplicatas
        If Not ParcelasMultiplas Then
            If EntrouDupl Then
                If XInserir = "A" Then
                    If Vetor4(0, 0) <> Empty Then
                        GravaVetorDuplDescAcresc
                    Else
                        Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & xNumFat
                        'subGravaVetor "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 9, Array(0, 1, 2, 3, 5, 7, 8, 9, 11), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                        subGravaVetor "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 10, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                    End If
                End If
                If Vetor2(0, 0) <> Empty And XInserir = "I" Then
                    'subGravaVetor "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 9, Array(0, 1, 2, 3, 5, 7, 8, 9, 11), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                    subGravaVetor "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 10, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                End If
            Else
                If XInserir = "I" Then
                   If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente ou tem status = previsto
                        'Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                              " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(TxtDocPagto1.Text) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                        Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                              " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(TxtDocPagto1.Text) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                        Conexao.Execute Sql
                   End If
                ElseIf XInserir = "A" And (Tipo = "1" Or XStatus = "P") Then
'                    Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & xNumFat
'                    sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
'                          " VALUES ('" & XCodDupl & "'," & funNuloData(MskVencParcelaUnica, NomeSgbd) & "," & funNuloVal(funTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & Nulo(TxtDocPagto1.Text) & "'," & NuloValor(CboCCorrente) & "," & NuloValor(CboBancoForn1) & ",'" & xdebcred & "'," & Str(xNumFat) & "," & funNuloData(dtppagtoparcelaunica.value, NomeSgbd) & ")"
'                    Conexao.Execute sql
                    'A rotina foi modificada para atender às alterações de Dupl que possuam Desc/Acresc
                    Conexao.Execute "UPDATE Duplicatas " & _
                                        "SET dupl_nr_duplicata ='" & XCodDupl & "', " & _
                                        "dupl_dt_vencimento = " & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & _
                                        "dupl_vl_GlosaDevolucao = " & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & _
                                        "dupl_vl_valor = " & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & "," & _
                                        "dupl_tx_status = '" & XStatus & "'," & _
                                        "dupl_tx_tipopagrec = '" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "'," & _
                                        "dupl_tx_docpagrec = '" & FunNulo(TxtDocPagto1.Text) & "'," & _
                                        "coco_cd_codigo= " & FunNuloValor(CboCCorrente) & "," & _
                                        "bafo_cd_codigo = " & FunNuloValor(CboBancoForn1) & "," & _
                                        "dupl_tx_debcred ='" & Xdebcred & "'," & _
                                        "dupl_nr_fatura = " & Str(xNumFat) & "," & _
                                        "dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & " " & _
                                        "WHERE dupl_nr_fatura = " & xNumFat
                End If
            End If
        End If
        
        'Distribuição
        If EntrouDist Then
            If XInserir = "A" Then
                  Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
            End If
            If Vetor3(0, 0) <> Empty Then
                subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N"), 4, Array(0, 2, 5, 6), "" & Chave & "", Vetor3, TDBGrid3
            Else
                'se entrou na orelha de distribuicao excluiu o rateio e cadastrou na orelha de apropriação
                Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal)" & _
                    " VALUES ('" & DatPcr.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ")"
                Conexao.Execute Sql
            End If
        Else
            If ParcelasMultiplas = False Then
                If XInserir = "A" Then
                    Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
                End If
                If ChkDistribuicao Then
                    If Vetor3(0, 0) <> Empty Then
                        'subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_valor,nofi_cd_notafiscal", Array("T", "T", "N", "N"), 3, Array(0, 2, 4), "" & Chave & "", Vetor3, TDBGrid3
                        subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N"), 4, Array(0, 2, 5, 6), "" & Chave & "", Vetor3, TDBGrid3
                    End If
                Else
                    'Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,nofi_cd_notafiscal)" & _
                          " VALUES ('" & DatPcr.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat(TxtValor.Text)) & "," & Chave & ")"
                    Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal)" & _
                          " VALUES ('" & DatPcr.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ")"
                    Conexao.Execute Sql
                End If
            End If
        End If
        Conexao.CommitTrans
        
        'CmdGravar.Caption = "&Novo"
        'Verificar se é preciso atualizar composição de fatura
        If XValorBruto <> CDbl(TxtValor.Text) And XTipoFatura = "3" And XInserir = "A" Then
            If MsgBox("É necessário atualizar a Composição da Fatura. Deseja atualizar agora ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                'Chamar tela de composição de fatura
                NFatura = xNumFat
                XFormulario = "FrmNotasFiscais"
                FrmFatura.Show 1
                Set Formulario = FrmNotasFiscais
            End If
        End If
        If XInserir = "A" Then
            If CmdNovo.Enabled = False Then Call CmdDesistir_Click
        Else
            XInserir = "A"
            CmdNovo.Enabled = True
            
            If Not EntrouDesc Then
                Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
                subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura"), Vetor1, TDBGrid1
            End If
            Sql = "SELECT DuplDescAcresc.* from DuplDescAcresc LEFT JOIN Duplicatas ON Duplicatas.dupl_cd_duplicata = DuplDescAcresc.dupl_cd_duplicata Where Duplicatas.dupl_nr_fatura = " & NFatura
            CarregaVetorDescAcresc Sql, Array("dupl_cd_duplicata", "deac_cd_descacresc", "dude_vl_descacresc"), Vetor4, 3
        End If
        
        'Habilitar o processo de pagamento
        'If XTipoFatura <> "4" And XStatus = "C" Then CmdProcPagto.Enabled = True
        
'    ElseIf CmdGravar.Caption = "&Novo" Then
'        Chave = -1
'        XInserir = "I"
'        XDistribuido = "N"
'        NumDocAnt = ""
'
'        TabNotasFiscais.Tab = 0
'        TabAuxiliar.Tab = 0
'
'        PrepararInsercao
'        CboStatus.ListIndex = -1
'        If EntrouDesc Then
'           subLimpaVetor Vetor1, TDBGrid1
'        End If
'        If EntrouDupl Then
'           subLimpaVetor Vetor2, TDBGrid2
'        End If
'        If EntrouDist Then
'           subLimpaVetor Vetor3, TDBGrid3
'        End If
'        CboTipoDoc.Enabled = False
'        ParcelasMultiplas = False
'        CboFornecedor.Enabled = False
'        MskNumdocumento.Enabled = False
'        'CmdProcPagto.Enabled = False
'        CmdEntMultiplas.Enabled = True
'        DtpEntrada.Enabled = False
'        DtpEmissao.Enabled = False
'        TxtValor.Enabled = False
'        CboHistorico.Enabled = False
'        CmdLimparHistorico.Enabled = False
'        LblValorLiquido1.Enabled = False
'        ChkDistribuicao.Value = 0
'
'        TabAuxiliar.Enabled = False
'        CmdGravar.Caption = "&Gravar"
'
'        ValorTotalDuplicata = 0
'        ValorTotalDistribuicao = 0
'
'        TabNotasFiscais.TabEnabled(1) = False
'        TabNotasFiscais.TabEnabled(2) = False
'        TabNotasFiscais.TabEnabled(3) = False
'
'        EntrouDesc = False
'        EntrouDupl = False
'        EntrouDist = False

'    End If
    Exit Sub
TrataErro:
    funTrataErros (ComMensagem)
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
Sub GravaVetorDuplDescAcresc()
    Dim ResFormulario As Object
    Dim x As Integer, y As Integer
    Dim AchouDupl As Boolean
         
    For x = 0 To Vetor2.UpperBound(1)
         If Vetor2(x, 15) <> Empty Then
             SubQOpenRecordset ResFormulario, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = " & Vetor2(x, 15), Dinamico
             xNumFat = ResFormulario!dupl_nr_fatura
         Else
             SubQOpenRecordset ResFormulario, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
             ResFormulario.AddNew
         End If
    
         ResFormulario!dupl_nr_duplicata = Vetor2(x, 0)
         ResFormulario!dupl_dt_vencimento = Vetor2(x, 1)
         'ResFormulario!dupl_vl_valor = Vetor2(x, 2)
         ResFormulario!dupl_vl_GlosaDevolucao = Vetor2(x, 3)
         ResFormulario!dupl_vl_valor = Vetor2(x, 4)
         ResFormulario!dupl_tx_status = Vetor2(x, 5)
         ResFormulario!dupl_tx_TipoPagRec = Vetor2(x, 7)
         ResFormulario!dupl_tx_DocPagRec = Vetor2(x, 9)
         'If Vetor2(x, 8) = "  /  /  " Then
         If Not IsDate(Vetor2(x, 10)) Then
             ResFormulario!dupl_dt_pagrec = Null
         Else
             ResFormulario!dupl_dt_pagrec = Format$(Vetor2(x, 10), "DD/MM/YY")
         End If
         ResFormulario!coco_cd_codigo = Vetor2(x, 11)
         ResFormulario!bafo_cd_codigo = Vetor2(x, 13)
         ResFormulario!dupl_nr_fatura = xNumFat
         ResFormulario!dupl_tx_debcred = Xdebcred
                     
         ResFormulario.Update
    Next
    ResFormulario.Close
    
    'Verificar se exite algum registro no vetor de Descontos/Acrescimos
    'sem um cod. correspondente no vetor de duplicata
    AchouDupl = False
    For y = 0 To Vetor4.UpperBound(1)
        x = 0
        While x <= Vetor2.UpperBound(1) And Not AchouDupl
            If Vetor4(y, 0) = Vetor2(x, 15) Then
                AchouDupl = True
            End If
            x = x + 1
        Wend
        If Not AchouDupl Then
            Conexao.Execute "DELETE from DuplDescAcresc WHERE dupl_cd_duplicata = " & Vetor4(y, 0)
            'Apaga a nota excluida na orelha de Comp.Dupl. da tabela de Duplicatas,
            'já que o vetor de dupl. não será apagado pois existe descontos/acrescimos relacionados a ele
            Conexao.Execute "DELETE from Duplicatas WHERE dupl_cd_duplicata = " & Vetor4(y, 0)
        End If
        
        AchouDupl = False
    Next
End Sub


Function ConsistenciaDatas() As Boolean
'Verifica se as datas seguem a ordem de preenchimento
'Se ConsistenciaDatas= TRUE    >> Datas consistentes
'Se ConsistenciaDatas= FALSE  >> Datas não consistentes
    ConsistenciaDatas = False
    If CboStatus.ListIndex <> 0 Then
        If TabNotasFiscais.Tab = 0 Then
            'SE FOR BASE COMPETÊNCIA COMENTAR O IF ABAIXO
'            If DateDiff("d", MskEmissao.Text, MskEntrada.Text) < 0 Then
'                MsgBox "A data de emissão é posterior a data de entrada !", vbCritical, "ATENÇÃO"
'                MskEmissao.SetFocus
'                Exit Function
'            End If
            If Not IsNull(DtpVencParcelaUnica.Value) And Not IsNull(DtpEmissao.Value) Then
                If DateDiff("d", DtpEmissao.Value, DtpVencParcelaUnica.Value) < 0 Then
                    MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    DtpVencParcelaUnica.SetFocus
                    Exit Function
                End If
            End If
            If Not IsNull(DtpPagtoParcelaUnica.Value) Then
            If CboStatus.ListIndex = 2 Then
                If DateDiff("d", DtpEmissao.Value, DtpPagtoParcelaUnica.Value) < 0 Then
                    MsgBox "A data de pagamento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    DtpPagtoParcelaUnica.SetFocus
                    Exit Function
                End If
            End If
            End If
        ElseIf TabNotasFiscais.Tab = 2 Then
           If DateDiff("d", DtpEmissao.Value, DtpVencDesconto.Value) < 0 Then
                MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                DtpVencDesconto.SetFocus
                Exit Function
           End If
        ElseIf TabNotasFiscais.Tab = 1 Then
           If Not IsNull(DtpVencimento.Value) Then
                If DateDiff("d", DtpEmissao.Value, DtpVencimento.Value) < 0 Then
                    MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    DtpVencimento.SetFocus
                    Exit Function
                End If
           End If
           If CboStatus.ListIndex = 2 Then
           If Not IsNull(DtpPagamento.Value) Then
                If DateDiff("d", DtpEmissao.Value, DtpPagamento.Value) < 0 Then
                    MsgBox "A data de pagamento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    DtpPagamento.SetFocus
                    Exit Function
                End If
            End If
           End If
        End If
    ElseIf CboStatus.ListIndex = 0 Then
        If Not IsNull(DtpVencParcelaUnica.Value) Then
           If DateDiff("d", DtpEntrada.Value, DtpVencParcelaUnica.Value) < 0 Then
              MsgBox "A data de vencimento é anterior a data da entrada!", vbCritical, "ATENÇÃO"
              DtpVencParcelaUnica.SetFocus
              Exit Function
           End If
        End If
    End If
    ConsistenciaDatas = True
End Function

Private Sub CmdInsereDesconto_Click()
    Dim Xvalor As Double
                    
    FrmNotasFiscais.MousePointer = vbHourglass
    'If ExisteChave(1, Array("4"), Array(Fatura), Vetor1, TDBGrid1) Then Exit Sub
    ChaveD = -1
    XDesconto = "I"
    PrepararInsercao
    pandescontos.Visible = True
    pandescontos.Enabled = True
    
    CboDescontos.SetFocus
    DesabilitaBotoes
    FrmNotasFiscais.MousePointer = vbDefault
End Sub


Private Sub CmdInsereDistribuicao_Click()
    Dim Xvalor As Double
    
    FrmNotasFiscais.MousePointer = vbHourglass
    XDistribuicao = "I"
    PrepararInsercao
    If LblValorNota.Caption = LblValorTotalDistrib Then
       MsgBox "As distribuições já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
    Else
        PanDistribuicao.Visible = True
        PanDistribuicao.Enabled = True
        MskPcr2.SetFocus
        DesabilitaBotoes
    End If
    
    FrmNotasFiscais.MousePointer = vbDefault
End Sub

Private Sub CmdInsereduplicata_Click()
    Dim Xvalor As Double
    
    FrmNotasFiscais.MousePointer = vbHourglass
    XDuplicata = "I"
    ChaveDp = -1
    If (LblValorLiquido1.Caption + CDbl(TxtDevolucao.Text)) = LblValorTotal.Caption Then
       MsgBox "As duplicatas já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
    Else
        PnlCompDuplicata.Visible = True
        PnlCompDuplicata.Enabled = True
        TxtDocPagto2.Enabled = False
        DtpPagamento.Enabled = False
        
        PrepararInsercao

        'DesabilitaBotoes
        
        TxtValorDuplicata.SetFocus
    End If
    If CboStatus.Text = "Realizado" Then
        CboStatusduplicata.ListIndex = 1
        CboStatusduplicata.Enabled = False
    Else
        CboStatusduplicata.Enabled = True
    End If

    FrmNotasFiscais.MousePointer = vbDefault
End Sub


Private Sub CmdLimparHistorico_Click()
    CboHistorico.BoundText = 0
    CboHistorico.Text = ""
End Sub


Private Sub CmdNovo_Click()
    Chave = -1
    XInserir = "I"
    XDistribuido = "N"
    NumDocAnt = ""
    
    TabNotasFiscais.Tab = 0
    TabAuxiliar.Tab = 0

    PrepararInsercao
    CboStatus.ListIndex = -1
    If EntrouDesc Then
       subLimpaVetor Vetor1, TDBGrid1
    End If
    If EntrouDupl Then
       subLimpaVetor Vetor2, TDBGrid2
    End If
    If EntrouDist Then
       subLimpaVetor Vetor3, TDBGrid3
    End If
    CboTipoDoc.Enabled = False
    ParcelasMultiplas = False
    XFT_INSPARCMULTIPLA = "I"
    CboFornecedor.Enabled = False
    MskNumdocumento.Enabled = False
    'CmdProcPagto.Enabled = False
    CmdEntMultiplas.Enabled = True
    DtpEntrada.Enabled = False
    DtpEmissao.Enabled = False
    TxtValor.Enabled = False
    CboHistorico.Enabled = False
    CmdLimparHistorico.Enabled = False
    LblValorLiquido1.Enabled = False
    ChkDistribuicao.Value = 0

    TabAuxiliar.Enabled = False
    'CmdGravar.Caption = "&Gravar"
    
    ValorTotalDuplicata = 0
    ValorTotalDistribuicao = 0
    
    TabNotasFiscais.TabEnabled(1) = False
    TabNotasFiscais.TabEnabled(2) = False
    TabNotasFiscais.TabEnabled(3) = False

    EntrouDesc = False
    EntrouDupl = False
    EntrouDist = False

    CmdNovo.Enabled = False
End Sub



'Private Sub CmdProcPagto_Click()
'
'    Dim FiltroCCusto As String
'    Dim ResNF As Object
'
'    FiltroCCusto = ""
'
'    MdiPrincipal.Report1.WindowTitle = "Processo de Pagamento"
'    MdiPrincipal.Report1.ReportFileName = App.Path + "\ProcPagto.rpt"
'
'    'Pegar Centro de Custo. Para registros de comp. de fatura será apresentado distribuido também.
'    If XDistribuido = "N" Then
'       FiltroCCusto = MskPcr.Text + " - " + CboPcr.Text
'    Else
'       FiltroCCusto = "Distribuído"
'    End If
'
'    MdiPrincipal.Report1.XGT_SELECAO = "({ConsFINFluxoCaixa.dupl_nr_fatura} = " & xNumFat & ")"
'    MdiPrincipal.Report1.XGT_SELECAO = MdiPrincipal.Report1.XGT_SELECAO + " AND uppercase({ConsFINFluxoCaixa.dupl_tx_status}) LIKE 'C' "
'
''********** Ativa os Relatórios
'    MdiPrincipal.Report1.Formulas(0) = "CCusto = """ + FiltroCCusto + """"
'    MdiPrincipal.Report1.Action = 0
'    TelaNotasFiscais.MousePointer = vbDefault
'    MdiPrincipal.Report1.XGT_SELECAO = ""
'    MdiPrincipal.Report1.Formulas(0) = ""
'
'    'Atualiza a tabela de Duplicatas
'    subQOpenRecordset ResNF, "SELECT dupl_dt_ProcPagamento FROM Duplicatas where dupl_nr_fatura =" & xxNumFat, "Estatico"
'    ResNF.MoveFirst
'    While Not ResNF.EOF
'        If Not IsDate(ResNF!dupl_dt_ProcPagamento = funNuloData(Now, NomeSgbd)) Then
'            ResNF!dupl_dt_ProcPagamento = Format(Now, "DD/MM/YY")
'            ResNF.MoveNext
'        End If
'    Wend
'    ResNF.Close
'
''*************** Imprimir relação de notas da fatura *******************
'    If XTipoFatura = "3" Then
'        If MsgBox("Deseja imprimir a relação de Notas Fiscais desta Fatura?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
'            MdiPrincipal.Report1.WindowTitle = "Relação de Notas da Fatura Nº " + CStr(xNumFat)
'            MdiPrincipal.Report1.ReportFileName = App.Path + "\RelNotasProcPagto.rpt"
'
'            MdiPrincipal.Report1.XGT_SELECAO = "({ConsFINNotasFiscais.nofi_nr_fatura} = " & xNumFat & ")"
'
'            '********** Ativa os Relatórios
'            MdiPrincipal.Report1.Action = 0
'            TelaNotasFiscais.MousePointer = vbDefault
'            MdiPrincipal.Report1.XGT_SELECAO = ""
'        Else
'            Exit Sub
'        End If
'    End If
'
'End Sub


Private Sub CmdRemoverDistribuicao_Click()
    Dim Houve_Erro As Integer

    Houve_Erro = False
    
    DesabilitaBotoes
   
    TDBGrid3.Enabled = False
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor3, TDBGrid3
    End If

    If Vetor3(0, 0) <> Empty Then
        HabilitaBotoes
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    Else
        LblValorTotalDistrib.Caption = ""
        ValorTotalDistribuicao = 0
        CmdInsereDistribuicao.Enabled = True
    End If
   
    TDBGrid3.Enabled = True
 
    Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next

End Sub

Private Sub CmdSairDesconto_Click()
    pandescontos.Visible = False
    pandescontos.Enabled = False
       
    XDesconto = ""
    TDBGrid1.Enabled = True
    
    If Vetor1(0, 0) <> Empty Then
       HabilitaBotoes
    Else
       DesabilitaBotoes
       CmdInsereDesconto.Enabled = True
    End If
End Sub

Private Sub CmdSairDistribuicao_Click()
    PanDistribuicao.Visible = False
    PanDistribuicao.Enabled = False
       
    XDistribuicao = ""
    TDBGrid3.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
       HabilitaBotoes
    Else
       CmdInsereDistribuicao.Enabled = True
    End If
End Sub

Private Sub CmdSairDuplicata_Click()
    
    PnlCompDuplicata.Visible = False
    PnlCompDuplicata.Enabled = False
    
    XDuplicata = ""
    TDBGrid2.Enabled = True
    
    TDBGrid2.MoveFirst
    If (TDBGrid2.ApproxCount) <> 0 Then
        CmdInsereduplicata.Enabled = True
        CmdAlteraduplicata.Enabled = True
        CmdExcluiduplicata.Enabled = True
        TDBGrid2.Bookmark = PosCursor
    Else
        CmdInsereduplicata.Enabled = True
    End If
        
End Sub

Private Sub CmdSairEntMult_Click()
    PanEntMultiplas.Visible = False
    PanEntMultiplas.Enabled = False
    TabNotasFiscais.Enabled = True
    CmdGravar.Enabled = True
    CmdDesistir.Enabled = True
End Sub

Private Sub Command1_Click()
    PanFornecedor.Visible = False
    PanFornecedor.Enabled = False
End Sub

Private Sub Command2_Click()
    Dim ChaveTemp As Variant
    
    ChaveTemp = Chave
        
    If EntSaida = "E" Then
        FornClie = "C" 'Cliente
    Else
        FornClie = "F"  'Fornecedor
    End If
    
    If OptJuridica.Value = True Then
        Pessoa = "J" 'Pessoa Jurídica
    ElseIf OptFisica.Value = True Then
        Pessoa = "F" 'Pessoa Física
    ElseIf OptProduto.Value = True Then
        Pessoa = "3" 'Produto
    ElseIf OptProdServ.Value = True Then
        Pessoa = "4" 'Produto e Serviço
    End If
    'XFormForn = "FrmNotasFiscais"
    FrmFornecedores.Show 1
    Chave = ChaveTemp
    Set Formulario = FrmNotasFiscais
    FrmNotasFiscais.DatFornecedor.Refresh
    FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
    
    PanFornecedor.Visible = False
    PanFornecedor.Enabled = False
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmNotasFiscais
    
End Sub

Private Sub Form_Load()
    Dim ResFat As Object
    'On Error GoTo TrataErro
    
    'PrimeiraVez = True
    subManutencaoJanelasAtivas "I", "FrmNotasFiscais"
    
    Set Formulario = FrmNotasFiscais
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If

   ' CliqueBarraFerramentas = False
 
    If XInserir = "I" Then
        subConectarControleDados DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') ORDER BY tido_tx_descricao", Estatico
    Else
        subConectarControleDados DatTipoDoc, "SELECT * FROM Tipodocumentos ORDER BY tido_tx_descricao", Estatico
    End If
        
    'Comum às orelhas de dados gerais e duplicatas
    subConectarControleDados DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
    'Comum às orelhas de dados gerais e distribuição
    subConectarControleDados DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    
    If EntSaida = "E" Then
        LblNomeFantasia.Caption = "Cliente:"
        'LblValorNota2.Left = 5950
        LblValorNota2.Caption = "Valor a Receber"
        'subConectarControleDados DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_fantasia ", Estatico
        'subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'C' and (empr_cd_empresa = " & PCodEmpresa & " or isnull(empr_cd_empresa)) ORDER BY focl_tx_fantasia ", Estatico
        subConectarControleDados DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
        subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'C' and empr_cd_empresa = " & PCodEmpresa & " ORDER BY focl_tx_razaosocial ", Estatico
        'subConectarControleDados DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
        subConectarControleDados DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        LblTipoPagto.Caption = "Tipo Recebimento:"
        LblDocPagto.Caption = "Doc. Recebimento:"
        LblBcForn.Visible = False
        CboBancoForn1.Visible = False
        LblPagto.Caption = "Recebimento:"
        PanPagto1.Top = 570
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        
        CboTipoPagto2.RemoveItem (1)
        CboTipoPagto2.RemoveItem (1)
        CboTipoPagto2.RemoveItem (1)
        CboTipoPagto2.RemoveItem (1)
        
        TDBGrid2.Columns("Tipo Pag.").Caption = "Tipo Rec."
        TDBGrid2.Columns("Doc. Pag.").Caption = "Doc. Rec."
        TDBGrid2.Columns("Data Pag.").Caption = "Data Rec."
        TDBGrid2.Columns("Conta Corrente (Fornecedor)").Visible = False
        TDBGrid1.Columns("Vencimento").Visible = False
        LblVlDesc.Left = LblVencimento3.Left
        TxtValorDesconto.Left = DtpVencDesconto.Left
        TabAuxiliar.TabCaption(1) = "Recebimento"
        LblVencimento3.Visible = False
        DtpVencDesconto.Visible = False
        'CmdProcPagto.Enabled = False
        'CmdGravar.Left = 2685
        'CmdDesistir.Left = 5370
    ElseIf EntSaida = "S" Then
        LblNomeFantasia.Caption = "Fornecedor:"
        LblValorNota2.Caption = "Valor a Pagar"
        subConectarControleDados DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
        'subConectarControleDados DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
        subConectarControleDados DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        'subConectarControleDados DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'F' ORDER BY focl_tx_fantasia ", Estatico
        'subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and (empr_cd_empresa = " & PCodEmpresa & " or isnull(empr_cd_empresa)) ORDER BY focl_tx_fantasia ", Estatico
        subConectarControleDados DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and empr_cd_empresa = " & PCodEmpresa & " ORDER BY focl_tx_razaosocial ", Estatico
        
        CboTipoPagto1.RemoveItem (5)
        CboTipoPagto2.RemoveItem (5)
    End If
     
    EntrouDesc = False
    EntrouDupl = False
    EntrouDist = False
    ParcelasMultiplas = False
    XFT_INSPARCMULTIPLA = "I"
    XTipoFatura = "Null"
    MskPcr.Mask = pPCRmascara
    MskPcr2.Mask = pPCRmascara
    NumDocAnt = ""

    XDistribuido = "N"
    
    If XInserir = "I" Then
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
        TabNotasFiscais.TabEnabled(3) = False
        LblValorTotal.Caption = ""
        LblValorTotalDistrib.Caption = ""
        PrepararInsercao
    ElseIf XInserir = "A" Then
'        subQOpenRecordset ResFat, "SELECT NotasFiscais.nofi_tx_tipofatura FROM NotasFiscais where NotasFiscais.nofi_nr_fatura =" & NFatura, "Estatico"
'        If Not IsNull(ResFat!nofi_tx_tipofatura) Then
'            XTipoFatura = ResFat!nofi_tx_tipofatura
'        End If
'        ResFat.Close
        PreencheCampos (XFormulario)
        If XFormulario = "TelaDuplicatas" Then 'Foi chamado da tela de Duplicata
            If Status <> "0" And Tipo = "0" Then 'Não é previsto e o tipo Doc. não gera dupl. automaticamente
                FrmNotasFiscais.TabNotasFiscais.Tab = 1
            End If
        End If
        If Status <> "2" Then
            HabilitaControles
        Else
            DesabilitaRealizado "Nota", False
        End If
    End If
Exit Sub

TrataErro:
    If Err.Number = NumberTravaRegistro And Err.Description = DescriptionTravaRegistro Then
        If MsgBox("O exemplo está sendo alterado por outro usuário. Deseja tentar novamente?", vbCritical + vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
            Resume
        Else
            Unload Me
            Exit Sub
        End If
    ElseIf Err.Number = NumberRegistroExcluido And Err.Description = DescriptionRegistroExcluido Then
        MsgBox "O exemplo foi excluído por outro usuário!", vbCritical, "ATENÇÃO"
        Unload Me
        Exit Sub
    Else
        funTrataErros (ComMensagem)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmNotasFiscais"
End Sub

Private Sub MskConta_GotFocus()
    Call subSelecionaMSK(Mskconta)
End Sub
Private Sub MskConta_LostFocus()
    If Mskconta.Text <> "" Then
        CboPlanoContas.BoundText = Mskconta.Text
        If CboPlanoContas.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            Mskconta.SetFocus
        End If
    
    End If
End Sub

Private Sub MskConta2_GotFocus()
    Call subSelecionaMSK(MskConta2)
End Sub


Private Sub MskConta2_LostFocus()
    If MskConta2.Text <> "" Then
        CboPlanoContas2.BoundText = MskConta2.Text
        If CboPlanoContas2.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta2.SetFocus
        End If
    End If
End Sub


Private Sub dtpEmissao_LostFocus()
    If Not ConsistenciaDatas Then Exit Sub
End Sub

Private Sub MskNumdocumento_LostFocus()
    Dim ResFormulario As Object
    'Checar se o documento já foi cadastrado
    'If MskNumdocumento.Text <> "" And XInserir = "I" Then
    If MskNumdocumento.Text <> "" And MskNumdocumento.Text <> NumDocAnt Then
        SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE focl_cd_forncli = " & CboFornecedor.BoundText & " and nofi_nr_documento = '" & MskNumdocumento.Text & "'", Estatico
        If Not ResFormulario.EOF Then
           MsgBox "Documento já cadastrado.", vbCritical + vbOKOnly, "ATENÇÃO"
           MskNumdocumento.SetFocus
           ResFormulario.Close
           Exit Sub
        End If
        ResFormulario.Close
    End If
End Sub



Private Sub dtpPagtoParcelaUnica_LostFocus()
    If IsDate(DtpPagtoParcelaUnica.Value) Then
        If Not ConsistenciaDatas Then Exit Sub
    End If
End Sub

Private Sub MskPcr_GotFocus()
    Call subSelecionaMSK(MskPcr)
End Sub


Private Sub MskPcr_LostFocus()
    MskPcr.PromptInclude = False
    If MskPcr.Text <> "" Then
        MskPcr.Mask = pPCRmascara
        MskPcr.PromptInclude = True
        CboPcr.BoundText = MskPcr.Text
        If CboPcr.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr.Visible = True Then MskPcr.SetFocus
        End If
    End If
    MskPcr.PromptInclude = True
End Sub


Private Sub MskPcr2_GotFocus()
    Call subSelecionaMSK(MskPcr2)
End Sub


Private Sub MskPcr2_LostFocus()
    MskPcr2.PromptInclude = False
    If MskPcr2.Text <> "" Then
        MskPcr2.Mask = pPCRmascara
        MskPcr2.PromptInclude = True
        CboPcr2.BoundText = MskPcr2.Text
        If CboPcr2.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr2.Visible = True Then MskPcr2.SetFocus
        End If
    End If
    MskPcr2.PromptInclude = True
End Sub

Private Sub dtpVencDesconto_LostFocus()
    If Not ConsistenciaDatas Then Exit Sub
End Sub

Private Sub dtpVencimento_LostFocus()
    If Not ConsistenciaDatas Then Exit Sub
End Sub

Private Sub dtpVencParcelaUnica_LostFocus()
    If IsDate(DtpVencParcelaUnica.Value) Then
        If Not ConsistenciaDatas Then Exit Sub
    End If
End Sub

Private Sub TabAuxiliar_Click(PreviousTab As Integer)
    If XInserir = "I" Then
        If IsNull(DtpVencParcelaUnica.Value) Then DtpVencParcelaUnica.Value = Format$(Now, "dd/MM/yy")
        DtpVencParcelaUnica.SetFocus
        If CboStatus.Text = "Realizado" Then
            If IsNull(DtpPagtoParcelaUnica.Value) Then DtpPagtoParcelaUnica.Value = Format$(Now, "dd/mm/yy")
        End If
    End If
End Sub

Private Sub TabNotasFiscais_Click(PreviousTab As Integer)
    Dim Sql As String
    Dim ResForn As Object
    Dim XPrevRecolhimento As String
    Dim XVlDesconto As Double
    Dim XData As Date
    Dim Ano As Integer
    Dim Mes As Integer
    'Variáveis utilizada na alteração da duplicata a partir da Tela de Duplicatas
    Dim x As Integer
    
    If TabNotasFiscais.Tab <> 0 And XFormulario = "TelaNotasFiscais" Then
        If ChecarCamposOrelhaPrincipal Then
            TabNotasFiscais.Tab = 0
            Exit Sub
        End If
    End If
      
    '****Preenche o cabecalho das orelhas com o nome do Forn/cliente, nº doc/Serie****
    PnlDados1_1.Caption = (" DOC.: " & MskNumdocumento.Text & " / " & TxtSerie.Text)
    PnlDados2_2.Caption = (" DOC.: " & MskNumdocumento.Text & " / " & TxtSerie.Text)
    PnlDados3_3.Caption = (" DOC.: " & MskNumdocumento.Text & " / " & TxtSerie.Text)
    If EntSaida = "S" Then
        PnlDados1.Caption = (" FORNECEDOR: " & CboFornecedor.Text)
        PnlDados2.Caption = (" FORNECEDOR: " & CboFornecedor.Text)
        PnlDados3.Caption = (" FORNECEDOR: " & CboFornecedor.Text)
    Else
        PnlDados1.Caption = (" CLIENTE: " & CboFornecedor.Text)
        PnlDados2.Caption = (" CLIENTE: " & CboFornecedor.Text)
        PnlDados3.Caption = (" CLIENTE: " & CboFornecedor.Text)
    End If
    '*********************************************************
      
    If TabNotasFiscais.Tab = 1 And Not EntrouDupl Then
               
        If EntSaida = "E" Then 'Prepara o painel para Contas a Receber
            LblTipoPagamento.Caption = "Tipo Recebimento:"
            LblDocPagRec.Caption = "Doc. Recebimento:"
            LblBcFornecedor.Visible = False
            CboBancoForn2.Visible = False
            LblPagamento.Caption = "Recebimento:"
            PanPagto2.Top = 375
'            CboTipoPagto2.RemoveItem (1)
'            CboTipoPagto2.RemoveItem (1)
'            CboTipoPagto2.RemoveItem (1)
        End If
        Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
        'subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2 'ACCESS (testar no SQL Server)
        'subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
        subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_vl_GlosaDevolucao", "Valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
    
        EntrouDupl = True
        If Vetor2(0, 0) = Empty Then
            DesabilitaBotoes
            CmdInsereduplicata.Enabled = True
        Else
            If XFormulario = "TelaDuplicatas" Then
                x = 0
                While x <= Vetor2.UpperBound(1) 'Utilizado para a alteração da duplicata, posicionando o cursor no registro desejado
                    If Vetor2(x, 15) = CStr(ChaveDp) Then
                        TDBGrid2.Bookmark = x
                        PosCursor = x
                        x = Vetor2.UpperBound(1)
                    End If
                    x = x + 1
                Wend
                CmdAlteraduplicata_Click
                TDBGrid2.Bookmark = PosCursor
            End If
        End If
        LblValorTotal = funSomaColuna(Vetor2, 2)
        ValorTotalDuplicata = CDbl(LblValorTotal.Caption)

    ElseIf TabNotasFiscais.Tab = 2 Then
        If Not EntrouDesc Then
            If EntSaida = "S" Then  'É um fornecedor
                'subConectarControleDados DatDescontosNF, "SELECT * FROM DescontosImpostos WHERE desc_tx_classe = 'D' ORDER BY desc_tx_descricao", "Estatico"
                subConectarControleDados DatDescontosNF, "SELECT * FROM ConsFINPlCDescImp WHERE desc_tx_classe = 'D' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao", "Estatico"
            Else  'É um Cliente
                'subConectarControleDados DatDescontosNF, "SELECT * FROM DescontosImpostos WHERE desc_tx_classe = 'I' and desc_tx_fonte = 'S' ORDER BY desc_tx_descricao", "Estatico"
                subConectarControleDados DatDescontosNF, "SELECT * FROM ConsFINPlCDescImp WHERE desc_tx_classe = 'I' and desc_tx_fonte = 'S' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao", "Estatico"
            End If
            
            Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
            subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura"), Vetor1, TDBGrid1
            If Vetor1(0, 0) = Empty Then
                DesabilitaBotoes
                CmdInsereDesconto.Enabled = True
            End If
        
            'Verifica se o fornecedor possui um imposto cadastrado
            If EntSaida = "S" And XInserir = "I" Then 'É um fornecedor
            'If XInserir = "I" Then
                'subQOpenRecordset ResForn, "SELECT DesImpForn.*,DescontosImpostos.desc_cd_desconto,DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase FROM DesImpForn LEFT JOIN DescontosImpostos ON DesImpForn.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DesImpForn.focl_cd_forncli=" & CboFornecedor.BoundText, "Estatico"
                SubQOpenRecordset ResForn, "SELECT DesImpForn.*,DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase FROM DesImpForn LEFT JOIN DescontosImpostos ON DesImpForn.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DesImpForn.focl_cd_forncli=" & CboFornecedor.BoundText, "Estatico"
                If Not (ResForn.BOF And ResForn.EOF) Then
                    ResForn.MoveFirst
                    Do While Not ResForn.EOF
                        If ResForn!desc_tx_tipobase = "1" Or ResForn!desc_tx_tipobase = "3" Then
                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            'Ano = Year(CDate(MskVencimento.Text))
                            'Mes = Month(CDate(MskVencimento.Text))
                            If EntrouDupl Then
                                If Vetor2(0, 0) <> Empty Then
                                    Ano = Year(CDate(Vetor2(0, 1)))
                                    Mes = Month(CDate(Vetor2(0, 1)))
                                    
                                    If Mes = 12 Then
                                        Mes = 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                                Else
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                                End If
                            Else
                                subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                            End If
                        ElseIf ResForn!desc_tx_tipobase = "2" Then
                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            'subInsereRegistroVetor Array(ResForn!desc_tx_descricao, Format(DtpVencimento.Value, "dd/mm/yyyy"), XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                            If EntrouDupl Then
                                If Vetor2(0, 0) <> Empty Then
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, Format(DtpVencimento.Value, "dd/mm/yyyy"), XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                                Else
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                                End If
                            Else
                                subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                            End If
                        End If
                        ResForn.MoveNext
                        HabilitaBotoes
                    Loop
                    ResForn.Close
                End If
                LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
            ElseIf EntSaida = "E" And XInserir = "I" Then 'É um Cliente
                SubQOpenRecordset ResForn, "SELECT DescontosImpostos.desc_cd_desconto, DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase, DescontosImpostos.desc_tx_fonte, DescontosImpostos.desc_vl_taxa FROM DescontosImpostos where descontosimpostos.desc_tx_fonte = 'S'", "Estatico"
                If Not (ResForn.BOF And ResForn.EOF) Then
                    ResForn.MoveFirst
                    Do While Not ResForn.EOF
'                       If ResForn!desc_tx_tipobase = "1" Then
                          XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desc_vl_taxa)) / 100), "##,##0.00")
'                          Ano = Year(CDate(MskVencimento.Text))
'                          Mes = Month(CDate(MskVencimento.Text))
'                          If Mes = 12 Then
'                             Mes = 1
'                          Else
'                             Mes = Mes + 1
'                          End If
'                          XData = DateSerial(Ano, Mes, ResForn!desc_tx_diarecolhimento)
'                          InsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
'                       ElseIf ResForn!desc_tx_tipobase = "2" Or ResForn!desc_tx_tipobase = "3" Then
'                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desc_vl_taxa)) / 100), "##,##0.00")
'                            InsereRegistroVetor Array(ResForn!desc_tx_descricao, MskVencimento.Text, XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
'                       End If
                       subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "  /  /  ", XVlDesconto, ResForn!desc_cd_desconto, ""), Vetor1, TDBGrid1
                       ResForn.MoveNext
                       HabilitaBotoes
                    Loop
                    ResForn.Close
                End If
                LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
            End If
            EntrouDesc = True
        End If
        
    ElseIf TabNotasFiscais.Tab = 3 And Not EntrouDist Then
        'Sql = "SELECT Distribuicao.*,Notasfiscais.nofi_vl_valor,(Distribuicao.dist_vl_valor / NotasFiscais.nofi_vl_valor) AS Percentual,CentroResultados.cere_cd_estruturado, PlanoContas.plco_cd_codreduzido " & _
        "FROM NotasFiscais,Distribuicao,CentroResultados, PlanoContas WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " and Distribuicao.nofi_cd_notafiscal = " & Chave & " and Distribuicao.cere_cd_pcr =CentroResultados.cere_cd_pcr and Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta"
        'subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
        'Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " and nofi_cd_notafiscal = " & Chave & ""
        'subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "percentual"), Vetor3, TDBGrid3
        Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " and nofi_cd_notafiscal = " & Chave & ""
        subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "dist_vl_GlosaDevolucao", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
        
        EntrouDist = True
        If Vetor3(0, 0) = Empty Then
            DesabilitaBotoes
            CmdInsereDistribuicao.Enabled = True
        End If
        LblValorNota.Caption = Format$(TxtValor.Text, "##,##0.00")
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    End If
End Sub


Private Sub TxtDocPagto1_LostFocus()
    If CboTipoPagto1.ListIndex = 0 Then
        TxtDocPagto1.Text = Format(TxtDocPagto1, "0000000")
    End If
End Sub


Private Sub TxtDocPagto2_LostFocus()
    If CboTipoPagto2.ListIndex = 0 Then
        TxtDocPagto2.Text = Format(TxtDocPagto2, "0000000")
    End If
End Sub


Private Sub TxtPercDistribuicao_GotFocus()
    Call subSelecionaTXT(TxtPercDistribuicao)
End Sub


Private Sub TxtPercDistribuicao_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtPercDistribuicao_LostFocus()
    If TxtPercDistribuicao.Text <> "" Then
        TxtVldistribuicao.Text = Format$(((CDbl(TxtValor.Text) * CDbl(TxtPercDistribuicao.Text)) / 100), "##,##0.00")
    End If
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtValor_LostFocus()
    If IsNumeric(TxtValor.Text) Then
        TxtValor.Text = Format$(TxtValor.Text, "##,##0.00")
        LblValorLiquido1.Caption = Format$(TxtValor.Text - CDbl(TxtDevolucao.Text), "##,##0.00")
    Else
        TxtValor.Text = ""
    End If
End Sub

Private Sub TxtValorDesconto_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub


Private Sub TxtValorDuplicata_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtValorDuplicata_LostFocus()
    If IsNumeric(TxtValorDuplicata.Text) Then
        TxtValorDuplicata.Text = Format$(TxtValorDuplicata.Text, "##,##0.00")
    Else
        TxtValorDuplicata.Text = ""
    End If
End Sub


Private Sub TxtVldistribuicao_GotFocus()
    Call subSelecionaTXT(TxtVldistribuicao)

End Sub


Private Sub TxtVldistribuicao_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtVldistribuicao_LostFocus()
    If TxtVldistribuicao.Text <> "" Then
        'TxtPercDistribuicao.Text = Format(CStr(((CDbl(TxtVldistribuicao.Text) / CDbl(TxtValor.Text)) * 100)), "standard")
        TxtPercDistribuicao.Text = ((CDbl(TxtVldistribuicao.Text) / CDbl(TxtValor.Text)) * 100)
    End If

End Sub


