VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmGlosaDevolucao 
   Caption         =   "Cadastro de Nota de Devolução"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   1590
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   9375
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
      TabIndex        =   15
      Top             =   5000
      Width           =   1815
   End
   Begin TabDlg.SSTab TabNotasFiscais 
      Height          =   4785
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8440
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "FrmGlosaDevolucao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TabAuxiliar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSPanel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PanNotasFiscais"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Composição de Duplicata"
      TabPicture(1)   =   "FrmGlosaDevolucao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "LblValorTotal"
      Tab(1).Control(2)=   "PnlDados1_1"
      Tab(1).Control(3)=   "PnlDados1"
      Tab(1).Control(4)=   "TDBGrid2"
      Tab(1).Control(5)=   "CmdExcluiduplicata"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Distribuição"
      TabPicture(2)   =   "FrmGlosaDevolucao.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "LblValorTotalDistrib"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "LblValorNota"
      Tab(2).Control(4)=   "PnlDados3_3"
      Tab(2).Control(5)=   "PnlDados3"
      Tab(2).Control(6)=   "TDBGrid3"
      Tab(2).Control(7)=   "CmdRemoverDistribuicao"
      Tab(2).ControlCount=   8
      Begin VB.CommandButton CmdRemoverDistribuicao 
         Caption         =   "Excl&uir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   37
         Top             =   975
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiduplicata 
         Caption         =   "&Excluir"
         Height          =   330
         Left            =   -74805
         TabIndex        =   14
         Top             =   4335
         Width           =   1080
      End
      Begin Threed.SSPanel PanNotasFiscais 
         Height          =   1050
         Left            =   120
         TabIndex        =   19
         Top             =   1770
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1852
         _StockProps     =   15
         Caption         =   "Nota de Devolução"
         ForeColor       =   12582912
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Begin VB.TextBox TxtDevGlosa 
            Height          =   315
            Left            =   7335
            TabIndex        =   4
            Top             =   645
            Width           =   1665
         End
         Begin MSAdodcLib.Adodc DatHistorico2 
            Height          =   345
            Left            =   2460
            Top             =   660
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
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
         Begin MSMask.MaskEdBox MskDocGlosaDev 
            Height          =   315
            Left            =   7995
            TabIndex        =   2
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboHistGlosaDev 
            Bindings        =   "FrmGlosaDevolucao.frx":0054
            Height          =   315
            Left            =   1260
            TabIndex        =   3
            Top             =   660
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_codigo"
            Text            =   "CboHistorico"
         End
         Begin Threed.SSCommand CmdLimparHistorico 
            Height          =   330
            Left            =   5355
            TabIndex        =   59
            Top             =   645
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            Enabled         =   0   'False
            MouseIcon       =   "FrmGlosaDevolucao.frx":0070
            Picture         =   "FrmGlosaDevolucao.frx":04C2
         End
         Begin MSComCtl2.DTPicker DtpCompGlosaDev 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   285
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   83034113
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker DtpEmissaoGlosaDev 
            Height          =   315
            Left            =   4245
            TabIndex        =   1
            Top             =   285
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   83034113
            CurrentDate     =   37460
         End
         Begin VB.Label LblHistGlosaDev 
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
            Height          =   270
            Left            =   90
            TabIndex        =   63
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label LblDocGlosaDev 
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
            Height          =   240
            Left            =   6585
            TabIndex        =   62
            Top             =   315
            Width           =   1365
         End
         Begin VB.Label LblEmissaoGlosaDev 
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
            Height          =   255
            Left            =   3285
            TabIndex        =   61
            Top             =   330
            Width           =   885
         End
         Begin VB.Label LblCompGlosaDev 
            Alignment       =   1  'Right Justify
            Caption         =   "Entrada:"
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
            Left            =   30
            TabIndex        =   60
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label LblDevGlosa 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Devolução:"
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
            Left            =   5820
            TabIndex        =   44
            Top             =   690
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   390
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   2355
         _StockProps     =   15
         Caption         =   "Nota de Origem"
         ForeColor       =   8421376
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Begin VB.TextBox TxtSerie 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8505
            MaxLength       =   3
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox TxtValor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   45
            Top             =   945
            Width           =   1800
         End
         Begin MSAdodcLib.Adodc DatTipoDoc 
            Height          =   330
            Left            =   6360
            Top             =   240
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
            Bindings        =   "FrmGlosaDevolucao.frx":05D4
            Height          =   315
            Left            =   6075
            TabIndex        =   13
            Top             =   255
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "tido_tx_descricao"
            BoundColumn     =   "tido_cd_tipodoc"
            Text            =   "CboTipoDoc"
         End
         Begin MSAdodcLib.Adodc DatFornecedor 
            Height          =   330
            Left            =   2820
            Top             =   255
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
            Left            =   2460
            Top             =   945
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
            Bindings        =   "FrmGlosaDevolucao.frx":05ED
            Height          =   315
            Left            =   1215
            TabIndex        =   47
            Top             =   255
            Width           =   3915
            _ExtentX        =   6906
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
            Left            =   6660
            TabIndex        =   48
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboHistorico 
            Bindings        =   "FrmGlosaDevolucao.frx":0609
            Height          =   315
            Left            =   1215
            TabIndex        =   49
            Top             =   945
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_codigo"
            Text            =   "CboHistorico"
         End
         Begin MSComCtl2.DTPicker DtpEntrada 
            Height          =   315
            Left            =   1215
            TabIndex        =   50
            Top             =   600
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            DateIsNull      =   -1  'True
            Format          =   83034113
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker DtpEmissao 
            Height          =   315
            Left            =   3600
            TabIndex        =   51
            Top             =   615
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   83034113
            CurrentDate     =   37460
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
            Height          =   270
            Left            =   90
            TabIndex        =   58
            Top             =   285
            Width           =   1125
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
            Height          =   285
            Left            =   90
            TabIndex        =   57
            Top             =   975
            Width           =   1125
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
            Height          =   270
            Left            =   5265
            TabIndex        =   56
            Top             =   645
            Width           =   1365
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
            Height          =   270
            Left            =   6585
            TabIndex        =   55
            Top             =   1005
            Width           =   570
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
            Height          =   255
            Left            =   7845
            TabIndex        =   54
            Top             =   645
            Width           =   630
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
            Left            =   2700
            TabIndex        =   53
            Top             =   660
            Width           =   885
         End
         Begin VB.Label LblEntrada 
            Alignment       =   1  'Right Justify
            Caption         =   "Entrada:"
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
            Left            =   30
            TabIndex        =   52
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label LblTipoDoc 
            Caption         =   "Tipo Doc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   21
            Top             =   300
            Width           =   915
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   1815
         Left            =   105
         TabIndex        =   22
         Top             =   2865
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   3201
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Apropriação"
         TabPicture(0)   =   "FrmGlosaDevolucao.frx":0624
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
         TabPicture(1)   =   "FrmGlosaDevolucao.frx":0640
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "PanPagto1"
         Tab(1).Control(1)=   "CboBancoForn1"
         Tab(1).Control(2)=   "LblBcForn"
         Tab(1).ControlCount=   3
         Begin Threed.SSPanel PanPagto1 
            Height          =   1080
            Left            =   -74940
            TabIndex        =   32
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
               ItemData        =   "FrmGlosaDevolucao.frx":065C
               Left            =   1860
               List            =   "FrmGlosaDevolucao.frx":0675
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   375
               Width           =   2655
            End
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "FrmGlosaDevolucao.frx":06CF
               Height          =   330
               Left            =   1860
               TabIndex        =   11
               Top             =   705
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
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
               TabIndex        =   9
               Top             =   30
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   83034113
               CurrentDate     =   37460
            End
            Begin MSComCtl2.DTPicker DtpPagtoParcelaUnica 
               Height          =   315
               Left            =   4860
               TabIndex        =   10
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   83034113
               CurrentDate     =   37460
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
               TabIndex        =   64
               Top             =   390
               Width           =   1725
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
               TabIndex        =   35
               Top             =   750
               Width           =   1365
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
               TabIndex        =   34
               Top             =   60
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
               TabIndex        =   33
               Top             =   45
               Width           =   1245
            End
         End
         Begin MSMask.MaskEdBox Mskconta 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
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
            Bindings        =   "FrmGlosaDevolucao.frx":06EE
            Height          =   330
            Left            =   2880
            TabIndex        =   8
            Top             =   1170
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
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
            Bindings        =   "FrmGlosaDevolucao.frx":0705
            Height          =   330
            Left            =   -73080
            TabIndex        =   12
            Top             =   1430
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
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
            TabIndex        =   5
            Top             =   705
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr 
            Bindings        =   "FrmGlosaDevolucao.frx":0726
            Height          =   330
            Left            =   2880
            TabIndex        =   6
            Top             =   705
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            Height          =   285
            Left            =   -74850
            TabIndex        =   23
            Top             =   1470
            Width           =   1725
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid2 
         Height          =   3255
         Left            =   -74805
         TabIndex        =   27
         Top             =   975
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
         Columns(3).Caption=   "CodStatus"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Status"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CodTipoPagto"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Tipo Pag."
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Doc. Pag."
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Data Pag."
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "dd/mm/yy"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "CodConta"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Conta Corrente (Empresa)"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "CodBancoforn"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Conta Corrente (Fornecedor)"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Cód. Duplicata"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   4
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
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
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1049362"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(24)=   "Column(2)._MinWidth=1953261894"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=2064"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=1049362"
         Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(32)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(34)=   "Column(4).Width=2143"
         Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=2064"
         Splits(0)._ColumnProps(37)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=1057296"
         Splits(0)._ColumnProps(39)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(40)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(41)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(42)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=1048848"
         Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(49)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(50)=   "Column(6).Width=1984"
         Splits(0)._ColumnProps(51)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(6)._WidthInPix=1905"
         Splits(0)._ColumnProps(53)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(6)._ColStyle=1057296"
         Splits(0)._ColumnProps(55)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(56)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(57)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(58)=   "Column(7).Width=1720"
         Splits(0)._ColumnProps(59)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(7)._WidthInPix=1640"
         Splits(0)._ColumnProps(61)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(7)._ColStyle=1057296"
         Splits(0)._ColumnProps(63)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(64)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(65)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(66)=   "Column(8).Width=1535"
         Splits(0)._ColumnProps(67)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(68)=   "Column(8)._WidthInPix=1455"
         Splits(0)._ColumnProps(69)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(70)=   "Column(8)._ColStyle=1049105"
         Splits(0)._ColumnProps(71)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(72)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(73)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(74)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(76)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(9)._ColStyle=1048848"
         Splits(0)._ColumnProps(78)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(79)=   "Column(9).WrapText=1"
         Splits(0)._ColumnProps(80)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(81)=   "Column(10).Width=3863"
         Splits(0)._ColumnProps(82)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(10)._WidthInPix=3784"
         Splits(0)._ColumnProps(84)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(85)=   "Column(10)._ColStyle=1057296"
         Splits(0)._ColumnProps(86)=   "Column(10).WrapText=1"
         Splits(0)._ColumnProps(87)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(88)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(89)=   "Column(11).Width=1270"
         Splits(0)._ColumnProps(90)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(11)._WidthInPix=1191"
         Splits(0)._ColumnProps(92)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(93)=   "Column(11)._ColStyle=1048848"
         Splits(0)._ColumnProps(94)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(95)=   "Column(11).WrapText=1"
         Splits(0)._ColumnProps(96)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(97)=   "Column(12).Width=3916"
         Splits(0)._ColumnProps(98)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(12)._WidthInPix=3836"
         Splits(0)._ColumnProps(100)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(101)=   "Column(12)._ColStyle=1057296"
         Splits(0)._ColumnProps(102)=   "Column(12).WrapText=1"
         Splits(0)._ColumnProps(103)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(104)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(105)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(106)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(108)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(109)=   "Column(13)._ColStyle=1048848"
         Splits(0)._ColumnProps(110)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(111)=   "Column(13).WrapText=1"
         Splits(0)._ColumnProps(112)=   "Column(13).Order=14"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(37)  =   ":id=28,.locked=-1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(42)  =   ":id=32,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2,.locked=0"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=1,.locked=0"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47,.locked=0"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(55)  =   ":id=66,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(64)  =   ":id=74,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(69)  =   ":id=78,.locked=-1"
         _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
         _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
         _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=2,.valignment=2,.locked=0"
         _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
         _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
         _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44"
         _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
         _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
         _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(82)  =   ":id=90,.locked=-1"
         _StyleDefs(83)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44,.alignment=2"
         _StyleDefs(84)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
         _StyleDefs(85)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
         _StyleDefs(86)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(87)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44"
         _StyleDefs(88)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
         _StyleDefs(89)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
         _StyleDefs(90)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(91)  =   ":id=98,.locked=-1"
         _StyleDefs(92)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44,.alignment=2"
         _StyleDefs(93)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
         _StyleDefs(94)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
         _StyleDefs(95)  =   "Splits(0).Columns(13).Style:id=102,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(96)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=44"
         _StyleDefs(97)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=45"
         _StyleDefs(98)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=47"
         _StyleDefs(99)  =   "Named:id=33:Normal"
         _StyleDefs(100) =   ":id=33,.parent=0"
         _StyleDefs(101) =   "Named:id=34:Heading"
         _StyleDefs(102) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(103) =   ":id=34,.wraptext=-1"
         _StyleDefs(104) =   "Named:id=35:Footing"
         _StyleDefs(105) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(106) =   "Named:id=36:Selected"
         _StyleDefs(107) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(108) =   "Named:id=37:Caption"
         _StyleDefs(109) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(110) =   "Named:id=38:HighlightRow"
         _StyleDefs(111) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(112) =   "Named:id=39:EvenRow"
         _StyleDefs(113) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(114) =   "Named:id=40:OddRow"
         _StyleDefs(115) =   ":id=40,.parent=33"
         _StyleDefs(116) =   "Named:id=41:RecordSelector"
         _StyleDefs(117) =   ":id=41,.parent=34"
         _StyleDefs(118) =   "Named:id=42:FilterBar"
         _StyleDefs(119) =   ":id=42,.parent=33"
      End
      Begin Threed.SSPanel PnlDados1 
         Height          =   360
         Left            =   -74280
         TabIndex        =   30
         Top             =   510
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
         TabIndex        =   31
         Top             =   510
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
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   3135
         Left            =   -74595
         TabIndex        =   36
         Top             =   975
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
         Columns(5).Caption=   "Percentual"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Percent"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   4
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(43)=   "Column(4)._MinWidth=86573588"
         Splits(0)._ColumnProps(44)=   "Column(5).Width=2408"
         Splits(0)._ColumnProps(45)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(5)._WidthInPix=2328"
         Splits(0)._ColumnProps(47)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(48)=   "Column(5)._ColStyle=1049106"
         Splits(0)._ColumnProps(49)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(50)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(51)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=20,.parent=43"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=44"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=47"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=45"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=47"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=45"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=47"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=47"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin Threed.SSPanel PnlDados3 
         Height          =   360
         Left            =   -74595
         TabIndex        =   38
         Top             =   495
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
         Left            =   -68355
         TabIndex        =   39
         Top             =   495
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
      Begin VB.Label LblValorNota 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72795
         TabIndex        =   43
         Top             =   4335
         Width           =   1485
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
         Left            =   -74115
         TabIndex        =   42
         Top             =   4380
         Width           =   1410
      End
      Begin VB.Label LblValorTotalDistrib 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -69195
         TabIndex        =   41
         Top             =   4335
         Width           =   1485
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
         Left            =   -70275
         TabIndex        =   40
         Top             =   4380
         Width           =   1170
      End
      Begin VB.Label LblValorTotal 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67965
         TabIndex        =   29
         Top             =   4350
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
         TabIndex        =   28
         Top             =   4395
         Width           =   1170
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
         TabIndex        =   26
         Top             =   4395
         Width           =   1050
      End
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   2775
      TabIndex        =   16
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
      Left            =   5310
      TabIndex        =   17
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
End
Attribute VB_Name = "FrmGlosaDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ValorTotalDuplicata As Double 'Acumula o total de duplicatas até compor o valor total da Nota Fiscal
Dim ValorTotalDistribuicao As Double 'Acumula o total de distribuições até compor o valor total da Nota Fiscal
Dim Xtipomov As String
Dim XDistribuido As String 'verifica se o valor total foi distribuido
Dim xNumFat As Long
Dim Tipo As String 'Verifica o tipo do documento
Dim XTipoFatura As String '1=(1 nota 1 dupl) 2=(1 nota 1 ou + dupl) 3=(2 ou + notas p/ 1 ou + dupl)
Dim XValorBruto As Double  'Armazena o valor bruto da nota
Dim Status As String 'Armazena o Status da Nota Fiscal
Dim XVlDuplicata As Double 'Valor inicial da duplicata que está sendo alterada
Dim NumDocAnt As String 'armazena o nº do documento no caso de uma alteração
Dim XLF_GLOSADEVANT As Double

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor2 As New XArray
Dim Vetor3 As New XArray
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

Dim ERRLOOP As ADODB.Error
Dim STRERROR As String

Function VerificaValores() As Boolean
    Dim Somadupl As Currency
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
    
    If Vetor2(0, 0) <> Empty Then
        Somadupl = funSomaColuna(Vetor2, 2)
        If Format(CCur(TxtDevGlosa.Text), "standard") <> Format(Somadupl, "standard") Then
            If EntSaida = "E" Then
                MsgBox "A soma das Parcelas é diferente do valor da Glosa!", vbCritical, "ATENÇÃO"
            Else
                MsgBox "A soma das Parcelas é diferente do valor da Devolução!", vbCritical, "ATENÇÃO"
            End If

            TabNotasFiscais.Tab = 1
            VerificaValores = True
            Exit Function
        End If
    Else 'entrou na orelha, mas não cadasrou nenhuma duplicata
        MsgBox "A composição da Duplicata é obrigatória.", vbCritical, "ATENÇÃO"
        VerificaValores = True
        Exit Function
    End If
End Function
Sub InsereRegistroVetorDupl(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)  'Carrega o vetores
    Dim Cont As Integer
    Dim X As Integer
    
    'Insere um regitro no vetor
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
        Vetor.Insert 1, Vetor.UpperBound(1) + 1
    End If
    Cont = Vetor.UpperBound(1)
    For X = 0 To Grid.Columns.Count - 2
        subSetaVal Vetor, Cont, X, FunNulo(VetorCampos(X))
    Next
    With Grid
        .ReBind
    End With
End Sub


Sub HabilitaControles()
    If Tipo = "1" Then 'gera duplicata automaticamente
        TabAuxiliar.Enabled = True
        TabAuxiliar.TabEnabled(1) = True
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
    Else     ' Não gera duplicata automaticamente
        TabAuxiliar.Enabled = True
        TabAuxiliar.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(1) = True
        TabNotasFiscais.TabEnabled(2) = True
    End If

    If Tipo = "1" Then  'gera duplicata automaticamente
        TabAuxiliar.Enabled = True
        TabAuxiliar.TabEnabled(1) = True
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
    Else     ' Não gera duplicata automaticamente
        TabAuxiliar.Enabled = True
        TabAuxiliar.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(1) = True
        TabNotasFiscais.TabEnabled(2) = True
    End If
End Sub

Sub PreencheCampos()
    Dim Sql As String
    Dim Status As String
    Dim ResDoc As Object
    Dim ResDist As Object
    Dim RESDUPL As Object
    Dim ResNota As Object
    Dim X As Integer
    
    SubQOpenRecordset ResNota, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal =" & Chave, "Estatico"
    If Not (ResNota.BOF And ResNota.EOF) Then
        XTipoFatura = ResNota!nofi_tx_tipofatura
        XValorBruto = ResNota!nofi_vl_valor
        Status = ResNota!nofi_tx_status
        CboTipoDoc.BoundText = FunNulo(ResNota!tido_cd_TipoDoc)
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        
        If Not IsNull(ResNota!focl_cd_forncli) Then
            CboFornecedor.BoundText = ResNota!focl_cd_forncli
        End If
        If EntSaida = "S" Then
            subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
        End If
        If ResNota!nofi_tx_distribuir = "S" Then
            ChkDistribuicao.Value = 1
            XDistribuido = "S"
        Else
            XDistribuido = "N"
        End If
        
        MskNumdocumento.Text = FunNulo(ResNota!nofi_nr_documento)
        NumDocAnt = MskNumdocumento.Text
        DtpEntrada.Value = Format$(ResNota!nofi_dt_entrada, "DD/MM/YY")
        DtpEmissao.Value = Format$(ResNota!nofi_dt_emissao, "DD/MM/YY")
        TxtSerie.Text = FunNulo(ResNota!nofi_tx_serie)
        CboHistorico.Text = FunNulo(ResNota!nofi_tx_historico)
        
        If Not IsNull(ResNota!nofi_vl_valor) Then
            TxtValor.Text = Format$(ResNota!nofi_vl_valor, "standard")
        End If
    End If
    ResNota.Close
    
    'Dados da Glosa
    If XInserir = "A" Then
        SubQOpenRecordset ResNota, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal =" & ChaveGlosaDev, "Estatico"
        If Not (ResNota.BOF And ResNota.EOF) Then
            XTipoFatura = ResNota!nofi_tx_tipofatura
            XValorBruto = ResNota!nofi_vl_valor
            If ResNota!nofi_tx_distribuir = "S" Then
                ChkDistribuicao.Value = 1
                XDistribuido = "S"
            Else
                XDistribuido = "N"
            End If
            
            MskDocGlosaDev.Text = FunNulo(ResNota!nofi_nr_documento)
            NumDocAnt = MskNumdocumento.Text
            DtpCompGlosaDev.Value = Format$(ResNota!nofi_dt_entrada, "DD/MM/YY")
            DtpEmissaoGlosaDev.Value = Format$(ResNota!nofi_dt_emissao, "DD/MM/YY")
            ChkDistribuicao.Enabled = True
            CboHistGlosaDev.Text = FunNulo(ResNota!nofi_tx_historico)
            
            If Not IsNull(ResNota!nofi_vl_valor) Then
                TxtDevGlosa.Text = Format$(ResNota!nofi_vl_valor, "standard")
            End If
            
            Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & ResNota!nofi_nr_fatura
            subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
            
            Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & ChaveGlosaDev & " and nofi_cd_notafiscal = " & ChaveGlosaDev & ""
            subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "percentual"), Vetor3, TDBGrid3
            
            XLF_GLOSADEVANT = TxtDevGlosa.Text
        End If
        ResNota.Close
    End If
        
    If Status = "P" Or XDistribuido = "N" Then 'É previsto ou não foi realizada distribuição e sim apropriação
        SubQOpenRecordset ResDist, "SELECT Distribuicao.*,PlanoContas.plco_cd_codreduzido, PlanoContas.plco_tx_desccontabil,CentroResultados.cere_cd_estruturado ,CentroResultados.cere_tx_descricao FROM ((Distribuicao LEFT JOIN PlanoContas ON Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta) LEFT JOIN CentroResultados ON Distribuicao.cere_cd_pcr = CentroResultados.cere_cd_pcr) where nofi_cd_notafiscal =" & Chave, "Estatico"
        If Not (ResDist.BOF And ResDist.EOF) Then
           MskPcr.Text = ResDist!cere_cd_estruturado
           CboPcr.BoundText = MskPcr.Text
           MskConta.Text = ResDist!plco_cd_CodReduzido
           CboPlanoContas.BoundText = MskConta.Text
        End If
        ResDist.Close
    End If
    
    SubQOpenRecordset ResDoc, "SELECT tipodocumentos.tido_cd_tipodoc,tipodocumentos.tido_tx_duplicata FROM TipoDocumentos where TipoDocumentos.tido_cd_tipodoc = " & CboTipoDoc.BoundText, "Estatico"
    If Not ResDoc.BOF And Not ResDoc.BOF Then
       Tipo = ResDoc!tido_tx_duplicata
    End If
    ResDoc.Close

    If Status = "P" Or Tipo = "1" Then 'É previsto ou gera duplicata automaticamente
        SubQOpenRecordset RESDUPL, "SELECT Duplicatas.*,ContasCorrente.coco_tx_conta,ContasCorrente.coco_cd_agencia FROM ((Duplicatas LEFT JOIN BancoFornecedor ON Duplicatas.bafo_cd_codigo = BancoFornecedor.bafo_cd_codigo) LEFT JOIN ContasCorrente ON Duplicatas.coco_cd_codigo = ContasCorrente.coco_cd_codigo) where dupl_nr_fatura =" & NFatura, "Estatico"
        If Not (RESDUPL.BOF And RESDUPL.EOF) Then
             DtpVencParcelaUnica.Value = Format$(RESDUPL!dupl_dt_vencimento, "DD/MM/YY")
             If Status = "2" Then
                DtpPagtoParcelaUnica.Value = Format$(RESDUPL!dupl_dt_pagrec, "DD/MM/YY")
             End If
             If Not IsNull(RESDUPL!dupl_tx_tipopagrec) Then
                For X = 0 To CboTipoPagto1.ListCount - 1
                    If CboTipoPagto1.ItemData(X) = CInt(RESDUPL!dupl_tx_tipopagrec) Then
                        CboTipoPagto1.ListIndex = X
                    End If
                Next
             End If
             'TxtDocPagto1.Text = FunNulo(ResDupl!dupl_tx_docpagrec)
             If Not IsNull(RESDUPL!coco_cd_codigo) Then
                CboCCorrente.BoundText = RESDUPL!coco_cd_codigo
             End If
             If Not IsNull(RESDUPL!bafo_cd_codigo) Then
                CboBancoForn1.BoundText = RESDUPL!bafo_cd_codigo
             End If
       End If
       RESDUPL.Close
    End If
End Sub

Function ChecarCamposOrelhaPrincipal() As Boolean
    ChecarCamposOrelhaPrincipal = True
    If TabNotasFiscais.Tab = 0 Then
        If IsNull(DtpEmissaoGlosaDev.Value) Then
            MsgBox "Data de Emissão é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
            Exit Function
        End If
        
        If XDistribuido = "N" And ChkDistribuicao.Value = 0 Then
            If FunObrigatorioCBO(CboPcr, "Centro Custo é Obrigatório") Then Exit Function
            If FunObrigatorioMSK(MskConta, "Conta Contábil é Obrigatória") Then Exit Function
        End If
        
        If Tipo = "1" Then 'Gera duplicata automaticamente
            If IsNull(DtpVencParcelaUnica.Value) Then
                MsgBox "O preenchimento do campo Vencimento é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
            If IsNull(DtpPagtoParcelaUnica.Value) Then
                MsgBox "Data do Pagamento é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO"
                Exit Function
            End If
            If EntSaida = "E" Then
                If FunObrigatorioCBOX(CboTipoPagto1, "O preenchimento do campo Tipo Recebimento é obrigatório.") Then Exit Function
            Else
                If FunObrigatorioCBOX(CboTipoPagto1, "O preenchimento do campo Tipo Pagamento é obrigatório.") Then Exit Function
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
        
        MskPcr.Mask = ""
        MskPcr.Text = ""
        MskPcr.Mask = pPCRmascara
        CboPcr.Text = ""
        MskConta.Mask = ""
        MskConta.Text = ""
        CboPlanoContas.Text = ""
        
        DtpVencParcelaUnica.Value = Null
        DtpPagtoParcelaUnica.Value = Null
        If Not TelaNotasFiscais.Adodc1.Recordset.EOF Then
            If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_TipoPag")) Then
                CboTipoPagto1.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_TipoPag")
            Else
                CboTipoPagto1.ListIndex = 0
            End If
        End If
        CboBancoForn1.Text = ""
        CboCCorrente.Text = ""
   End If
End Sub

Private Sub CboBancoForn1_Change()
    If CboBancoForn1.Text <> "" Then
        DatBancoFornecedor.Recordset.Bookmark = CboBancoForn1.SelectedItem
    End If
End Sub

Private Sub CboCCorrente_Change()
    If CboCCorrente.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    End If
End Sub

Private Sub CboPcr_Change()
    If CboPcr.Text <> "" Then
        DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            Exit Sub
        End If
        MskPcr.Text = CboPcr.BoundText
    End If
End Sub

Private Sub cboPlanoContas_Change()
    If CboPlanoContas.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = MskConta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        MskConta.Text = CboPlanoContas.BoundText
    End If
End Sub

Private Sub CboTipoPagto1_CLICK()
    Dim Achou As Boolean
    
    Achou = False
    If EntSaida = "S" Then
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

Private Sub ChkDistribuicao_Click()
    If ChkDistribuicao.Value = 1 Then
       MskPcr.Mask = ""
       MskPcr.Text = ""
       MskPcr.Mask = pPCRmascara
       MskConta.Text = ""
       CboPlanoContas.Text = ""
       CboPcr.Text = ""
       TabAuxiliar.TabEnabled(0) = False
       MskPcr.Enabled = False
       CboPcr.Enabled = False
       MskConta.Enabled = False
       CboPlanoContas.Enabled = False
       TabNotasFiscais.TabEnabled(2) = True
    Else
       MskConta.Text = ""
       CboPlanoContas.Text = ""
       CboPcr.Text = ""
       TabAuxiliar.TabEnabled(0) = True
       MskPcr.Enabled = True
       CboPcr.Enabled = True
       MskConta.Enabled = True
       CboPlanoContas.Enabled = True
       TabNotasFiscais.TabEnabled(2) = False
       TabNotasFiscais.Tab = 0
    End If
    XDistribuido = "N"
End Sub

Private Sub CmdDesistir_Click()
    If XFormulario = "TelaNotasFiscais" Then
        If Not FunTabelaVazia(TelaNotasFiscais.Adodc1) Then
          TelaNotasFiscais.CmdAlterar.Enabled = True
          TelaNotasFiscais.CmdExcluir.Enabled = True
          TelaNotasFiscais.CmdImprimir.Enabled = True
          TelaNotasFiscais.CmdOrdem.Enabled = True
'          TelaNotasFiscais.CmdDuplicata.Enabled = True
'          TelaNotasFiscais.CmdRecibo.Enabled = True
          TelaNotasFiscais.CmdSair.Enabled = True
        End If
    ElseIf XFormulario = "TelaDuplicatas" Then
        TelaDuplicatas.Adodc1.Recordset.CancelUpdate
    ElseIf XFormulario = "TelaFatura" Then
        TelaFatura.Adodc1.Recordset.CancelUpdate
    End If
    Unload Me
End Sub

Private Sub CmdExcluiduplicata_Click()
   Dim Houve_Erro As Integer
   Dim ValorExcluido As Double

   Houve_Erro = False
   TDBGrid2.Enabled = False
   TDBGrid2.MoveFirst
   
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        ValorExcluido = TDBGrid2.Columns("Valor").Value
        subRemoveRegistroVetor Vetor2, TDBGrid2
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

Private Sub cmdGravar_Click()
    Dim Sql As String
    Dim XStatus As String
    Dim XSerie As String
    Dim ResNF As Object
    Dim X As Byte
    Dim XNumDoc As String
    Dim XCodDupl As String
    Dim ResFat As Object
    Dim XLO_NOTA As Object
    Dim DataVencimento As Date
    Dim XLI_TIPODOC As Integer
    Dim XLT_DOCPAGREC   'Armazena o nº do doc. pag/rec
    Dim XLT_DEBCREDGLOSA As String
       
    TabNotasFiscais.Tab = 0
    If ChecarCamposOrelhaPrincipal Then Exit Sub
    
    If TxtDevGlosa.Text = "" Then
        If EntSaida = "S" Then MsgBox "Informe o valor da devolução.", vbCritical, "ATENÇÃO"
        If EntSaida = "E" Then MsgBox "Informe o valor da glosa.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    '***************CONTABILIDADE*******************
    If XGB_PQUALICONT And XGT_TIPOEXPORTACAO = "PC" Then  'Se possui o módulo de contabilidade
        If funMesFechado(DtpCompGlosaDev.Value) <> "A" Then
            MsgBox "Não é possível cadastrar/alterar o documento nessa data, pois o mês contábil já foi fechado. " _
                + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If
    '***********************************************
    
    If XDistribuido = "N" And (MskPcr.Text = "" And MskConta.Text = "") Then
        MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
        Exit Sub
    ElseIf ChkDistribuicao.Value = 1 Then
        If TxtDevGlosa.Text <> funSomaColuna(Vetor3, 4) Then
            If EntSaida = "E" Then
                MsgBox "A soma das distribuições é diferente do valor da Glosa!", vbCritical, "ATENÇÃO"
            Else
                MsgBox "A soma das distribuições é diferente do valor da Devolução!", vbCritical, "ATENÇÃO"
            End If
            Exit Sub
        End If
    End If
    
    If Tipo = 0 Then 'Não gera duplicata automaticamente.
        'Verifica se o total da nota é igual a soma das duplicatas com os descontos ou se alterou o valor da nota
        If VerificaValores Then
            Exit Sub
        End If
        
        'Verifica se a posição do grid do doc.pag/rec está vazia, se sim gera o nº do doc.
        TDBGrid2.MoveFirst
        If TDBGrid2.Columns(7).Text = "" Then
            While Not TDBGrid2.EOF
                TDBGrid2.Columns(7).Text = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
                TDBGrid2.Update
        
                TDBGrid2.MoveNext
            Wend
        End If
    Else
        XCodDupl = "U"
        XLT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    End If
       
    If EntSaida = "E" Then
        Xtipomov = "E"
        
        'Tipo Documento
        SubQOpenRecordset XLO_NOTA, "SELECT tido_cd_TipoDoc FROM TipoDocumentos WHERE tido_tx_Sigla = 'GLOS'", Estatico
        XLI_TIPODOC = XLO_NOTA!tido_cd_TipoDoc
        XLO_NOTA.Close
    ElseIf EntSaida = "S" Then
        Xtipomov = "S"
        
        'Tipo Documento
        SubQOpenRecordset XLO_NOTA, "SELECT tido_cd_TipoDoc FROM TipoDocumentos WHERE tido_tx_Sigla = 'NFDE'", Estatico
        XLI_TIPODOC = XLO_NOTA!tido_cd_TipoDoc
        XLO_NOTA.Close
    End If

'-------------------------------------------------------------------
    'Definição de XTipoFatura
    If IsNull(DtpVencParcelaUnica.Value) Then
        XTipoFatura = 2    '2 - Uma nota para várias duplicatas
    Else
        XTipoFatura = 1
    End If
'-------------------------------------------------------------------

    'Definir o código da duplicata como "U" para caso de unica e
    '"A" para caso de ser a primeira de uma série
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
        SubQOpenRecordset XLO_NOTA, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        XLO_NOTA.AddNew
    Else    'Alteração
        SubQOpenRecordset XLO_NOTA, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal = " & ChaveGlosaDev, Dinamico
        xNumFat = XLO_NOTA!nofi_nr_fatura
    End If
    
    'Gera nº da fatura
    If XInserir = "I" Then
        xNumFat = FunPegaGuardaUltimo("F")
    Else 'NA Alteração não precisa gerar novo nº de fatura
        xNumFat = XLO_NOTA!nofi_nr_fatura
    End If
    
    'Gera Novo nº documento
    If MskDocGlosaDev.Text = "" Then
        XNumDoc = Format(FunPegaGuardaUltimo("D"), "A000000000")
    Else
        XNumDoc = MskDocGlosaDev.Text
    End If
    
    If Xtipomov = "E" Then
        XLO_NOTA!nofi_tx_tipomov = "S"
        XLT_DEBCREDGLOSA = "D"
    Else
        XLO_NOTA!nofi_tx_tipomov = "E"
        XLT_DEBCREDGLOSA = "C"
    End If
    
    XLO_NOTA!nofi_tx_status = "R"
    XLO_NOTA!nofi_nr_documento = XNumDoc
    XLO_NOTA!focl_cd_forncli = CboFornecedor.BoundText
    XLO_NOTA!nofi_dt_emissao = Format$(DtpEmissaoGlosaDev.Value, "DD/MM/YY")
    XLO_NOTA!nofi_tx_serie = TxtSerie.Text
    XLO_NOTA!tido_cd_TipoDoc = XLI_TIPODOC
    XLO_NOTA!nofi_dt_entrada = DtpCompGlosaDev.Value
    XLO_NOTA!nofi_tx_historico = FunNulo(CboHistGlosaDev.Text)
    XLO_NOTA!nofi_vl_valor = FunNuloVal(TxtDevGlosa.Text)
    XLO_NOTA!nofi_tx_distribuir = XDistribuido
    XLO_NOTA!nofi_nr_fatura = xNumFat
    XLO_NOTA!nofi_tx_tipofatura = XTipoFatura
    XLO_NOTA!empr_cd_empresa = PCodEmpresa
    XLO_NOTA!nofi_cd_notaGlosaDev = Chave
    
    XLO_NOTA.Update
    If XInserir = "I" Then
        ChaveGlosaDev = XLO_NOTA!nofi_cd_notafiscal
    End If
    XLO_NOTA.Close
    
    'Duplicatas
    If Tipo = 0 Then
        If XInserir = "A" Then
            Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & xNumFat
            subGravaVetorGlosa "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 9, Array(0, 1, 2, 3, 5, 7, 8, 9, 11), Str(xNumFat) & ",'" & XLT_DEBCREDGLOSA & "'", Vetor2, TDBGrid2, "P"
        End If
        If Vetor2(0, 0) <> Empty And XInserir = "I" Then
            subGravaVetorGlosa "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 9, Array(0, 1, 2, 3, 5, 7, 8, 9, 11), Str(xNumFat) & ",'" & XLT_DEBCREDGLOSA & "'", Vetor2, TDBGrid2, "P"
        End If
    Else
        If XInserir = "I" Then
            Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                  " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevGlosa.Text)) & ",'R','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & XLT_DEBCREDGLOSA & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
            Conexao.Execute Sql
            
            SubQOpenRecordset ResNF, "select dupl_vl_Valor, dupl_vl_GlosaDevolucao from Duplicatas where dupl_nr_fatura = (select nofi_nr_fatura from NotasFiscais where nofi_cd_notaFiscal = " & Chave & ")", Dinamico 'código da duplicata de origem
            ResNF!dupl_vl_valor = (ResNF!dupl_vl_valor - CDbl(TxtDevGlosa.Text))
            ResNF!dupl_vl_GlosaDevolucao = CDbl(TxtDevGlosa.Text)
            ResNF.Update
            ResNF.Close
        ElseIf XInserir = "A" And (Tipo = "1" Or XStatus = "P") Then
            Conexao.Execute "UPDATE Duplicatas " & _
                                "SET dupl_nr_duplicata ='" & XCodDupl & "', " & _
                                "dupl_dt_vencimento = " & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & _
                                "dupl_vl_valor = " & FunNuloVal(FunTrataFloat(TxtDevGlosa.Text)) & "," & _
                                "dupl_tx_status = 'R'," & _
                                "dupl_tx_tipopagrec = '" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "'," & _
                                "dupl_tx_docpagrec = '" & FunNulo(XLT_DOCPAGREC) & "'," & _
                                "coco_cd_codigo= " & FunNuloValor(CboCCorrente) & "," & _
                                "bafo_cd_codigo = " & FunNuloValor(CboBancoForn1) & "," & _
                                "dupl_tx_debcred ='" & XLT_DEBCREDGLOSA & "'," & _
                                "dupl_nr_fatura = " & Str(xNumFat) & "," & _
                                "dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & " " & _
                                "WHERE dupl_nr_fatura = " & xNumFat
                                
            SubQOpenRecordset ResNF, "select dupl_vl_Valor, dupl_vl_GlosaDevolucao from Duplicatas where dupl_nr_fatura = (select nofi_nr_fatura from NotasFiscais where nofi_cd_notaFiscal = " & Chave & ")", Dinamico 'código da duplicata de origem
            ResNF!dupl_vl_valor = ((ResNF!dupl_vl_valor + ResNF!dupl_vl_GlosaDevolucao) - CDbl(TxtDevGlosa.Text))
            ResNF!dupl_vl_GlosaDevolucao = CDbl(TxtDevGlosa.Text)
            ResNF.Update
            ResNF.Close
        End If
   End If
        
    'Distribuição
    If ChkDistribuicao Then
        If XInserir = "A" Then
              Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & ChaveGlosaDev
        End If
        If Vetor3(0, 0) <> Empty Then
            subGravaVetorGlosa "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_valor,nofi_cd_notafiscal", Array("T", "T", "N", "N"), 3, Array(0, 2, 4), "" & ChaveGlosaDev & "", Vetor3, TDBGrid3, "D"
        Else
            'se entrou na orelha de distribuicao excluiu o rateio e cadastrou na orelha de apropriação
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,nofi_cd_notafiscal)" & _
                " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat(TxtValor.Text)) & "," & ChaveGlosaDev & ")"
            Conexao.Execute Sql
            
            SubQOpenRecordset ResNF, "SELECT dist_vl_Valor, dist_vl_GlosaDevolucao FROM Distribuicao WHERE nofi_cd_notafiscal = " & Chave & "", Dinamico 'código da distribuicao de origem
            ResNF!dist_vl_valor = ((ResNF!dist_vl_valor + ResNF!dist_vl_GlosaDevolucao) - CDbl(TxtDevGlosa.Text))
            ResNF!dist_vl_GlosaDevolucao = CDbl(TxtDevGlosa.Text)
            ResNF.Update
            ResNF.Close
        End If
    Else
        If XInserir = "A" Then
            Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & ChaveGlosaDev
        End If
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,nofi_cd_notafiscal)" & _
                  " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat(TxtDevGlosa.Text)) & "," & ChaveGlosaDev & ")"
            Conexao.Execute Sql
            
            SubQOpenRecordset ResNF, "SELECT dist_vl_Valor, dist_vl_GlosaDevolucao FROM Distribuicao WHERE nofi_cd_notafiscal = " & Chave & "", Dinamico 'código da distribuicao de origem
            ResNF!dist_vl_valor = ((ResNF!dist_vl_valor + ResNF!dist_vl_GlosaDevolucao) - CDbl(TxtDevGlosa.Text))
            ResNF!dist_vl_GlosaDevolucao = CDbl(TxtDevGlosa.Text)
            ResNF.Update
            ResNF.Close
        'End If
    End If
    
    'Update na tabela de Notas Fiscais para gravar o nº da glosa
    Conexao.Execute "UPDATE NotasFiscais SET nofi_cd_NotaGlosaDev = " & ChaveGlosaDev & " WHERE nofi_cd_notafiscal = " & Chave & ""
    
    '***************CONTABILIDADE*******************
    If XGT_TIPOEXPORTACAO = "PC" Then  'Exporta base provisão
        Sql = "select * from lancamentos where (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T') and lanc_cd_OrigemFin = " & Chave & ""
        SubQOpenRecordset ResNF, Sql, Estatico
        
        If Not ResNF.EOF Then 'A devolução só será contabilizada se tiver ocorrido a provisão
'            SubProvisao xNumFat, DtpEntrada.Value
        End If
    End If
    '***********************************************
    
    Conexao.CommitTrans
    
    
    '****************REGISTRA LOG DA OPERAÇÃO***********************
     'Nota Fiscal:
     ReDim XGM_MATRIZLOG(11, 2) As Variant
     
     'Nota de origem
     XGM_MATRIZLOG(0, 0) = "Nota Fiscal Origem"
     XGM_MATRIZLOG(1, 0) = "Fornecedor"
     XGM_MATRIZLOG(2, 0) = "Data Competencia"
     XGM_MATRIZLOG(3, 0) = "Data Emissao"
     XGM_MATRIZLOG(4, 0) = "Valor"
     If CboHistorico.Text <> "" Then XGM_MATRIZLOG(5, 0) = "Historico"
     XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
     XGM_MATRIZLOG(1, 1) = CboFornecedor.Text
     XGM_MATRIZLOG(2, 1) = Format$(DtpEntrada.Value, "DD/MM/YYYY")
     XGM_MATRIZLOG(3, 1) = Format$(DtpEmissao.Value, "DD/MM/YYYY")
     XGM_MATRIZLOG(4, 1) = Format$(TxtValor.Text, "##,##0.00")
     XGM_MATRIZLOG(5, 1) = FunNulo(CboHistorico.Text)
     'Nota de devolução
     XGM_MATRIZLOG(6, 0) = "Nota Fiscal Devolucao"
     XGM_MATRIZLOG(7, 0) = "Data Competencia"
     XGM_MATRIZLOG(8, 0) = "Data Emissao"
     XGM_MATRIZLOG(9, 0) = "Valor"
     If CboHistGlosaDev.Text <> "" Then XGM_MATRIZLOG(10, 0) = "Historico"
     XGM_MATRIZLOG(6, 1) = XNumDoc
     XGM_MATRIZLOG(7, 1) = Format$(DtpCompGlosaDev.Value, "DD/MM/YYYY")
     XGM_MATRIZLOG(8, 1) = Format$(DtpEmissaoGlosaDev.Value, "DD/MM/YYYY")
     XGM_MATRIZLOG(9, 1) = Format$(TxtDevGlosa.Text, "##,##0.00")
     XGM_MATRIZLOG(10, 1) = FunNulo(CboHistGlosaDev.Text)
     
     If EntSaida = "E" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, GLOSA, NOTAFISCAL_RECEBER))
     If EntSaida = "S" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, DEVOLVER, NOTAFISCAL_PAGAR))
     
     
    'Duplicata:
      ReDim XGM_MATRIZLOG(9, 2) As Variant
    
      XGM_MATRIZLOG(0, 0) = "Descrição"
      XGM_MATRIZLOG(1, 0) = "Nota Fiscal" 'Número do documento da nota fiscal
      XGM_MATRIZLOG(2, 0) = "Cod"
      XGM_MATRIZLOG(3, 0) = "Valor"
      XGM_MATRIZLOG(4, 0) = "Status"
      XGM_MATRIZLOG(5, 0) = "Data Venc"
      
      XGM_MATRIZLOG(8, 0) = "Conta Corrente"
      
      XGM_MATRIZLOG(0, 1) = "Duplicata ref. a glosa/devolução"
      XGM_MATRIZLOG(1, 1) = XNumDoc
      
      TDBGrid2.MoveFirst
      If Tipo <> 0 Then 'composicao de duplicata
          XGM_MATRIZLOG(2, 1) = "U"
          XGM_MATRIZLOG(3, 1) = Format$(TxtDevGlosa.Text, "##,##0.00")
          XGM_MATRIZLOG(4, 1) = "realizado"
          XGM_MATRIZLOG(5, 1) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
          XGM_MATRIZLOG(6, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
          XGM_MATRIZLOG(7, 1) = CboTipoPagto1.List(CboTipoPagto1.ListIndex)
          XGM_MATRIZLOG(8, 1) = CboCCorrente.Text
      Else 'apropriação
          XGM_MATRIZLOG(2, 1) = TDBGrid2.Columns("Código").Text
          XGM_MATRIZLOG(3, 1) = Format$(TDBGrid2.Columns("valor").Text, "##,##0.00")
          XGM_MATRIZLOG(4, 1) = TDBGrid2.Columns("status").Text
          XGM_MATRIZLOG(5, 1) = Format$(TDBGrid2.Columns("vencimento").Text, "DD/MM/YYYY")
          If EntSaida = "S" Then
                XGM_MATRIZLOG(6, 0) = "Data Pagto"
                XGM_MATRIZLOG(7, 0) = "Tipo Pagto"
                
                XGM_MATRIZLOG(6, 1) = Format$(TDBGrid2.Columns("Data Pag.").Text, "DD/MM/YYYY")
                XGM_MATRIZLOG(7, 1) = TDBGrid2.Columns("Tipo Pag.").Text
          Else
                XGM_MATRIZLOG(6, 0) = "Data Rec."
                XGM_MATRIZLOG(7, 0) = "Tipo Rec."
                XGM_MATRIZLOG(6, 1) = Format$(TDBGrid2.Columns("Data Rec.").Text, "DD/MM/YYYY")
                XGM_MATRIZLOG(7, 1) = TDBGrid2.Columns("Tipo Rec.").Text
          End If
          
          XGM_MATRIZLOG(8, 1) = TDBGrid2.Columns("Conta Corrente (Empresa)").Text
      End If
    
      If EntSaida = "E" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_RECEBER))
      If EntSaida = "S" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_PAGAR))
    
    
    'Distribuicao:
      TDBGrid3.MoveFirst
      
      While Not TDBGrid3.BOF
          ReDim XGM_MATRIZLOG(6, 2) As Variant
            
          XGM_MATRIZLOG(0, 0) = "Descricao"
          XGM_MATRIZLOG(1, 0) = "Nota Fiscal"
          XGM_MATRIZLOG(2, 0) = "Centro Custo"
          XGM_MATRIZLOG(3, 0) = "Conta Contábil"
          XGM_MATRIZLOG(4, 0) = "Percentual"
          XGM_MATRIZLOG(5, 0) = "Valor"
          
          XGM_MATRIZLOG(0, 1) = "Distribuição de glosa/devolução"
          XGM_MATRIZLOG(1, 1) = XNumDoc
          
          If ChkDistribuicao Then
              XGM_MATRIZLOG(2, 1) = TDBGrid3.Columns("Centro Custo").Text
              XGM_MATRIZLOG(3, 1) = TDBGrid3.Columns("Conta Contábil").Text
              XGM_MATRIZLOG(4, 1) = TDBGrid3.Columns("Percentual")
              XGM_MATRIZLOG(5, 1) = Format$(TDBGrid3.Columns("Valor").Text, "##,##0.00")
          Else
              XGM_MATRIZLOG(2, 1) = CboPcr.Text
              XGM_MATRIZLOG(3, 1) = CboPlanoContas.Text
              XGM_MATRIZLOG(4, 1) = "100,00%"
              XGM_MATRIZLOG(5, 1) = Format$(TxtDevGlosa.Text, "##,##0.00")
          End If
          
         If EntSaida = "E" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_RECEBER))
         If EntSaida = "S" Then Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_PAGAR))
         
         TDBGrid3.MoveNext
    Wend
    '****************FIM REGISTRA LOG DA OPERAÇÃO********************
    
    Call CmdDesistir_Click
    
    Exit Sub
TrataErro:
Conexao.RollbackTrans

    funTrataErros (ComMensagem)
   ' Enumera a coleção de erros e apresenta
   ' as propriedades de cada erro ocorrido
   For Each ERRLOOP In Conexao.Errors
      STRERROR = "Error #" & ERRLOOP.Number & vbCr & _
         "   " & ERRLOOP.Description & vbCr & _
         "   (Source: " & ERRLOOP.Source & ")" & vbCr & _
         "   (SQL State: " & ERRLOOP.SQLState & ")" & vbCr & _
         "   (NativeError: " & ERRLOOP.NativeError & ")" & vbCr

   MsgBox STRERROR, vbCritical + vbOKOnly, "ATENÇÃO !!!"
   Next

   Resume Next
End Sub


Sub subGravaVetorGlosa(Tabela As String, CamposTabela As String, Tipos As Variant, QtdCampos As Byte, Colunas As Variant, Complemento As String, Vetor As XArray, Grid As TDBGrid, XLT_ParcDistrib As String)
    Dim Valores As String
    Dim X As Integer, Y As Integer
    Dim XLO_REGISTRO As Object
    'Tabela - nome da tabela onde serão inseridos os registros do grid
    'CamposTabela - string simples com os campos na respectiva ordem do grid
    
    'Tipos - vetor simples com os respectivos tipos de dados das colunas do grid
    'N: Número; D: Data; T: Texto
    
    'QtdCampos - quantidade de colunas do grid que serão gravadas
    'Colunas - vetor simples com o índice das respectivas colunas que serão gravadas
    'Complemento - Valores de campos fixos que deseje gravar e não existe no grid para evitar
    'repetição Ex.: Todos os campos possuem o mesmo código de serviço
    
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    If Vetor(0, 0) <> Empty Then
        If Complemento <> "" Then
            Complemento = ", " & Complemento
        End If
        
        For X = 0 To Vetor.UpperBound(1)
            Valores = ""
            For Y = 0 To QtdCampos - 1
                If Valores <> "" Then
                    If Tipos(Y) = "N" Then
                        Valores = Valores & ", " & FunNuloBancoVal(Vetor(X, Colunas(Y)))
                    ElseIf Tipos(Y) = "T" Then
                        Valores = Valores & ", '" & Vetor(X, Colunas(Y)) & "'"
                    ElseIf Tipos(Y) = "D" Then
                        Valores = Valores & "," & FunNuloData(Vetor(X, Colunas(Y)), NomeSgbd) & ""
                    End If
                Else
                    If Tipos(Y) = "N" Then
                        Valores = FunNuloBancoVal(Vetor(X, Colunas(Y)))
                    ElseIf Tipos(Y) = "T" Then
                        Valores = "'" & Vetor(X, Colunas(Y)) & "'"
                    ElseIf Tipos(Y) = "D" Then
                        Valores = "" & FunNuloData(Vetor(X, Colunas(Y)), NomeSgbd) & ""
                    End If
                End If
            Next
    
            If XInserir = "I" Then
                'Atualizar a origem (Duplicatas/Distribuicao)
                If XLT_ParcDistrib = "P" Then 'Parcela
                    SubQOpenRecordset XLO_REGISTRO, "SELECT dupl_vl_valor, dupl_vl_GlosaDevolucao FROM Duplicatas WHERE dupl_cd_duplicata = " & Vetor(X, 13) & "", Dinamico 'código da duplicata de origem
                    XLO_REGISTRO!dupl_vl_valor = (XLO_REGISTRO!dupl_vl_valor - CDbl(Vetor(X, 2)))
                    XLO_REGISTRO!dupl_vl_GlosaDevolucao = CDbl(Vetor(X, 2))
                Else
                    SubQOpenRecordset XLO_REGISTRO, "SELECT dist_vl_Valor, dist_vl_GlosaDevolucao FROM Distribuicao WHERE nofi_cd_notafiscal = " & Chave & " and cere_cd_pcr = " & Vetor(X, 0) & " and plco_cd_conta =" & Vetor(X, 2) & "", Dinamico 'código da distribuicao de origem
                    XLO_REGISTRO!dist_vl_valor = (XLO_REGISTRO!dist_vl_valor - CDbl(Vetor(X, 4)))
                    XLO_REGISTRO!dist_vl_GlosaDevolucao = CDbl(Vetor(X, 4))
                End If
                XLO_REGISTRO.Update
                XLO_REGISTRO.Close
            End If
            
            Conexao.Execute "INSERT INTO " & Tabela & " (" & CamposTabela & ") VALUES (" & Valores & Complemento & ")"
        Next
    End If
End Sub

Function ConsistenciaDatas() As Boolean
'Verifica se as datas seguem a ordem de preenchimento
'Se ConsistenciaDatas= TRUE    >> Datas consistentes
'Se ConsistenciaDatas= FALSE  >> Datas não consistentes
    ConsistenciaDatas = False

    If TabNotasFiscais.Tab = 0 Then
        If Not IsNull(DtpVencParcelaUnica.Value) And Not IsNull(DtpEmissao.Value) Then
            If DateDiff("d", DtpEmissao.Value, DtpVencParcelaUnica.Value) < 0 Then
                MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                DtpVencParcelaUnica.SetFocus
                Exit Function
            End If
        End If
        If Not IsNull(DtpPagtoParcelaUnica.Value) Then
            If DateDiff("d", DtpEmissao.Value, DtpPagtoParcelaUnica.Value) < 0 Then
                MsgBox "A data de pagamento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                DtpPagtoParcelaUnica.SetFocus
                Exit Function
            End If
        End If
    End If
    ConsistenciaDatas = True
End Function

Private Sub CmdLimparHistorico_Click()
    CboHistorico.BoundText = 0
    CboHistorico.Text = ""
End Sub




Private Sub CmdRemoverDistribuicao_Click()
    Dim Houve_Erro As Integer
    Dim XLF_RESIDUO As Currency
    Dim XLF_TOTAL As Currency

    Houve_Erro = False
       
    TDBGrid3.Enabled = False
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor3, TDBGrid3
    End If

    If Vetor3(0, 0) <> Empty Then
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        XLF_TOTAL = CCur(LblValorTotalDistrib.Caption)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    Else
        LblValorTotalDistrib.Caption = ""
        ValorTotalDistribuicao = 0
    End If
    
    If CCur(TxtDevGlosa.Text) <> XLF_TOTAL Then
        XLF_RESIDUO = (CCur(TxtDevGlosa.Text) - XLF_TOTAL) / Vetor3.Count(1)
    
        TDBGrid3.MoveFirst
        While Not TDBGrid3.EOF
            TDBGrid3.Columns("Valor").Text = Format(CDbl(FunNuloVal(TDBGrid3.Columns("Valor").Text)) + XLF_RESIDUO, "standard")
            TDBGrid3.Update
    
            TDBGrid3.MoveNext
        Wend
        
        XLF_TOTAL = CCur(funSomaColuna(Vetor3, 4))
        If CCur(TxtDevGlosa.Text) < XLF_TOTAL Then
            TDBGrid3.MovePrevious
            TDBGrid3.Columns("Valor").Text = CCur(FunNuloVal(TDBGrid3.Columns("Valor").Text)) - (XLF_TOTAL - CCur(TxtDevGlosa.Text))
            TDBGrid3.Update
        ElseIf CCur(TxtDevGlosa.Text) > XLF_TOTAL Then
            TDBGrid3.MovePrevious
            TDBGrid3.Columns("Valor").Text = CCur(FunNuloVal(TDBGrid3.Columns("Valor").Text)) + (CCur(TxtDevGlosa.Text) - XLF_TOTAL)
            TDBGrid3.Update
        End If
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
    End If
   
    TDBGrid3.Enabled = True
 
    Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next

End Sub


Private Sub DtpCompGlosaDev_LostFocus()
    If DtpCompGlosaDev.Value < DtpEntrada.Value Then
        MsgBox "A data de entrada da Glosa é menor que a data de entrada da Nota de origem!", vbCritical, "ATENÇÃO"
        DtpCompGlosaDev.SetFocus
        Exit Sub
    End If
End Sub


Private Sub DtpEmissaoGlosaDev_LostFocus()
    If DtpEmissaoGlosaDev.Value < DtpEmissao.Value Then
        MsgBox "A data de emissão da Glosa é menor que a data de emissão da Nota de origem!", vbCritical, "ATENÇÃO"
        DtpEmissaoGlosaDev.SetFocus
        Exit Sub
    End If
End Sub


Private Sub Form_Activate()
    Set Formulario = FrmGlosaDevolucao
    MskDocGlosaDev.SetFocus
End Sub

Private Sub Form_Load()
    Dim ResFat As Object
    'On Error GoTo TrataErro
    
    'PrimeiraVez = True
    subManutencaoJanelasAtivas "I", "FrmGlosaDevolucao"
    
    Set Formulario = FrmGlosaDevolucao
          
    If XInserir = "I" Then
        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS' or tido_tx_sigla='RGLO') ORDER BY tido_tx_descricao", Estatico
        CmdGravar.Enabled = True
    Else
        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos ORDER BY tido_tx_descricao", Estatico
        CmdGravar.Enabled = False
    End If
    subConectarControleDadosNV DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY focl_tx_razaosocial ", Estatico
    
    'Comum às orelhas de dados gerais e duplicatas
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
    'Comum às orelhas de dados gerais e distribuição
    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    
    If EntSaida = "E" Then
        LblNomeFantasia.Caption = "Cliente:"
        LblDevGlosa.Caption = "Valor Glosa:"
        PanNotasFiscais.Caption = "Dados da Glosa"
        FrmGlosaDevolucao.Caption = "Cadastro de Glosa"
        subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
        subConectarControleDadosNV DatHistorico2, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        LblTipoPagto.Caption = "Tipo Recebimento:"
        LblBcForn.Visible = False
        CboBancoForn1.Visible = False
        LblPagto.Caption = "Recebimento:"
        PanPagto1.Top = 570
        
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        CboTipoPagto1.RemoveItem (1)
        
        TDBGrid2.Columns("Tipo Pag.").Caption = "Tipo Rec."
        TDBGrid2.Columns("Doc. Pag.").Caption = "Doc. Rec."
        TDBGrid2.Columns("Data Pag.").Caption = "Data Rec."
        TDBGrid2.Columns("Conta Corrente (Fornecedor)").Visible = False
        TabAuxiliar.TabCaption(1) = "Recebimento"
    ElseIf EntSaida = "S" Then
        LblNomeFantasia.Caption = "Fornecedor:"
        subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
        subConectarControleDadosNV DatHistorico2, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
       
        CboTipoPagto1.RemoveItem (5)
    End If
     
    DtpCompGlosaDev.Value = Format(Now, "dd/MM/yy")
    DtpEmissaoGlosaDev.Value = Format(Now, "dd/MM/yy")
    CboHistGlosaDev.Text = ""
    EntrouDesc = False
    EntrouDupl = False
    EntrouDist = False
    MskPcr.Mask = pPCRmascara
    NumDocAnt = ""

    XDistribuido = "N"
    PreencheCampos
    If XInserir = "I" Then
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
        LblValorTotal.Caption = ""
        LblValorTotalDistrib.Caption = ""
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
    subManutencaoJanelasAtivas "R", "FrmGlosaDevolucao"
End Sub

Private Sub MskConta_GotFocus()
    Call subSelecionaMSK(MskConta)
End Sub
Private Sub mskConta_LostFocus()
    If MskConta.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskConta.Text = Format(MskConta.Text, XGT_MSKCODREDUZIDO)
        End If
        CboPlanoContas.BoundText = MskConta.Text
        If CboPlanoContas.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta.SetFocus
        End If
    
    End If
End Sub

Private Sub dtpEmissao_LostFocus()
    If Not ConsistenciaDatas Then Exit Sub
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
        If Len(MskPcr.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            MskPcr.Text = funFormataCCusto(MskPcr.Text)
        End If
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


Private Sub dtpVencParcelaUnica_LostFocus()
    If IsDate(DtpVencParcelaUnica.Value) Then
        If Not ConsistenciaDatas Then Exit Sub
    End If
End Sub

Private Sub TabAuxiliar_Click(PreviousTab As Integer)
'    If XInserir = "I" Then
'        If IsNull(DtpVencParcelaUnica.Value) Then DtpVencParcelaUnica.Value = Format$(Now, "dd/MM/yy")
'        DtpVencParcelaUnica.SetFocus
'        ' If CboStatus.Text = "Realizado" Then
'        If IsNull(DtpPagtoParcelaUnica.Value) Then DtpPagtoParcelaUnica.Value = Format$(Now, "dd/mm/yy")
'        ' End If
'    End If
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
    Dim X As Integer
    
    If TabNotasFiscais.Tab <> 0 And XFormulario = "TelaNotasFiscais" Then
        If ChecarCamposOrelhaPrincipal Then
            TabNotasFiscais.Tab = 0
            Exit Sub
        End If
    End If
      
    '****Preenche o cabecalho das orelhas com o nome do Forn/cliente, nº doc/Serie****
    PnlDados1_1.Caption = (" DOC.: " & MskNumdocumento.Text & " / " & TxtSerie.Text)
    PnlDados3_3.Caption = (" DOC.: " & MskNumdocumento.Text & " / " & TxtSerie.Text)
    If EntSaida = "S" Then
        PnlDados1.Caption = (" FORNECEDOR: " & CboFornecedor.Text)
        PnlDados3.Caption = (" FORNECEDOR: " & CboFornecedor.Text)
    Else
        PnlDados1.Caption = (" CLIENTE: " & CboFornecedor.Text)
        PnlDados3.Caption = (" CLIENTE: " & CboFornecedor.Text)
    End If
    '*********************************************************
      
    If TabNotasFiscais.Tab = 1 And Not EntrouDupl Then
        EntrouDupl = True
        LblValorTotal = funSomaColuna(Vetor2, 2)
        ValorTotalDuplicata = CDbl(LblValorTotal.Caption)
    ElseIf TabNotasFiscais.Tab = 2 And Not EntrouDist Then
        EntrouDist = True
        LblValorNota.Caption = Format$(TxtValor.Text, "##,##0.00")
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    End If
End Sub

Private Sub TDBGrid2_AfterColEdit(ByVal ColIndex As Integer)
    TDBGrid2.Update
    LblValorTotal = funSomaColuna(Vetor2, 2)
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    TDBGrid2.Update
    LblValorTotal = funSomaColuna(Vetor2, 2)
End Sub

Private Sub TDBGrid3_AfterColEdit(ByVal ColIndex As Integer)
    TDBGrid3.Columns("Percentual").Text = CDbl(TDBGrid3.Columns("Valor").Text) / CDbl(TxtDevGlosa.Text)
    TDBGrid3.Update
End Sub

Private Sub TDBGrid3_LostFocus()
    If CDbl(TxtDevGlosa.Text) <> 0 Then
        TDBGrid3.Columns("Percentual").Text = CDbl(TDBGrid3.Columns("Valor").Text) / CDbl(TxtDevGlosa.Text)
        TDBGrid3.Update
    End If
End Sub

Private Sub TxtDevGlosa_LostFocus()
    Dim XLI_X As Integer
    Dim XLF_TOTAL As Currency
    Dim XLF_RESIDUO As Currency
    
    If TxtDevGlosa.Text <> "" Then
        If CDbl(TxtDevGlosa.Text) > CDbl(TxtValor.Text) Then
            If EntSaida = "E" Then
                MsgBox "O valor da Glosa é maior que o valor da Nota.", vbCritical, "ATENÇÃO"
            Else
                MsgBox "O valor da Devolução é maior que o valor da Nota.", vbCritical, "ATENÇÃO"
            End If
            TxtDevGlosa.SetFocus
            Exit Sub
        End If
         
        If IsNumeric(TxtDevGlosa.Text) Then
            TxtDevGlosa.Text = Format$(TxtDevGlosa.Text, "##,##0.00")
        Else
            TxtDevGlosa.Text = ""
        End If
         
        If XInserir = "I" Then
            If EntSaida = "E" Then
                'Quando for Glosa não ratear para parcelas realizadas
                Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_tx_status <> 'R' and dupl_nr_fatura = " & NFatura
            Else
                'Quando for Devolução pode ratear para parcelas realizadas
                Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
            End If
            
            subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
            
            Sql = "SELECT Distribuicao.*,Notasfiscais.nofi_vl_valor,(Distribuicao.dist_vl_valor / NotasFiscais.nofi_vl_valor) AS Percentual,CentroResultados.cere_cd_estruturado, PlanoContas.plco_cd_codreduzido " & _
            "FROM NotasFiscais,Distribuicao,CentroResultados, PlanoContas WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " and Distribuicao.nofi_cd_notafiscal = " & Chave & " and Distribuicao.cere_cd_pcr =CentroResultados.cere_cd_pcr and Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta"
            
            subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
        Else
            If XLF_GLOSADEVANT <> TxtDevGlosa.Text Then
                Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
                subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2
            
                Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " and nofi_cd_notafiscal = " & Chave & ""
                subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "percentual"), Vetor3, TDBGrid3
                XLF_GLOSADEVANT = TxtDevGlosa.Text
            Else
                Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
                subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "dupl_vl_GlosaDevolucao", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata"), Vetor2, TDBGrid2

                Sql = "SELECT Distribuicao.*,Notasfiscais.nofi_vl_valor,(Distribuicao.dist_vl_GlosaDevolucao / NotasFiscais.nofi_vl_valor) AS Percentual,CentroResultados.cere_cd_estruturado, PlanoContas.plco_cd_codreduzido " & _
                "FROM NotasFiscais,Distribuicao,CentroResultados, PlanoContas WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " and Distribuicao.nofi_cd_notafiscal = " & Chave & " and Distribuicao.cere_cd_pcr =CentroResultados.cere_cd_pcr and Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta"
                subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
            End If
        End If
        
        TDBGrid2.Enabled = True
        TDBGrid3.Enabled = True
           
        If Tipo = 0 Then
            TDBGrid2.MoveFirst
            While Not TDBGrid2.EOF
                TDBGrid2.Columns("Valor").Text = Format((CDbl(TxtDevGlosa.Text) / CDbl(TxtValor.Text)) * CDbl(FunNuloVal(TDBGrid2.Columns("Valor").Text)), "standard")
                TDBGrid2.Columns("CodStatus").Text = "R"
                TDBGrid2.Columns("Status").Text = "Realizado"
                TDBGrid2.Columns(5).Text = 6
                TDBGrid2.Columns(6).Text = "Dinheiro"
                TDBGrid2.Columns(8).Text = Format(Now, "dd/MM/yy")
                TDBGrid2.Update
                TDBGrid2.MoveNext
            Wend
            
            TabNotasFiscais.TabEnabled(1) = True
            TabAuxiliar.TabEnabled(1) = False
        
            XLF_TOTAL = CCur(funSomaColuna(Vetor2, 2))
            
            If CCur(TxtDevGlosa.Text) <> XLF_TOTAL Then
                XLF_RESIDUO = (CCur(TxtDevGlosa.Text) - XLF_TOTAL) / Vetor2.Count(1)
            
                TDBGrid2.MoveFirst
                While Not TDBGrid2.EOF
                    TDBGrid2.Columns("Valor").Text = Format(CDbl(FunNuloVal(TDBGrid2.Columns("Valor").Text)) + XLF_RESIDUO, "standard")
                    TDBGrid2.Update
                    TDBGrid2.MoveNext
                Wend
                
                XLF_TOTAL = CCur(funSomaColuna(Vetor2, 2))
                
                If CCur(TxtDevGlosa.Text) < XLF_TOTAL Then
                    TDBGrid2.MovePrevious
                    TDBGrid2.Columns("Valor").Text = CCur(FunNuloVal(TDBGrid2.Columns("Valor").Text)) - (XLF_TOTAL - CCur(TxtDevGlosa.Text))
                ElseIf CCur(TxtDevGlosa.Text) > XLF_TOTAL Then
                    TDBGrid2.MovePrevious
                    TDBGrid2.Columns("Valor").Text = CCur(FunNuloVal(TDBGrid2.Columns("Valor").Text)) + (CCur(TxtDevGlosa.Text) - XLF_TOTAL)
                End If
            End If
        Else
            TabNotasFiscais.TabEnabled(1) = False
            TabAuxiliar.Enabled = True
            TabAuxiliar.TabEnabled(1) = True
            CboTipoPagto1.Text = "Dinheiro"
            CboTipoPagto1.Enabled = False
            DtpPagtoParcelaUnica.Value = Format(Now, "dd/MM/yy")
        End If
        
        If ChkDistribuicao.Value = 1 Then
            TDBGrid3.MoveFirst
            While Not TDBGrid3.EOF
                TDBGrid3.Columns("Valor").Text = (CDbl(TxtDevGlosa.Text) / CDbl(TxtValor.Text)) * CDbl(TDBGrid3.Columns("Valor").Text)
                TDBGrid3.Columns("Percentual").Text = CDbl(TDBGrid3.Columns("Valor").Text) / CDbl(TxtDevGlosa.Text)
                TDBGrid3.Update
                
                TDBGrid3.MoveNext
            Wend
            TabNotasFiscais.TabEnabled(2) = True
            TabAuxiliar.TabEnabled(0) = False
        Else
            TabAuxiliar.Enabled = True
            TabAuxiliar.TabEnabled(0) = True
            TabNotasFiscais.TabEnabled(2) = False
        End If
        
        'Não habilita a aba de distribuição porque o usuário não pode alterá-la.
        'A devolução deve gerar uma provisão inversa à provisão da nota original.
        If EntSaida = "S" Then
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            MskConta.Enabled = False
            CboPlanoContas.Enabled = False
            CmdRemoverDistribuicao.Enabled = False
        End If
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValor_LostFocus()
    If IsNumeric(TxtValor.Text) Then
        TxtValor.Text = Format$(TxtValor.Text, "##,##0.00")
    Else
        TxtValor.Text = ""
    End If
End Sub
