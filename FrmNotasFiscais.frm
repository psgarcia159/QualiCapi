VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmNotasFiscais 
   Caption         =   "Cadastro de Contas a Pagar"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   1305
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc DatContaContabilCredito 
      Height          =   375
      Left            =   120
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adodc2"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkPendente 
      Caption         =   "Pendente"
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
      TabIndex        =   156
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   57
      Top             =   6435
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TabDlg.SSTab TabNotasFiscais 
      Height          =   6345
      Left            =   240
      TabIndex        =   60
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11192
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
      Tab(0).Control(2)=   "PanNotasFiscais"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraDadosGerais"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSPanel1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Composição de Duplicata"
      TabPicture(1)   =   "FrmNotasFiscais.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "LblValorTotal"
      Tab(1).Control(2)=   "CmdDuplMultiplas"
      Tab(1).Control(3)=   "PnlDados1_1"
      Tab(1).Control(4)=   "PnlDados1"
      Tab(1).Control(5)=   "TDBGrid2"
      Tab(1).Control(6)=   "CmdExcluiduplicata"
      Tab(1).Control(7)=   "CmdAlteraduplicata"
      Tab(1).Control(8)=   "CmdInsereduplicata"
      Tab(1).Control(9)=   "PnlCompDuplicata"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Descontos"
      TabPicture(2)   =   "FrmNotasFiscais.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PnlDados2_2"
      Tab(2).Control(1)=   "PnlDados2"
      Tab(2).Control(2)=   "TDBGrid1"
      Tab(2).Control(3)=   "pandescontos"
      Tab(2).Control(4)=   "CmdInsereDesconto"
      Tab(2).Control(5)=   "CmdExcluiDesconto"
      Tab(2).Control(6)=   "CmdAlteraDesconto"
      Tab(2).Control(7)=   "CmdExcluirTodos"
      Tab(2).ControlCount=   8
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
      Tab(3).Control(9)=   "CmdCriterio"
      Tab(3).Control(10)=   "TDBGrid3"
      Tab(3).Control(11)=   "PainelCriterios"
      Tab(3).Control(12)=   "PanDistribuicao"
      Tab(3).Control(13)=   "CmdImportarPlanilha"
      Tab(3).ControlCount=   14
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   120
         TabIndex        =   67
         Top             =   1460
         Width           =   9615
         _Version        =   65536
         _ExtentX        =   16960
         _ExtentY        =   979
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
         Begin VB.TextBox TxtCodTipoDoc 
            Height          =   315
            Left            =   4875
            MaxLength       =   4
            TabIndex        =   1
            Top             =   120
            Width           =   495
         End
         Begin VB.ComboBox CboStatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmNotasFiscais.frx":0070
            Left            =   840
            List            =   "FrmNotasFiscais.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2415
         End
         Begin MSAdodcLib.Adodc DatTipoDoc 
            Height          =   330
            Left            =   6360
            Top             =   120
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
            Bindings        =   "FrmNotasFiscais.frx":0097
            Height          =   315
            Left            =   5400
            TabIndex        =   2
            Top             =   120
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
            Left            =   3375
            TabIndex        =   69
            Top             =   165
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
            Left            =   165
            TabIndex        =   68
            Top             =   150
            Width           =   645
         End
      End
      Begin VB.Frame FraDadosGerais 
         Height          =   945
         Left            =   120
         TabIndex        =   174
         Top             =   480
         Width           =   9615
         Begin MSDataListLib.DataCombo CboImovel 
            Bindings        =   "FrmNotasFiscais.frx":00B0
            Height          =   315
            Left            =   1620
            TabIndex        =   175
            Top             =   540
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "Imovel"
            BoundColumn     =   "imov_cd_imovel"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MskEmpreendimento 
            Height          =   315
            Left            =   1650
            TabIndex        =   176
            Top             =   165
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboEmpreendimento 
            Bindings        =   "FrmNotasFiscais.frx":00C8
            Height          =   315
            Left            =   2370
            TabIndex        =   177
            Top             =   165
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "empd_tx_nome"
            BoundColumn     =   "empd_cd_empreendimento"
            Text            =   "CboEmpreendimento"
         End
         Begin MSAdodcLib.Adodc DatEmpreendimento 
            Height          =   330
            Left            =   2580
            Top             =   225
            Visible         =   0   'False
            Width           =   2280
            _ExtentX        =   4022
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
            Caption         =   "DatEmpreendimento"
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
         Begin MSAdodcLib.Adodc DatImovel 
            Height          =   330
            Left            =   2880
            Top             =   480
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
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
            Caption         =   "DatImovel"
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
         Begin VB.Label LblEmpreendimento 
            Alignment       =   1  'Right Justify
            Caption         =   "Empreendimento:"
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
            Left            =   60
            TabIndex        =   179
            Top             =   210
            Width           =   1545
         End
         Begin VB.Label LblImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Imóvel:"
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
            Left            =   780
            TabIndex        =   178
            Top             =   600
            Width           =   810
         End
      End
      Begin Threed.SSPanel PanNotasFiscais 
         Height          =   1905
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   9615
         _Version        =   65536
         _ExtentX        =   16960
         _ExtentY        =   3360
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
         Begin VB.TextBox TxtHistorico 
            Height          =   315
            Left            =   1455
            MaxLength       =   4
            TabIndex        =   11
            Top             =   1350
            Width           =   450
         End
         Begin VB.TextBox TxtPrazo 
            Height          =   315
            Left            =   5880
            MaxLength       =   3
            TabIndex        =   6
            Top             =   495
            Width           =   480
         End
         Begin VB.TextBox TxtValor 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   7
            Top             =   480
            Width           =   1800
         End
         Begin VB.TextBox TxtDevolucao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7215
            TabIndex        =   145
            Text            =   "0"
            Top             =   495
            Width           =   1755
         End
         Begin VB.TextBox TxtSerie 
            Height          =   315
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   5
            Top             =   510
            Width           =   495
         End
         Begin MSAdodcLib.Adodc DatFornecedor 
            Height          =   330
            Left            =   2880
            Top             =   105
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
            Left            =   2745
            Top             =   1350
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
         Begin Threed.SSCommand CmdLimparHistorico 
            Height          =   315
            Left            =   8670
            TabIndex        =   97
            Top             =   1350
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            _StockProps     =   78
            Enabled         =   0   'False
            MouseIcon       =   "FrmNotasFiscais.frx":00E8
            Picture         =   "FrmNotasFiscais.frx":053A
         End
         Begin MSComCtl2.DTPicker DtpEmissao 
            Height          =   315
            Left            =   4080
            TabIndex        =   9
            Top             =   930
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   81657857
            CurrentDate     =   37460
         End
         Begin MSDataListLib.DataCombo CboFornecedor 
            Bindings        =   "FrmNotasFiscais.frx":064C
            Height          =   330
            Left            =   1440
            TabIndex        =   3
            Top             =   120
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "DESC_FORNCLIENTE"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   "CboFornecedor"
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
         Begin MSMask.MaskEdBox MskNumdocumento 
            Height          =   315
            Left            =   1455
            TabIndex        =   4
            Top             =   525
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboHistorico 
            Bindings        =   "FrmNotasFiscais.frx":0668
            Height          =   315
            Left            =   1965
            TabIndex        =   12
            Top             =   1350
            Width           =   6690
            _ExtentX        =   11800
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DtpEntrada 
            Height          =   315
            Left            =   1455
            TabIndex        =   8
            Top             =   935
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81657857
            CurrentDate     =   37460
         End
         Begin VB.Label LblPrazo 
            Alignment       =   1  'Right Justify
            Caption         =   "Prazo 1ª Parcela:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4365
            TabIndex        =   151
            Top             =   540
            Width           =   1500
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
            Left            =   6615
            TabIndex        =   65
            Top             =   540
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
            Left            =   6600
            TabIndex        =   146
            Top             =   555
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
            Left            =   5655
            TabIndex        =   86
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label LblValorLiquido1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   7215
            TabIndex        =   10
            Top             =   930
            Width           =   1800
         End
         Begin VB.Label LblEntrada 
            Alignment       =   1  'Right Justify
            Caption         =   "Dt. Entrada:"
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
            Left            =   240
            TabIndex        =   71
            Top             =   960
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
            Left            =   3180
            TabIndex        =   70
            Top             =   960
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
            Left            =   2880
            TabIndex        =   66
            Top             =   540
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
            Left            =   60
            TabIndex        =   64
            Top             =   540
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
            Left            =   300
            TabIndex        =   63
            Top             =   1365
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
            Left            =   300
            TabIndex        =   62
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.CommandButton CmdExcluirTodos 
         Caption         =   "Excluir Todos"
         Height          =   365
         Left            =   -67080
         TabIndex        =   165
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton CmdImportarPlanilha 
         Caption         =   "Importar"
         Height          =   330
         Left            =   -66885
         TabIndex        =   164
         Top             =   2215
         Width           =   1095
      End
      Begin Threed.SSPanel PanDistribuicao 
         Height          =   1845
         Left            =   -74760
         TabIndex        =   76
         Top             =   1320
         Visible         =   0   'False
         Width           =   9420
         _Version        =   65536
         _ExtentX        =   16616
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
         Enabled         =   0   'False
         Begin VB.TextBox TxtPercDistribuicao 
            Height          =   330
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   52
            Top             =   990
            Width           =   660
         End
         Begin VB.TextBox TxtVldistribuicao 
            Height          =   330
            Left            =   2865
            TabIndex        =   53
            Top             =   990
            Width           =   1455
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
            Left            =   7800
            TabIndex        =   56
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
            Left            =   6600
            TabIndex        =   55
            Top             =   1380
            Width           =   1140
         End
         Begin MSMask.MaskEdBox MskConta2 
            Height          =   330
            Left            =   1530
            TabIndex        =   50
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContas2 
            Bindings        =   "FrmNotasFiscais.frx":0683
            Height          =   330
            Left            =   2400
            TabIndex        =   51
            Top             =   600
            Width           =   6570
            _ExtentX        =   11589
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
         Begin MSMask.MaskEdBox MskPcr2 
            Height          =   375
            Left            =   1530
            TabIndex        =   48
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr2 
            Bindings        =   "FrmNotasFiscais.frx":069A
            Height          =   330
            Left            =   2400
            TabIndex        =   49
            Top             =   120
            Width           =   6570
            _ExtentX        =   11589
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
         Begin MSDataListLib.DataCombo CboIdentificador 
            Bindings        =   "FrmNotasFiscais.frx":06AF
            Height          =   315
            Left            =   5535
            TabIndex        =   54
            Top             =   990
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "iden_tx_descricao"
            BoundColumn     =   "iden_cd_Projeto"
            Text            =   ""
         End
         Begin Threed.SSCommand CmdLimparIdentif2 
            Height          =   315
            Left            =   8610
            TabIndex        =   150
            Top             =   990
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":06CE
            Picture         =   "FrmNotasFiscais.frx":0B20
         End
         Begin Threed.SSCommand CmdLimparCR2 
            Height          =   315
            Left            =   9000
            TabIndex        =   154
            Top             =   180
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":0C32
            Picture         =   "FrmNotasFiscais.frx":1084
         End
         Begin Threed.SSCommand CmdLimparCC2 
            Height          =   315
            Left            =   9000
            TabIndex        =   155
            Top             =   570
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":1196
            Picture         =   "FrmNotasFiscais.frx":15E8
         End
         Begin VB.Label LblIdentificador 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Identificador:"
            DataField       =   "c"
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
            Left            =   4350
            TabIndex        =   147
            Top             =   1035
            Width           =   1155
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
            Left            =   450
            TabIndex        =   87
            Top             =   1020
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
            Left            =   2040
            TabIndex        =   79
            Top             =   1035
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
            Left            =   -240
            TabIndex        =   78
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
            Left            =   90
            TabIndex        =   77
            Top             =   600
            Width           =   1350
         End
      End
      Begin Threed.SSPanel PainelCriterios 
         Height          =   4215
         Left            =   -74640
         TabIndex        =   157
         Top             =   480
         Visible         =   0   'False
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   7435
         _StockProps     =   15
         Caption         =   "Relação de Critérios"
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
         Begin VB.CommandButton CmdConfirmar 
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
            Left            =   6120
            TabIndex        =   159
            Top             =   3720
            Width           =   1140
         End
         Begin VB.CommandButton Command4 
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
            Left            =   7440
            TabIndex        =   158
            Top             =   3720
            Width           =   1140
         End
         Begin MSAdodcLib.Adodc DatCriterios 
            Height          =   330
            Left            =   2160
            Top             =   3720
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
            Caption         =   "DatCriterios"
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
         Begin MSDataListLib.DataCombo CboCriterios 
            Bindings        =   "FrmNotasFiscais.frx":16FA
            Height          =   330
            Left            =   1080
            TabIndex        =   160
            Top             =   360
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "crdi_tx_descricao"
            BoundColumn     =   "cere_cd_Estruturado"
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   240
            Top             =   3720
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid10 
            Bindings        =   "FrmNotasFiscais.frx":1715
            Height          =   2730
            Left            =   240
            TabIndex        =   163
            Top             =   840
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   4815
            _LayoutType     =   4
            _RowHeight      =   24
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código Estruturado"
            Columns(0).DataField=   "cere_cd_Estruturado"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Centro de Resultado"
            Columns(1).DataField=   "cere_tx_Descricao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   16
            Columns(2)._MaxComboItems=   5
            Columns(2).ValueItems(0)._DefaultItem=   0
            Columns(2).ValueItems(0).Value=   "S"
            Columns(2).ValueItems(0).Value.vt=   8
            Columns(2).ValueItems(0).DisplayValue=   "Sintético"
            Columns(2).ValueItems(0).DisplayValue.vt=   8
            Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems(1)._DefaultItem=   0
            Columns(2).ValueItems(1).Value=   "A"
            Columns(2).ValueItems(1).Value.vt=   8
            Columns(2).ValueItems(1).DisplayValue=   "Analítico"
            Columns(2).ValueItems(1).DisplayValue.vt=   8
            Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems.Count=   2
            Columns(2).Caption=   "Código Reduzido"
            Columns(2).DataField=   "plco_cd_CodReduzido"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Conta Contábil"
            Columns(3).DataField=   "plco_tx_DescContabil"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Percentual"
            Columns(4).DataField=   "cocd_vl_percentual"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(10)=   "Column(1).Width=3519"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3440"
            Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=528"
            Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=55313848"
            Splits(0)._ColumnProps(19)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=2566"
            Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=528"
            Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(25)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(26)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(27)=   "Column(3).Width=2646"
            Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=2566"
            Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=20"
            Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
            Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(34)=   "Column(4).Width=1402"
            Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=1323"
            Splits(0)._ColumnProps(37)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=20"
            Splits(0)._ColumnProps(39)=   "Column(4).WrapText=1"
            Splits(0)._ColumnProps(40)=   "Column(4).Order=5"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1,.bold=0,.fontsize=825"
            _StyleDefs(25)  =   ":id=13,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(28)  =   ":id=22,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=22,.fontname=MS Sans Serif"
            _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2,.bold=0"
            _StyleDefs(42)  =   ":id=25,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(43)  =   ":id=25,.fontname=MS Sans Serif"
            _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2,.bold=0"
            _StyleDefs(48)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(49)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2,.bold=0"
            _StyleDefs(54)  =   ":id=43,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(55)  =   ":id=43,.fontname=MS Sans Serif"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.bold=0,.fontsize=825"
            _StyleDefs(60)  =   ":id=47,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(61)  =   ":id=47,.fontname=MS Sans Serif"
            _StyleDefs(62)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.bold=0,.fontsize=825"
            _StyleDefs(66)  =   ":id=51,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(67)  =   ":id=51,.fontname=MS Sans Serif"
            _StyleDefs(68)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(70)  =   "Named:id=33:Normal"
            _StyleDefs(71)  =   ":id=33,.parent=0"
            _StyleDefs(72)  =   "Named:id=34:Heading"
            _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   ":id=34,.wraptext=-1"
            _StyleDefs(75)  =   "Named:id=35:Footing"
            _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=36:Selected"
            _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=37:Caption"
            _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(81)  =   "Named:id=38:HighlightRow"
            _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=39:EvenRow"
            _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(85)  =   "Named:id=40:OddRow"
            _StyleDefs(86)  =   ":id=40,.parent=33"
            _StyleDefs(87)  =   "Named:id=41:RecordSelector"
            _StyleDefs(88)  =   ":id=41,.parent=34"
            _StyleDefs(89)  =   "Named:id=42:FilterBar"
            _StyleDefs(90)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Critério:"
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
            Left            =   240
            TabIndex        =   162
            Top             =   480
            Width           =   735
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   3135
         Left            =   -74280
         TabIndex        =   88
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
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "CodIdentificador"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Identificador"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1048848"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=2566"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2487"
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
         Splits(0)._ColumnProps(27)=   "Column(3).Width=2619"
         Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=1049104"
         Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(35)=   "Column(3)._MinWidth=86574240"
         Splits(0)._ColumnProps(36)=   "Column(4).Width=2566"
         Splits(0)._ColumnProps(37)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._WidthInPix=2487"
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
         Splits(0)._ColumnProps(61)=   "Column(7).Width=1746"
         Splits(0)._ColumnProps(62)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(64)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(7)._ColStyle=1049106"
         Splits(0)._ColumnProps(66)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(67)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(68)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(69)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(74)=   "Column(8)._ColStyle=1049104"
         Splits(0)._ColumnProps(75)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(76)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(77)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(78)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(79)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(81)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(82)=   "Column(9)._ColStyle=1049104"
         Splits(0)._ColumnProps(83)=   "Column(9).WrapText=1"
         Splits(0)._ColumnProps(84)=   "Column(9).Order=10"
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
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=43,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=44,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=47"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=43"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44,.alignment=2"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=43"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=45"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=47"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton CmdCriterio 
         Caption         =   "&Critério de Rateio"
         Height          =   450
         Left            =   -66885
         TabIndex        =   161
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton CmdAlteraDesconto 
         Caption         =   "A&lterar"
         Height          =   330
         Left            =   -67080
         TabIndex        =   41
         Top             =   1380
         Width           =   1080
      End
      Begin VB.CommandButton CmdAlteraDistribuicao 
         Caption         =   "&Al&terar"
         Height          =   330
         Left            =   -66885
         TabIndex        =   104
         Top             =   1380
         Width           =   1080
      End
      Begin Threed.SSPanel PnlCompDuplicata 
         Height          =   3255
         Left            =   -74400
         TabIndex        =   94
         Top             =   1320
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   5741
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
            Height          =   1875
            Left            =   75
            TabIndex        =   117
            Top             =   105
            Width           =   8145
            _Version        =   65536
            _ExtentX        =   14367
            _ExtentY        =   3307
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
            Begin MSDataListLib.DataCombo CboFormaPagamento2 
               Bindings        =   "FrmNotasFiscais.frx":172A
               Height          =   330
               Left            =   1830
               TabIndex        =   166
               Top             =   1480
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "DescCombo"
               BoundColumn     =   "fopa_cd_codigo"
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
            Begin VB.TextBox TxtCCorrente2 
               Height          =   330
               Left            =   1860
               MaxLength       =   4
               TabIndex        =   33
               Top             =   1050
               Width           =   450
            End
            Begin VB.TextBox TxtDocPagto2 
               Height          =   315
               Left            =   6600
               MaxLength       =   7
               TabIndex        =   32
               Top             =   570
               Width           =   1515
            End
            Begin VB.TextBox TxtDuplicata 
               Height          =   315
               Left            =   1860
               MaxLength       =   4
               TabIndex        =   28
               Top             =   105
               Width           =   630
            End
            Begin VB.TextBox TxtValorDuplicata 
               Height          =   315
               Left            =   3465
               TabIndex        =   29
               Top             =   105
               Width           =   1605
            End
            Begin VB.ComboBox CboTipoPagto2 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":174A
               Left            =   1875
               List            =   "FrmNotasFiscais.frx":1763
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   570
               Width           =   3000
            End
            Begin VB.ComboBox CboStatusduplicata 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":17BD
               Left            =   6180
               List            =   "FrmNotasFiscais.frx":17CA
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   105
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo CboCCorrente2 
               Bindings        =   "FrmNotasFiscais.frx":17F3
               Height          =   330
               Left            =   2340
               TabIndex        =   34
               Top             =   1050
               Width           =   5775
               _ExtentX        =   10186
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
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Forma de Pagto:"
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
               Left            =   240
               TabIndex        =   167
               Top             =   1485
               Width           =   1485
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
               TabIndex        =   123
               Top             =   585
               Width           =   1815
            End
            Begin VB.Label LblCodDupl 
               Alignment       =   1  'Right Justify
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
               Left            =   1065
               TabIndex        =   122
               Top             =   150
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
               TabIndex        =   121
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
               Left            =   60
               TabIndex        =   120
               Top             =   600
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
               TabIndex        =   119
               Top             =   1065
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
               TabIndex        =   118
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
            TabIndex        =   39
            Top             =   2670
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
            TabIndex        =   38
            Top             =   2670
            Width           =   1140
         End
         Begin Threed.SSPanel PanDatas 
            Height          =   495
            Left            =   120
            TabIndex        =   95
            Top             =   2520
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
               TabIndex        =   36
               Top             =   90
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               DateIsNull      =   -1  'True
               Format          =   81657857
               CurrentDate     =   36494
            End
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3855
               TabIndex        =   37
               Top             =   105
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   81657857
               CurrentDate     =   37603
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
               TabIndex        =   98
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
               TabIndex        =   96
               Top             =   120
               Width           =   1245
            End
         End
         Begin MSDataListLib.DataCombo CboBancoForn2 
            Bindings        =   "FrmNotasFiscais.frx":1812
            Height          =   330
            Left            =   1920
            TabIndex        =   35
            Top             =   2040
            Width           =   6270
            _ExtentX        =   11060
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
            TabIndex        =   99
            Top             =   2070
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdInsereduplicata 
         Caption         =   "&Inserir"
         Height          =   330
         Left            =   -74790
         TabIndex        =   25
         Top             =   4785
         Width           =   1080
      End
      Begin VB.CommandButton CmdAlteraduplicata 
         Caption         =   "&Alterar"
         Height          =   330
         Left            =   -73480
         TabIndex        =   26
         Top             =   4785
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiduplicata 
         Caption         =   "&Excluir"
         Height          =   330
         Left            =   -72170
         TabIndex        =   27
         Top             =   4785
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiDesconto 
         Caption         =   "E&xcluir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   42
         Top             =   1800
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDesconto 
         Caption         =   "I&nserir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   40
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton CmdRemoverDistribuicao 
         Caption         =   "Excl&uir"
         Height          =   330
         Left            =   -66885
         TabIndex        =   81
         Top             =   1800
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDistribuicao 
         Caption         =   "Ins&erir"
         Height          =   330
         Left            =   -66885
         TabIndex        =   80
         Top             =   960
         Width           =   1080
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   2295
         Left            =   120
         TabIndex        =   72
         Top             =   3960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Apropriação"
         TabPicture(0)   =   "FrmNotasFiscais.frx":1833
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblPcr"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblConta"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblIdentificador2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblContaCredito"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdLimparIdentCcredito"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cboPlanoContasCredito"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "MskcontaCredito"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CboPlanoContas"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CmdLimparCC"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "CmdLimparCR"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "CmdLimparIdentif"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "CboPcr"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "MskPcr"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Mskconta"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "CboIdentificador2"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Pagamento"
         TabPicture(1)   =   "FrmNotasFiscais.frx":184F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LblBcForn"
         Tab(1).Control(1)=   "CboBancoForn1"
         Tab(1).Control(2)=   "PanPagto1"
         Tab(1).ControlCount=   3
         Begin MSDataListLib.DataCombo CboIdentificador2 
            Bindings        =   "FrmNotasFiscais.frx":186B
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   1395
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "iden_tx_descricao"
            BoundColumn     =   "iden_cd_Projeto"
            Text            =   ""
         End
         Begin Threed.SSPanel PanPagto1 
            Height          =   1440
            Left            =   -74940
            TabIndex        =   111
            Top             =   360
            Width           =   9015
            _Version        =   65536
            _ExtentX        =   15901
            _ExtentY        =   2540
            _StockProps     =   15
            BevelOuter      =   0
            Begin VB.TextBox TxtCCorrente 
               Height          =   330
               Left            =   1860
               MaxLength       =   4
               TabIndex        =   22
               Top             =   705
               Width           =   450
            End
            Begin VB.ComboBox CboTipoPagto1 
               Height          =   315
               ItemData        =   "FrmNotasFiscais.frx":188A
               Left            =   1860
               List            =   "FrmNotasFiscais.frx":18A3
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   367
               Width           =   2655
            End
            Begin VB.TextBox TxtDocPagto1 
               Height          =   315
               Left            =   6555
               MaxLength       =   7
               TabIndex        =   21
               Top             =   360
               Width           =   1635
            End
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "FrmNotasFiscais.frx":18FD
               Height          =   330
               Left            =   2340
               TabIndex        =   23
               Top             =   705
               Width           =   6570
               _ExtentX        =   11589
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
               TabIndex        =   18
               Top             =   30
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   81657857
               CurrentDate     =   37460
            End
            Begin MSComCtl2.DTPicker DtpPagtoParcelaUnica 
               Height          =   315
               Left            =   4860
               TabIndex        =   19
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   81657857
               CurrentDate     =   37460
            End
            Begin Threed.SSCommand CmdEntMultiplas 
               Height          =   285
               Left            =   7410
               TabIndex        =   124
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
            Begin MSDataListLib.DataCombo CboFormaPagamento1 
               Bindings        =   "FrmNotasFiscais.frx":191C
               Height          =   330
               Left            =   1860
               TabIndex        =   168
               Top             =   1070
               Width           =   7050
               _ExtentX        =   12435
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "DescCombo"
               BoundColumn     =   "fopa_cd_codigo"
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Forma de Pagto:"
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
               Left            =   0
               TabIndex        =   169
               Top             =   1080
               Width           =   1785
            End
            Begin VB.Label LblCCorrente 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   116
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label LblTipoPagto 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
               Top             =   0
               Width           =   1245
            End
         End
         Begin MSMask.MaskEdBox Mskconta 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   1005
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboBancoForn1 
            Bindings        =   "FrmNotasFiscais.frx":193C
            Height          =   330
            Left            =   -73080
            TabIndex        =   24
            Top             =   1781
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
            TabIndex        =   13
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr 
            Bindings        =   "FrmNotasFiscais.frx":195D
            Height          =   330
            Left            =   2880
            TabIndex        =   14
            Top             =   600
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "cere_cd_Estruturado"
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
         Begin Threed.SSCommand CmdLimparIdentif 
            Height          =   330
            Left            =   9000
            TabIndex        =   149
            Top             =   1395
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":1972
            Picture         =   "FrmNotasFiscais.frx":1DC4
         End
         Begin Threed.SSCommand CmdLimparCR 
            Height          =   330
            Left            =   9000
            TabIndex        =   152
            Top             =   600
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":1ED6
            Picture         =   "FrmNotasFiscais.frx":2328
         End
         Begin Threed.SSCommand CmdLimparCC 
            Height          =   330
            Left            =   9000
            TabIndex        =   153
            Top             =   1005
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":243A
            Picture         =   "FrmNotasFiscais.frx":288C
         End
         Begin MSDataListLib.DataCombo CboPlanoContas 
            Bindings        =   "FrmNotasFiscais.frx":299E
            Height          =   330
            Left            =   2880
            TabIndex        =   16
            Top             =   1005
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
         Begin MSMask.MaskEdBox MskcontaCredito 
            Height          =   315
            Left            =   1920
            TabIndex        =   170
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cboPlanoContasCredito 
            Bindings        =   "FrmNotasFiscais.frx":29B5
            Height          =   330
            Left            =   2880
            TabIndex        =   171
            Top             =   1800
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
         Begin Threed.SSCommand CmdLimparIdentCcredito 
            Height          =   330
            Left            =   9000
            TabIndex        =   173
            Top             =   1800
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "FrmNotasFiscais.frx":29DB
            Picture         =   "FrmNotasFiscais.frx":2E2D
         End
         Begin VB.Label lblContaCredito 
            Alignment       =   1  'Right Justify
            Caption         =   "Conta Crédito:"
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
            TabIndex        =   172
            Top             =   1845
            Width           =   1605
         End
         Begin VB.Label LblIdentificador2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Identificador:"
            DataField       =   "c"
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
            Left            =   675
            TabIndex        =   148
            Top             =   1440
            Width           =   1155
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
            TabIndex        =   75
            Top             =   1050
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
            TabIndex        =   74
            Top             =   660
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
            TabIndex        =   73
            Top             =   1800
            Width           =   1725
         End
      End
      Begin Threed.SSPanel pandescontos 
         Height          =   1695
         Left            =   -74280
         TabIndex        =   89
         Top             =   960
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
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   1200
            Width           =   1140
         End
         Begin VB.TextBox TxtValorDesconto 
            Height          =   315
            Left            =   4800
            TabIndex        =   45
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
            Bindings        =   "FrmNotasFiscais.frx":2F3F
            Height          =   315
            Left            =   1575
            TabIndex        =   43
            Top             =   240
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "DescricaoCombo"
            BoundColumn     =   "desc_cd_desconto"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DtpVencDesconto 
            Height          =   315
            Left            =   1575
            TabIndex        =   44
            Top             =   705
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   81657857
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
            TabIndex        =   92
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
            TabIndex        =   91
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
            TabIndex        =   90
            Top             =   735
            Width           =   1125
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2535
         Left            =   -74280
         TabIndex        =   93
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
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
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Recolhimento"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "NumDarf"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(43)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(48)=   "Column(5)._ColStyle=1057040"
         Splits(0)._ColumnProps(49)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(50)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(51)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(52)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(53)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(57)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(58)=   "Column(6)._ColStyle=1057040"
         Splits(0)._ColumnProps(59)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(61)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(62)=   "Column(6).Order=7"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=16,.parent=43,.locked=-1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=44"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=45"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=47"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.locked=-1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid2 
         Height          =   3735
         Left            =   -74805
         TabIndex        =   100
         Top             =   945
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   6588
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
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Estorno"
         Columns(16).DataField=   ""
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Consolidado"
         Columns(17).DataField=   ""
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Conciliacao"
         Columns(18).DataField=   ""
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "PP"
         Columns(19).DataField=   ""
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "Cód Forma Pagto"
         Columns(20).DataField=   ""
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "Forma Pagto"
         Columns(21).DataField=   ""
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   22
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=22"
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
         Splits(0)._ColumnProps(131)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(132)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(133)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(134)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(135)=   "Column(16)._ColStyle=1048848"
         Splits(0)._ColumnProps(136)=   "Column(16).WrapText=1"
         Splits(0)._ColumnProps(137)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(138)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(139)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(140)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(141)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(142)=   "Column(17)._ColStyle=1048848"
         Splits(0)._ColumnProps(143)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(144)=   "Column(17).WrapText=1"
         Splits(0)._ColumnProps(145)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(146)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(147)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(148)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(149)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(150)=   "Column(18)._ColStyle=1048848"
         Splits(0)._ColumnProps(151)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(152)=   "Column(18).WrapText=1"
         Splits(0)._ColumnProps(153)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(154)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(155)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(156)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(157)=   "Column(19)._EditAlways=0"
         Splits(0)._ColumnProps(158)=   "Column(19)._ColStyle=1048848"
         Splits(0)._ColumnProps(159)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(160)=   "Column(19).WrapText=1"
         Splits(0)._ColumnProps(161)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(162)=   "Column(19)._MinWidth=69228880"
         Splits(0)._ColumnProps(163)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(164)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(165)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(166)=   "Column(20)._EditAlways=0"
         Splits(0)._ColumnProps(167)=   "Column(20)._ColStyle=1048848"
         Splits(0)._ColumnProps(168)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(169)=   "Column(20).WrapText=1"
         Splits(0)._ColumnProps(170)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(171)=   "Column(21).Width=2725"
         Splits(0)._ColumnProps(172)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(173)=   "Column(21)._WidthInPix=2646"
         Splits(0)._ColumnProps(174)=   "Column(21)._EditAlways=0"
         Splits(0)._ColumnProps(175)=   "Column(21)._ColStyle=1048848"
         Splits(0)._ColumnProps(176)=   "Column(21).WrapText=1"
         Splits(0)._ColumnProps(177)=   "Column(21).Order=22"
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
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(47)  =   ":id=58,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=20,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=44,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=16,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=44,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=44,.alignment=1,.locked=0"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=45"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=47,.locked=0"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(64)  =   ":id=66,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=70,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=44"
         _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=45"
         _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=47"
         _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=74,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(73)  =   ":id=74,.locked=-1"
         _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=45"
         _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=47"
         _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=78,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(78)  =   ":id=78,.locked=-1"
         _StyleDefs(79)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=44,.alignment=2"
         _StyleDefs(80)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=45"
         _StyleDefs(81)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=47"
         _StyleDefs(82)  =   "Splits(0).Columns(10).Style:id=82,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(83)  =   ":id=82,.locked=-1"
         _StyleDefs(84)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=44,.alignment=2"
         _StyleDefs(85)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=45"
         _StyleDefs(86)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=47"
         _StyleDefs(87)  =   "Splits(0).Columns(11).Style:id=86,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(88)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=44"
         _StyleDefs(89)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=45"
         _StyleDefs(90)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=47"
         _StyleDefs(91)  =   "Splits(0).Columns(12).Style:id=90,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(92)  =   ":id=90,.locked=-1"
         _StyleDefs(93)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=44,.alignment=2"
         _StyleDefs(94)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=45"
         _StyleDefs(95)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=47"
         _StyleDefs(96)  =   "Splits(0).Columns(13).Style:id=94,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(97)  =   "Splits(0).Columns(13).HeadingStyle:id=91,.parent=44"
         _StyleDefs(98)  =   "Splits(0).Columns(13).FooterStyle:id=92,.parent=45"
         _StyleDefs(99)  =   "Splits(0).Columns(13).EditorStyle:id=93,.parent=47"
         _StyleDefs(100) =   "Splits(0).Columns(14).Style:id=98,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(101) =   ":id=98,.locked=-1"
         _StyleDefs(102) =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=44,.alignment=2"
         _StyleDefs(103) =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=45"
         _StyleDefs(104) =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=47"
         _StyleDefs(105) =   "Splits(0).Columns(15).Style:id=102,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(106) =   "Splits(0).Columns(15).HeadingStyle:id=99,.parent=44"
         _StyleDefs(107) =   "Splits(0).Columns(15).FooterStyle:id=100,.parent=45"
         _StyleDefs(108) =   "Splits(0).Columns(15).EditorStyle:id=101,.parent=47"
         _StyleDefs(109) =   "Splits(0).Columns(16).Style:id=24,.parent=43"
         _StyleDefs(110) =   "Splits(0).Columns(16).HeadingStyle:id=21,.parent=44"
         _StyleDefs(111) =   "Splits(0).Columns(16).FooterStyle:id=22,.parent=45"
         _StyleDefs(112) =   "Splits(0).Columns(16).EditorStyle:id=23,.parent=47"
         _StyleDefs(113) =   "Splits(0).Columns(17).Style:id=106,.parent=43"
         _StyleDefs(114) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=44"
         _StyleDefs(115) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=45"
         _StyleDefs(116) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=47"
         _StyleDefs(117) =   "Splits(0).Columns(18).Style:id=110,.parent=43"
         _StyleDefs(118) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=44"
         _StyleDefs(119) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=45"
         _StyleDefs(120) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=47"
         _StyleDefs(121) =   "Splits(0).Columns(19).Style:id=114,.parent=43"
         _StyleDefs(122) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=44"
         _StyleDefs(123) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=45"
         _StyleDefs(124) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=47"
         _StyleDefs(125) =   "Splits(0).Columns(20).Style:id=118,.parent=43"
         _StyleDefs(126) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=44"
         _StyleDefs(127) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=45"
         _StyleDefs(128) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=47"
         _StyleDefs(129) =   "Splits(0).Columns(21).Style:id=122,.parent=43"
         _StyleDefs(130) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=44"
         _StyleDefs(131) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=45"
         _StyleDefs(132) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=47"
         _StyleDefs(133) =   "Named:id=33:Normal"
         _StyleDefs(134) =   ":id=33,.parent=0"
         _StyleDefs(135) =   "Named:id=34:Heading"
         _StyleDefs(136) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(137) =   ":id=34,.wraptext=-1"
         _StyleDefs(138) =   "Named:id=35:Footing"
         _StyleDefs(139) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(140) =   "Named:id=36:Selected"
         _StyleDefs(141) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(142) =   "Named:id=37:Caption"
         _StyleDefs(143) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(144) =   "Named:id=38:HighlightRow"
         _StyleDefs(145) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(146) =   "Named:id=39:EvenRow"
         _StyleDefs(147) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(148) =   "Named:id=40:OddRow"
         _StyleDefs(149) =   ":id=40,.parent=33"
         _StyleDefs(150) =   "Named:id=41:RecordSelector"
         _StyleDefs(151) =   ":id=41,.parent=34"
         _StyleDefs(152) =   "Named:id=42:FilterBar"
         _StyleDefs(153) =   ":id=42,.parent=33"
      End
      Begin Threed.SSPanel PnlDados1 
         Height          =   360
         Left            =   -74280
         TabIndex        =   105
         Top             =   480
         Width           =   5880
         _Version        =   65536
         _ExtentX        =   10372
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados1_1 
         Height          =   360
         Left            =   -68280
         TabIndex        =   106
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados2 
         Height          =   360
         Left            =   -74280
         TabIndex        =   107
         Top             =   480
         Width           =   5880
         _Version        =   65536
         _ExtentX        =   10372
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados2_2 
         Height          =   360
         Left            =   -68280
         TabIndex        =   108
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados3 
         Height          =   360
         Left            =   -74400
         TabIndex        =   109
         Top             =   480
         Width           =   6120
         _Version        =   65536
         _ExtentX        =   10795
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSPanel PnlDados3_3 
         Height          =   360
         Left            =   -68160
         TabIndex        =   110
         Top             =   480
         Width           =   2280
         _Version        =   65536
         _ExtentX        =   4022
         _ExtentY        =   635
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
         Alignment       =   1
      End
      Begin Threed.SSCommand CmdDuplMultiplas 
         Height          =   330
         Left            =   -70860
         TabIndex        =   136
         Top             =   4785
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
         TabIndex        =   137
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
            TabIndex        =   139
            Top             =   1335
            Width           =   780
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00000000&
            Caption         =   "&Confirmar"
            Height          =   240
            Left            =   1080
            TabIndex        =   138
            Top             =   1335
            Width           =   780
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   1215
            Left            =   90
            TabIndex        =   140
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
               TabIndex        =   144
               Top             =   915
               Width           =   1725
            End
            Begin VB.OptionButton OptProduto 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produtos"
               Height          =   255
               Left            =   210
               TabIndex        =   143
               Top             =   630
               Width           =   1725
            End
            Begin VB.OptionButton OptFisica 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pessoa Física"
               Height          =   255
               Left            =   210
               TabIndex        =   142
               Top             =   330
               Width           =   1725
            End
            Begin VB.OptionButton OptJuridica 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pessoa Jurídica"
               Height          =   255
               Left            =   210
               TabIndex        =   141
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
         TabIndex        =   103
         Top             =   4365
         Width           =   1170
      End
      Begin VB.Label LblValorTotal 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67965
         TabIndex        =   102
         Top             =   4800
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
         TabIndex        =   101
         Top             =   4845
         Width           =   1170
      End
      Begin VB.Label LblValorTotalDistrib 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -69000
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
         Top             =   4365
         Width           =   1410
      End
      Begin VB.Label LblValorNota 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72600
         TabIndex        =   82
         Top             =   4320
         Width           =   1485
      End
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   3840
      TabIndex        =   58
      Top             =   6720
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
      Left            =   5520
      TabIndex        =   59
      Top             =   6720
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
      Left            =   2640
      Top             =   6960
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
      Left            =   7560
      Top             =   7320
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
      Left            =   5640
      Top             =   7320
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
      Left            =   5160
      Top             =   6960
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
      Left            =   2280
      TabIndex        =   135
      Top             =   6720
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
   Begin MSAdodcLib.Adodc DatIdentificador 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Caption         =   "DatIdentificador"
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
   Begin Threed.SSPanel PanEntMultiplas 
      Height          =   1590
      Left            =   2280
      TabIndex        =   125
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
         TabIndex        =   130
         Top             =   1260
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox TxtNParcelas 
         Height          =   315
         Left            =   1545
         TabIndex        =   126
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdSairEntMult 
         BackColor       =   &H00000000&
         Caption         =   "&Retornar"
         Height          =   240
         Left            =   2160
         TabIndex        =   132
         Top             =   1260
         Width           =   840
      End
      Begin VB.CommandButton CmdConfEntMult 
         BackColor       =   &H00000000&
         Caption         =   "&Confirmar"
         Height          =   240
         Left            =   1275
         TabIndex        =   131
         Top             =   1260
         Width           =   840
      End
      Begin VB.TextBox TxtPeriodicidade 
         Height          =   315
         Left            =   1545
         TabIndex        =   127
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
         TabIndex        =   128
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
         TabIndex        =   129
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
         TabIndex        =   134
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
         TabIndex        =   133
         Top             =   735
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc DatFormaPagamento 
      Height          =   330
      Left            =   2640
      Top             =   7320
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
End
Attribute VB_Name = "FrmNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Reload As Boolean

Dim ValorTotalDuplicata As Double       'Acumula o total de duplicatas até compor o valor total da Nota Fiscal
Dim ValorTotalDistribuicao As Double    'Acumula o total de distribuições até compor o valor total da Nota Fiscal
Dim Xtipomov As String
Dim Xdebcred As String
Dim Xbancoforn As Integer
Dim XConta As Integer
Dim xcont As Integer                    'Controle na ditribuicao
Dim XContDesc As Integer                'Controle em descontos
Dim XDistribuido As String              'verifica se o valor total foi distribuido
Dim XPcr As String                      'Armazena o pcr digitado na distribuicao
Dim xcc As String                       'Armazena a conta contábil digitada na distribuicao
Dim xcod As String                      'Armazena o desconto escolhido na alteracao do desconto
Dim xNumFat As Long
Dim Tipo As String                      'Verifica o tipo do documento
Dim ParcelasMultiplas As Boolean        'Verifica se foi gerada parcelas multiplas
Dim XTipoFatura As String               '1=(1 nota 1 dupl) 2=(1 nota 1 ou + dupl) 3=(2 ou + notas p/ 1 ou + dupl)
Dim XValorBruto As Double               'Armazena o valor bruto da nota
Dim Status As String                    'Armazena o Status da Nota Fiscal
Dim XFT_STATUSANT As String             'Armazena o Status anterior da Nota Fiscal (alteração de NF)
Dim XVlDuplicata As Double              'Valor inicial da duplicata que está sendo alterada
Dim NumDocAnt As String                 'armazena o nº do documento no caso de uma alteração
Dim XFB_PROVISIONA As Boolean           'Indica se a NF deve ou não ser provisionada
Dim XFB_ESTOQUE As Boolean              'Indica se a nota é do estoque
Dim XFB_QUALIFAT As Boolean             'Indica se a nota é do Qualifat
Dim XFB_QUALICAPI As Boolean            'Indica se a nota é do QualiCapi - 16/10/09

Dim XLT_DOCPAGRECVERSO As String        'guarda o documento de pagamento para exclusão do verso de cheque
Dim XLT_CONTAVERSO As Integer           'guarda o número da conta corrente para exclusão do verso de cheque

'Variável que armazena as alterações realizadas na nota fiscal
Dim XFT_DESCRICAO As String
  

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor1 As New XArray
Dim Vetor2 As New XArray
Dim Vetor3 As New XArray
Dim Vetor4 As New XArray                'Descontos/Acrescimos de duplicata
Dim VetAnt As New XArray                'Armazena os valores do vetor antes das alterações
Dim EntrouDesc As Boolean               'Controla a entrada na orelha de desconto
Dim EntrouDupl As Boolean               'Controla a entrada na orelha de duplicata
Dim EntrouDist As Boolean               'Controla a entrada na orelha de distribuicao
Dim XDistribuicao As String             'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDuplicata As String                'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDesconto As String                 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XFT_INSPARCMULTIPLA As String
Dim XFT_INSDUPLMULTIPLA As String
Dim XFT_NOTADUPLICATA As String         'Informar se é Nota ou Duplicata
Dim XCLASSE_CENTROCUSTO As String * 1   'Declaração em 05/06/08 - Patrícia
Dim XCLASSE_CCONTABIL As String * 1     'Declaração em 05/06/08 - PatríciA
Dim Data_EmpObr As String               'Declaração em 09/06/08 - Patrícia
Dim XMENSAGEM As String                 'Declaração em 10/06/08 - Patrícia
Dim XLT_CONSOLIDADO As String           'Armazena se a duplicata realizada é consolidada - Declaração em 09/07/09 - Larissa


'Variáveis usadas na geração dos logs das operações realizadas - Larissa 25/06/2008
Dim XFV_VETORLOG() As Variant
Dim XFT_LOG As String
Dim XFT_CONTADORLOG As Integer
Dim XFM_MATRIZNOTAFISCAL As Variant
Dim XFM_MATRIZPAGAMENTO As Variant
Dim XFM_MATRIZAPROPRIACAO As Variant
Dim XLT_VALORDIFFNOTA As Double


'Variável utilizada na alteração da duplicata
Dim PosCursor As Integer 'Guarda o Bookmark do grid

Dim ERRLOOP  As ADODB.Error
Dim STRERROR As String
Dim XFT_NOVO As Byte '13/11/09

'Variáveis utilizadas para controle de Contas associadas a CRs (CboPcr/2 e CboPlanoContas/2) - 21/03/2022 - Paulo Garcia
Dim XFT_CR  As String
Dim XFT_CT As String

Function FunVerificaClassEmprObr(CentroCusto As String, CContabil As String, Empresa As Integer) As String
'Criação da Função em 09/06/08 - Patrícia
'Função que realiza a verificação da classificação "Empreendimento\Obra" dos Centros de Custos e Contas Contábeis
'Só serão aceitos CRs e Contas Contábeis com a mesma Classificação "Empreendimento\Obra" ou Classificação Nula.

    Dim ResEmpObr As Object
    Dim Sql As String
    Dim Desc_CREmpObr As String
    Dim Desc_CCEmpObr As String
    Dim Mensagen As String

    '******** Início Rotina de Classificação "Empreendimento\Obra"
    Mensagen = ""
    If (DtpEmissao) >= CDate(Data_EmpObr) Then
                    Sql = "SELECT * FROM CentroResultados WHERE cere_cd_Pcr= " & CentroCusto & " and empr_cd_Empresa= " & Str(Empresa)
                    SubQOpenRecordset ResEmpObr, Sql, "Estatico"
                    If Not (ResEmpObr.BOF And ResEmpObr.EOF) Then
                        XCLASSE_CENTROCUSTO = RTrim(LTrim(FunNulo(ResEmpObr!cere_tx_EmpObr)))
                        If XCLASSE_CENTROCUSTO = "E" Then
                                    Desc_CREmpObr = "Empreendimento"
                        ElseIf XCLASSE_CENTROCUSTO = "O" Then
                                    Desc_CREmpObr = "Obra"
                       End If
                    End If
                    ResEmpObr.Close
                    Sql = "SELECT * FROM PlanoContas WHERE plco_cd_CodReduzido= '" & CContabil & "'" & " and empr_cd_Empresa= " & Str(Empresa)
                    SubQOpenRecordset ResEmpObr, Sql, "Estatico"
                    If Not (ResEmpObr.BOF And ResEmpObr.EOF) Then
                            If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "E" Then
                                Desc_CCEmpObr = "Empreendimento"
                            ElseIf Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "O" Then
                                Desc_CCEmpObr = "Obra"
                                Else
                                    Desc_CCEmpObr = "Sem Classificação"
                            End If
            
                            If XCLASSE_CENTROCUSTO = "E" Or XCLASSE_CENTROCUSTO = "O" Then 'Se o CR estiver classificado, só poderá informar conta com a mesma classificação ou classificação = Nullo
                                If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "E" Or Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "O" Then
                                    If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) <> XCLASSE_CENTROCUSTO Then
                                        Mensagen = "O Centro de Custo é do tipo " & Desc_CREmpObr & " e a Conta Contábil é do tipo " & Desc_CCEmpObr & "!"
                                    End If
                                End If
                            End If
                    End If
                    ResEmpObr.Close
    End If
'******** Fim Rotina de Classificação "Empreendimento\Obra"
  FunVerificaClassEmprObr = Mensagen
End Function

Public Property Get CARREGAGRID()
    CARREGAGRID = Reload
End Property

Function FunVerificaDuplicataCancelada() As Boolean
    
    Do While Not TDBGrid2.EOF
        If Me.TDBGrid2.Columns("Status").Text = "Cancelado" Then
            FunVerificaDuplicataCancelada = False
            Exit Function
        End If
        TDBGrid2.MoveNext
    Loop
    
    FunVerificaDuplicataCancelada = True
    
End Function

Sub subGravaVetorDupl(Tabela As String, CamposTabela As String, Tipos As Variant, QtdCampos As Byte, Colunas As Variant, Complemento As String, Vetor As XArray, Grid As TDBGrid)
    Dim Valores As String
    Dim x As Integer, y As Integer
    Dim XLB_ESTORNO As Boolean
    
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
        
        For x = 0 To Vetor.UpperBound(1)
            If Vetor(x, 16) <> "S" Then
                Valores = ""
                For y = 0 To QtdCampos - 1
                    If Valores <> "" Then
                        If Tipos(y) = "N" Then
                            'Valores = Valores & ", " & Str(Vetor(x, Colunas(y)))
                            'Qdo o valor é "" ocorre Type Mismatch com o uso da função Str
                            Valores = Valores & ", " & FunNuloBancoVal(Vetor(x, Colunas(y)))
                        ElseIf Tipos(y) = "T" Then
                            Valores = Valores & ", '" & Vetor(x, Colunas(y)) & "'"
                        ElseIf Tipos(y) = "D" Then
                            Valores = Valores & "," & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                        ElseIf Tipos(y) = "TN" Then
                            Valores = Valores & ", Null"
                        End If
                    Else
                        If Tipos(y) = "N" Then
                            'Valores = Valores & ", " & Str(Vetor(x, Colunas(y)))
                            'Qdo o valor é "" ocorre Type Mismatch com o uso da função Str
                            Valores = FunNuloBancoVal(Vetor(x, Colunas(y)))
                        ElseIf Tipos(y) = "T" Then
                            Valores = "'" & Vetor(x, Colunas(y)) & "'"
                        ElseIf Tipos(y) = "D" Then
                            Valores = "" & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                        ElseIf Tipos(y) = "TN" Then
                            Valores = Valores & ", Null"
                        End If
                    End If
                Next
                Conexao.Execute "INSERT INTO " & Tabela & " (" & CamposTabela & ") VALUES (" & Valores & Complemento & ")"
            End If
        Next
    End If
End Sub

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
        If XTipoFatura <> "4" Then 'não possui fatura
            If Not XEntrouDupl Then
                XEntrouDupl = True
            End If
            If Not XEntrouDesc Then
                XEntrouDesc = True
            End If
        End If
    End If
    
    If XEntrouDesc And XEntrouDupl Then 'Entrou na orelha de descontos e em duplicata
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
    ElseIf XEntrouDesc Then 'Só entrou na orelha de descontos
        If Vetor1(0, 0) <> Empty Then
            SomaDesc = funSomaColuna(Vetor1, 2)
            If Format(TxtValor.Text, "standard") <> Format(SomaDesc, "standard") Then
                MsgBox "A soma dos Descontos é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                TabNotasFiscais.Tab = 1
                VerificaValores = True
                Exit Function
            End If
        End If
    ElseIf XEntrouDupl Then 'só entrou na orelha de duplicatas
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
    ElseIf Not XEntrouDupl And ((CboStatus.ListIndex = "2") Or (CboStatus.ListIndex = "1")) And Tipo = "0" Then 'Não entrou na orelha de duplicata é realizada e não compoe duplicata automat.
        MsgBox "A composição da Duplicata é obrigatória para Notas Fiscais com status 'Realizado ou Compromissado'.", vbCritical, "ATENÇÃO"
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
    For x = 0 To Grid.Columns.Count - 6
        subSetaVal Vetor, Cont, x, FunNulo(VetorCampos(x))
    Next
    With Grid
        .ReBind
    End With
End Sub

Function ExisteDistribuicao(Numchaves As Byte, coluna As Variant, VALOR As Variant, Vetor As XArray, Grid As TDBGrid) As Boolean
    Dim Cont As Integer, Achou As Boolean
    Dim x As Integer, y As Integer
    
    'Verifica se o registro está cadastrado no vetor
    
    ExisteDistribuicao = False
    
    While x <= Vetor.UpperBound(1)
        If Vetor(x, coluna(0)) = CStr(VALOR(0)) Then
            If Numchaves = 1 Then
                ExisteDistribuicao = True
                x = Vetor.UpperBound(1)
            Else
                Achou = True
                y = 1
                While Achou And y <= Numchaves - 1
                    'If coluna(Y) <> Valor(Y) Then
                    If Vetor(x, coluna(y)) <> CStr(VALOR(y)) Then
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
        DtpEntrada.Enabled = True
        
        If XInserir = "I" Then
            CmdEntMultiplas.Enabled = True
            DtpEntrada.Value = Format(Now, "dd/MM/yy")
            DtpEmissao.Value = Format(Now, "dd/MM/yy")
        End If
        
        DtpEmissao.Enabled = False
        TxtValor.Enabled = True
        CboHistorico.Enabled = True
        TxtHistorico.Enabled = True
        CmdLimparHistorico.Enabled = True
        LblValorLiquido1.Enabled = True
'        TxtChamado.Enabled = True

        
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
        TabNotasFiscais.TabEnabled(3) = False
        TabAuxiliar.Enabled = True
        CboPcr.Enabled = True
        
        MskConta.Enabled = True
        CboPlanoContas.Enabled = True
        TabAuxiliar.TabEnabled(0) = True
        TabAuxiliar.TabEnabled(1) = True
        If DtpPagtoParcelaUnica.Enabled Then DtpPagtoParcelaUnica.Value = ""
        DtpPagtoParcelaUnica.Enabled = False
        If TxtDocPagto1.Enabled Then TxtDocPagto1.Text = ""
        TxtDocPagto1.Enabled = False
        ChkDistribuicao.Value = 0
        ChkDistribuicao.Enabled = False

    ElseIf CboStatus.ListIndex > 0 Then
        CmdEntMultiplas.Enabled = False
        PanNotasFiscais.Enabled = True
        CboFornecedor.Enabled = True
        DtpEntrada.Enabled = True
        
        '/////////CONTABILIDADE///////////
        'Checa se foi provisionado e não deixa mudar se o mês estiver fechado
        If XGT_TIPOEXPORTACAO = "PC" Then
            Dim XLO_CHECK As Object
            Dim XLT_SQL As String
        
            XLT_SQL = "SELECT lancamentos.lanc_cd_lanc FROM lancamentos LEFT JOIN notasfiscais ON lancamentos.lanc_cd_origemfin = notasfiscais.nofi_cd_notafiscal " & _
                        "WHERE   (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T')  AND notasfiscais.nofi_cd_notafiscal = " & Chave & ""
            SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
            If Not XLO_CHECK.EOF Then ' foi provisionado
                If funMesFechado(DtpEntrada.Value) <> "A" Then
                    DtpEntrada.Enabled = False
                End If
            End If
            
            XLO_CHECK.Close
            
            Set XLO_CHECK = Nothing
        End If
        '////////////////////////////////
        
        DtpEmissao.Enabled = True
        TxtValor.Enabled = True
        CboHistorico.Enabled = True
        TxtHistorico.Enabled = True
        CmdLimparHistorico.Enabled = True
        LblValorLiquido1.Enabled = True
'        TxtChamado.Enabled = True
        
        If XInserir = "I" Then
            ChkDistribuicao.Value = 0
        End If
        ChkDistribuicao.Enabled = True

        If ChkDistribuicao.Value = "1" Then
            TabAuxiliar.TabEnabled(0) = False
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            MskConta.Enabled = False
            CboPlanoContas.Enabled = False
        Else
            TabAuxiliar.TabEnabled(0) = True
            MskPcr.Enabled = True
            CboPcr.Enabled = True
            MskConta.Enabled = True
            CboPlanoContas.Enabled = True
        End If

        If CboStatus.ListIndex = 1 Then 'Compromissado
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
                TabNotasFiscais.TabEnabled(2) = True
            End If
        ElseIf CboStatus.ListIndex = 2 Or CboStatus.ListIndex = 3 Then 'Realizado ou Cancelado - Alterado em 10/03/09 - Larissa
            If Tipo = "1" Then  'gera duplicata automaticamente
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = True
                TabNotasFiscais.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(2) = False
                If XLT_CONSOLIDADO <> "S" Then
                    DtpPagtoParcelaUnica.Enabled = True
                    TxtDocPagto1.Enabled = True
                Else
                    DtpPagtoParcelaUnica.Enabled = False
                    TxtDocPagto1.Enabled = False
                    CboTipoPagto1.Enabled = False
                    CboCCorrente.Enabled = False
                End If
            Else     ' Não gera duplicata automaticamente
                DtpPagamento.Enabled = True
                TxtDocPagto2.Enabled = True
                TabAuxiliar.Enabled = True
                TabAuxiliar.TabEnabled(1) = False
                TabNotasFiscais.TabEnabled(1) = True
                TabNotasFiscais.TabEnabled(2) = True
            End If
        End If
    End If
    If XTipoFatura = "3" Then
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
    End If
    
    'Cássio 21/06/2013
    'Se o centro de custo for inativo desabilita o combo
    If Not DatPCR.Recordset.EOF Then
        If DatPCR.Recordset.Fields("cere_tx_status") = "I" And CboStatus.ListIndex = 2 Then
                CboPcr.Enabled = False
                MskPcr.Enabled = False
        Else
                CboPcr.Enabled = True
                MskPcr.Enabled = True
        End If
    End If
End Sub

Sub DesabilitaRealizado(XLT_DOCUMENTO As String, XLB_TRUEFALSE As Boolean)
    If XLT_DOCUMENTO = "Nota" Or XLT_DOCUMENTO = "Estoque" Then
        CmdGravar.Enabled = False
        PanNotasFiscais.Enabled = True
        CboStatus.Enabled = False
        CboTipoDoc.Enabled = False
        TxtCodTipoDoc.Enabled = False
        CboFornecedor.Enabled = False
        MskNumdocumento.Enabled = False
        TxtSerie.Enabled = False
        DtpEntrada.Enabled = False
        DtpEmissao.Enabled = False
        TxtValor.Enabled = False
        CboHistorico.Enabled = False
        TxtHistorico.Enabled = False
        CmdLimparHistorico.Enabled = False
        TabAuxiliar.Enabled = True
'        TxtChamado.Enabled = True
        
        
        If XLT_DOCUMENTO = "Nota" Then
            ChkDistribuicao.Enabled = True
            
            If ChkDistribuicao.Value = "1" Then
                TabAuxiliar.TabEnabled(0) = False
                MskPcr.Enabled = False
                CboPcr.Enabled = False
                MskConta.Enabled = False
                CboPlanoContas.Enabled = False
            Else
                TabAuxiliar.TabEnabled(0) = True
                MskPcr.Enabled = True
                CboPcr.Enabled = True
                MskConta.Enabled = True
                CboPlanoContas.Enabled = True
            End If
        Else
            ChkDistribuicao.Enabled = False
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            MskConta.Enabled = False
            CboPlanoContas.Enabled = False
                
            If ChkDistribuicao.Value = "1" Then
                TabAuxiliar.TabEnabled(0) = False
                CmdInsereDistribuicao.Enabled = False
                CmdRemoverDistribuicao.Enabled = False
                MskPcr2.Enabled = False
                CboPcr2.Enabled = False
                MskConta2.Enabled = False
                CboPlanoContas2.Enabled = False
                TxtPercDistribuicao.Enabled = False
                TxtVldistribuicao.Enabled = False
                CboIdentificador2.Enabled = False
            Else
                TabAuxiliar.TabEnabled(0) = True
                CboIdentificador.Enabled = True
            End If
        End If
           
        If Tipo = "1" Then  'gera duplicata automaticamente
            TabAuxiliar.TabEnabled(1) = True
            DtpVencParcelaUnica.Enabled = False
            DtpPagtoParcelaUnica.Enabled = False
            CboTipoPagto1.Enabled = False
            TxtDocPagto1.Enabled = False
            CboCCorrente.Enabled = False
            TxtCCorrente.Enabled = False
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
        TxtCCorrente2.Enabled = XLB_TRUEFALSE
        CboBancoForn2.Enabled = XLB_TRUEFALSE
        DtpVencimento.Enabled = XLB_TRUEFALSE
        CmdConfirmaduplicata.Enabled = XLB_TRUEFALSE
        If TDBGrid2.Columns("CodStatus") = "R" Then
            DtpPagamento.Enabled = XLB_TRUEFALSE
            TxtDocPagto2.Enabled = XLB_TRUEFALSE
        End If
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
    Dim RESDUPL As Object
    Dim ResNota As Object
    Dim x As Integer
    Dim ResCentroResultado As Object
    Dim ResPlano As Object
    
    'Tab 0
    TabNotasFiscais.TabEnabled(1) = True
    MskNumdocumento.Enabled = True
    SubQOpenRecordset ResNota, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal =" & Chave, "Estatico"
    If Not (ResNota.BOF And ResNota.EOF) Then
        If Not IsNull(ResNota!nofi_tx_pendestoque) Then
            XFB_ESTOQUE = True 'NOTA DO ESTOQUE
        Else
            XFB_ESTOQUE = False
        End If
        
        If Not IsNull(ResNota!nofi_tx_OrigemQualiFat) Then
            XFB_QUALIFAT = True 'NOTA DO QUALIFAT
        Else
            XFB_QUALIFAT = False
        End If
        
        If IsNull(ResNota!nofi_nr_TituloCapi) Or ResNota!nofi_nr_TituloCapi = "" Then
            XFB_QUALICAPI = False
        Else
            XFB_QUALICAPI = True 'NOTA DO QUALICAPI
        End If
        
        XTipoFatura = ResNota!nofi_tx_tipofatura
        XValorBruto = ResNota!nofi_vl_valor
        XFT_STATUSANT = ResNota!nofi_tx_status
        Status = ResNota!nofi_tx_status
        
        If Status = "P" Then 'Previsto
            Status = "0"
        ElseIf Status = "C" Then 'Compromissado
            Status = "1"
        ElseIf Status = "R" Then 'Realizado
            Status = "2"
        'Inserido em 10/03/09 - Larissa
        ElseIf Status = "X" Then 'Cancelado
            Status = "3"
        End If
        
        If (FunNulo(ResNota!nofi_tx_pendestoque) = "P") Then
            chkPendente.Value = 1
        Else
            chkPendente.Value = 0
        End If
        
        'Inserido por Jeferson em 14/06/21
        If (EntSaida = "S") And (PCliente = "QUALIDADOS") Then
            If (FunNulo(ResNota!plco_cd_contacredito) <> "") Then
                SubQOpenRecordset ResPlano, "SELECT * FROM PlanoContas WHERE plco_cd_Conta =" & ResNota!plco_cd_contacredito, "Estatico"
                MskcontaCredito.Text = ResPlano!plco_cd_CodReduzido
                cboPlanoContasCredito.BoundText = MskcontaCredito.Text
            End If
        End If

        
        CboStatus.ListIndex = Status
        CboTipoDoc.BoundText = FunNulo(ResNota!tido_cd_TipoDoc)
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        
        If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "TB" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "PC" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "DEV" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "REST" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "RA" Or DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "NP" Then
            CboTipoDoc.Enabled = False
            TxtCodTipoDoc.Enabled = False
        End If
        
        If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "DEV" Then
            subConectarControleDadosNV DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and (empr_cd_empresa = " & PCodEmpresa & " or empr_cd_empresa is null) ORDER BY focl_tx_razaosocial ", Estatico
        End If
        
        If Not IsNull(ResNota!focl_cd_forncli) Then
            CboFornecedor.BoundText = ResNota!focl_cd_forncli
        End If
        
        If EntSaida = "S" Then
            If Not IsNull(ResNota!focl_cd_forncli) Then
                subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
            End If
            TDBGrid2.Columns(3).Caption = "Devolução"
            TDBGrid3.Columns(5).Caption = "Devolução"
            'Verifica se existe Devolução
            If Not IsNull(ResNota!nofi_cd_notaGlosaDev) Then
                LblPrazo.Visible = False
                TxtPrazo.Visible = False
            
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
                TDBGrid2.Columns(3).Visible = True
                
                TDBGrid3.Columns("Valor Líquido").Visible = True
                TDBGrid3.Columns(5).Visible = True
            End If
        Else
            TDBGrid2.Columns(3).Caption = "Glosa"
            TDBGrid3.Columns(5).Caption = "Glosa"
            'Verifica se existe Glosa
            If Not IsNull(ResNota!nofi_cd_notaGlosaDev) Then
                LblPrazo.Visible = False
                TxtPrazo.Visible = False
                
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
                TDBGrid2.Columns(3).Visible = True
                
                TDBGrid3.Columns("Valor Líquido").Visible = True
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
        
'        If Not IsNull(ResNota!nofi_nr_chamado) Then
'            TxtChamado.Text = ResNota!nofi_nr_chamado
'        End If
        
    End If
'    chkSigiloso.Value = IIf(Not IsNull(ResNota!nofi_bl_Sigiloso) And ResNota!nofi_bl_Sigiloso = True, 1, 0)
    ResNota.Close
    
    Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
    subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_vl_GlosaDevolucao", "Valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata", "dupl_tx_estorno", "dupl_tx_Consolidado", "dupl_dt_concilia", "dupl_dt_ProcPagamento", "fopa_cd_codigo", "fopa_tx_descricao"), Vetor2, TDBGrid2
    TxtDevolucao.Text = Format(funSomaColuna(Vetor2, 3), "standard")
    
    'Desabilitei o botão gravar quando a duplicata era consolidado
    'Cássio 31/07/2012
'    Dim XLI_I As Integer
'    Dim XLD_CONSOLIDADO As Boolean
'    If Status = "2" Then
'        For XLI_I = 0 To Vetor2.UpperBound(1)
'            If Vetor2(XLI_I, 17) = "S" Then
'                XLD_CONSOLIDADO = True
'            End If
'        Next
'        CmdGravar.Enabled = Not XLD_CONSOLIDADO
'    End If
         
    Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
    subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura", "deen_dt_recolhimento", "deen_nr_DarfDam"), Vetor1, TDBGrid1
    
    Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " ORDER BY cere_cd_estruturado"
    subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "dist_vl_GlosaDevolucao", "dist_vl_valor", "percentual", "iden_cd_Projeto", "iden_tx_Descricao"), Vetor3, TDBGrid3
'
'    'Cássio Medeiros 25/06/2013
'    Sql = "SELECT CentroResultados.cere_cd_Pcr FROM Distribuicao INNER JOIN CentroResultados ON CentroResultados.cere_cd_Pcr = Distribuicao.cere_cd_Pcr" & _
'          " LEFT JOIN NotasFiscais ON NotasFiscais.nofi_cd_notafiscal = Distribuicao.nofi_cd_notafiscal" & _
'          " WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " AND CentroResultados.cere_tx_status = 'I'"
'
'    'Se na distribuição existir centro de resultado inativo, desabilita os botões excluir e alterar.
'    SubQOpenRecordset ResCentroResultado, Sql, "Estatico"
'    If Not (ResCentroResultado.BOF And ResCentroResultado.EOF) Then
'        CmdAlteraDistribuicao.Enabled = False
'        CmdRemoverDistribuicao.Enabled = False
'    End If
        
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
        
            If (Len(Trim(ResDist!cere_cd_estruturado)) = Len(MskPcr.Mask)) Then
                MskPcr.Text = ResDist!cere_cd_estruturado
            Else
                MskPcr.Mask = pPCRmascara
                MskPcr.Text = ""
                CboPcr.BoundText = 0
                CboPcr.Text = ""
            End If
            
           MskConta.Text = ResDist!plco_cd_CodReduzido
         
           CboPcr.BoundText = MskPcr.Text 'Inserido em 09/06/08 - Patrícia
           Verifica_DataEmissao 'Inserido em 31/07/08 - Larissa
           
           CboPlanoContas.BoundText = MskConta.Text
           CboIdentificador2.BoundText = FunNulo(ResDist!iden_cd_Projeto)
           
           XFT_CR = MskPcr.Text             ' armazena o codigo CR
           XFT_CT = MskConta.Text           ' armazena o codigo Conta Contábil
           
        End If
        
        ResDist.Close
    End If
    
    SubQOpenRecordset ResDoc, "SELECT tipodocumentos.tido_cd_tipodoc,tipodocumentos.tido_tx_duplicata FROM TipoDocumentos where TipoDocumentos.tido_cd_tipodoc = " & CboTipoDoc.BoundText, "Estatico"
    If Not ResDoc.BOF And Not ResDoc.BOF Then
       Tipo = ResDoc!tido_tx_duplicata
       ResDoc.Close
    End If

    If Status = "0" Or Tipo = "1" Then 'É previsto ou gera duplicata automaticamente
        
        '26/09/08 - Larissa: inseri a condição (dupl_tx_Estorno IS NULL OR dupl_tx_Estorno = 'N') para tratar a existência de estornos associados à nota.
        SubQOpenRecordset RESDUPL, "SELECT Duplicatas.*,ContasCorrente.coco_tx_conta,ContasCorrente.coco_cd_agencia FROM ((Duplicatas LEFT JOIN BancoFornecedor ON Duplicatas.bafo_cd_codigo = BancoFornecedor.bafo_cd_codigo) LEFT JOIN ContasCorrente ON Duplicatas.coco_cd_codigo = ContasCorrente.coco_cd_codigo) where (dupl_tx_Estorno IS NULL OR dupl_tx_Estorno = 'N') AND dupl_nr_fatura =" & NFatura, "Estatico"
        
        If Not (RESDUPL.BOF And RESDUPL.EOF) Then
            DtpVencParcelaUnica.Value = Format$(RESDUPL!dupl_dt_vencimento, "DD/MM/YY")
             
            If IsNull(RESDUPL!dupl_tx_consolidado) Then
                XLT_CONSOLIDADO = "N"
            Else
                XLT_CONSOLIDADO = RESDUPL!dupl_tx_consolidado
            End If
             
            If Status = "2" Then
               DtpPagtoParcelaUnica.Value = Format$(RESDUPL!dupl_dt_pagrec, "DD/MM/YY")
            End If
            
            If Not IsNull(RESDUPL!dupl_tx_tipopagrec) Then
               For x = 0 To CboTipoPagto1.ListCount - 1
                   If CboTipoPagto1.ItemData(x) = CInt(RESDUPL!dupl_tx_tipopagrec) Then
                       CboTipoPagto1.ListIndex = x
                   End If
               Next
            End If
             
            TxtDocPagto1.Text = FunNulo(RESDUPL!dupl_tx_docpagrec)
            
            If Not IsNull(RESDUPL!coco_cd_codigo) Then
               CboCCorrente.BoundText = RESDUPL!coco_cd_codigo
               TxtCCorrente.Text = RESDUPL!coco_cd_codigo
            End If
            
            If Not IsNull(RESDUPL!fopa_cd_codigo) Then
                CboFormaPagamento1.BoundText = RESDUPL!fopa_cd_codigo
            End If
            
            If Not IsNull(RESDUPL!bafo_cd_codigo) Then
               CboBancoForn1.BoundText = RESDUPL!bafo_cd_codigo
            End If
            
            
       End If
       RESDUPL.Close
    End If
    
    If Status = "2" Then
        CboStatus.Enabled = False
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
            If FunObrigatorioMSK(MskConta, "Conta Contábil é Obrigatória") Then Exit Function
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
            If CboTipoPagto2.ListIndex = 2 Then
                If FunObrigatorioCBO(CboBancoForn2, "O preenchimento do Campo Banco do Fornecedor é obrigatório.") Then Exit Function
            End If
        End If
    ElseIf TabNotasFiscais.Tab = 3 Then
        If FunObrigatorioTXT(TxtVldistribuicao, "O preenchimento do campo Valor é obrigatório.") Then Exit Function
        If FunObrigatorioCBO(CboPcr2, "O preenchimento do campo Centro Custo é obrigatório.") Then Exit Function
        If FunObrigatorioMSK(MskConta2, "O preenchimento do campo Conta Contábil é obrigatório.") Then Exit Function
    End If
    
    ChecarCampos = False
End Function

Sub PrepararInsercao()
    If TabNotasFiscais.Tab = 0 Then
        CboTipoDoc.Text = ""
        TxtCodTipoDoc.Text = ""
        CboFornecedor.Text = ""
        MskNumdocumento.Text = ""
        TxtSerie.Text = ""
        DtpEntrada.Value = Format$(Now, "dd/mm/yy")
        DtpEmissao.Value = Format$(Now, "dd/mm/yy")
        TxtValor.Text = ""
        TxtHistorico.Text = ""
        CboHistorico.Text = ""
        LblValorLiquido1.Caption = ""
'        TxtChamado.Text = ""
        
        'TabAuxiliar (0)
        MskPcr.Mask = ""
        MskPcr.Text = ""
        MskPcr.Mask = pPCRmascara
        CboPcr.Text = ""
        MskConta.Mask = ""
        MskConta.Text = ""
        CboPlanoContas.Text = ""
        CboIdentificador2.Text = ""
        
        DtpVencParcelaUnica.Value = Null
        DtpPagtoParcelaUnica.Value = Null
        TxtDocPagto1.Text = ""
        CboBancoForn1.Text = ""
        CboCCorrente.Text = ""
        TxtCCorrente.Text = ""
                
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
        TxtCCorrente2.Text = ""
        CboStatusduplicata.ListIndex = 0
        CboBancoForn2.Text = ""
        If TxtPrazo.Text <> "" Then
            DtpVencimento.Value = DateAdd("d", CInt(TxtPrazo.Text), DtpEmissao.Value)
        Else
          DtpVencimento.Value = Now
        End If
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
        TxtCCorrente.Text = CboCCorrente.BoundText
        
        Dim XLO_CHECAR As ADODB.Recordset
    
        'Checar se o nº do doc. pagto. já não está associado a um pagto nesta data e para esta conta
        If TxtDocPagto1.Text <> "" And CboStatus.Text = "Realizado" And XInserir = "I" Then
            If EntSaida = "S" Then ' Despesa
                SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex) & "' and dupl_tx_docpagrec ='" & TxtDocPagto1.Text & "' and dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "", "Estatico"
                If Not XLO_CHECAR.EOF Then
                    MsgBox "Este nº de Doc. de Pagto. já foi emitido!", vbCritical, "ATENÇÃO!"
                    TxtDocPagto1.SetFocus
                    Exit Sub
                End If
                XLO_CHECAR.Close
            Else
                SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='C' and coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex) & "' and dupl_tx_docpagrec ='" & TxtDocPagto1.Text & "'  and dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "", "Estatico"
                If Not XLO_CHECAR.EOF Then
                    MsgBox "Este nº doc. de crédito já foi registrado!", vbCritical, "ATENÇÃO!"
                    TxtDocPagto1.SetFocus
                    Exit Sub
                End If
                XLO_CHECAR.Close
            End If
        End If

        subConectarControleDadosNV DatFormaPagamento, "SELECT fopa_cd_codigo, banc_cd_codigo, CONVERT(VARCHAR, fopa_nr_numero) + ' - ' +  fopa_tx_descricao AS DescCombo FROM FormasPagamento WHERE banc_cd_codigo = " & DatContaCorrente.Recordset.Fields("banc_cd_codigo") & " ORDER BY fopa_nr_numero", Estatico
        
    End If
End Sub

Private Sub CboCCorrente2_Change()
    If CboCCorrente2.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente2.SelectedItem
        TxtCCorrente2.Text = CboCCorrente2.BoundText
        subConectarControleDadosNV DatFormaPagamento, "SELECT fopa_cd_codigo, banc_cd_codigo, CONVERT(VARCHAR, fopa_nr_numero) + ' - ' +  fopa_tx_descricao AS DescCombo FROM FormasPagamento WHERE banc_cd_codigo = " & DatContaCorrente.Recordset.Fields("banc_cd_codigo") & " ORDER BY fopa_nr_numero", Estatico
    End If
End Sub

Private Sub CboCriterios_Change()
    If CboCriterios.Text <> "" Then
        DatCriterios.Recordset.Bookmark = CboCriterios.SelectedItem
        subCarregaDadosNV Adodc1, "", "cocd_vl_percentual", "", "SELECT CentroResultados.cere_cd_Estruturado,CentroResultados.cere_tx_Descricao, PlanoContas.plco_cd_CodReduzido,PlanoContas.plco_tx_DescContabil, ComposicaoCriteriosDistribuicao.cocd_vl_percentual, ComposicaoCriteriosDistribuicao.cere_cd_Pcr, ComposicaoCriteriosDistribuicao.plco_cd_Conta " & _
            " FROM ComposicaoCriteriosDistribuicao JOIN CentroResultados ON (ComposicaoCriteriosDistribuicao.cere_cd_Pcr = CentroResultados.cere_cd_Pcr) JOIN PlanoContas ON (ComposicaoCriteriosDistribuicao.plco_cd_Conta = PlanoContas.plco_cd_Conta) " & _
            " WHERE ComposicaoCriteriosDistribuicao.crdi_cd_criteriodistribuicao = " & CStr(DatCriterios.Recordset.Fields("crdi_cd_criteriodistribuicao"))
    End If
End Sub

Private Sub CboDescontos_Change()
    'TESTE
    If CboDescontos.Text <> "" And EntSaida = "S" And EntrouDupl Then
        DatDescontosNF.Recordset.Bookmark = CboDescontos.SelectedItem
        If Vetor2(0, 1) <> Empty Then
            If DatDescontosNF.Recordset.Fields("desc_tx_tipobase") = "2" Then
                DtpVencDesconto.Value = Format$(Vetor2(0, 1), "DD/MM/YY")
            ElseIf Not IsNull(DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")) Then
                    Dim Ano As Integer
                    Dim Mes As Integer
                    
                    Ano = Year(CDate(Vetor2(0, 1)))
                    Mes = Month(CDate(Vetor2(0, 1)))
                    If Mes = 12 Then
                       Mes = 1
                       Ano = Ano + 1
                    Else
                       Mes = Mes + 1
                    End If
                    DtpVencDesconto.Value = Format$(DateSerial(Ano, Mes, DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")), "dd/mm/yy")
                Else
                    DtpVencDesconto.Value = Format(Now, "dd/mm/yy")
            End If
        End If
    End If
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.Bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        
        CboImovel.Enabled = True
        CboImovel.Text = ""
        subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='A' ORDER BY imov_cd_Imovel", Estatico
        
'        TabFinanceiro.Enabled = False
'        FraDatas.Enabled = False
'        TxtTitulo.Text = ""
'        CboTitulos.Text = ""
    Else
      Call cmdNovo_Click
    End If
    
End Sub

Private Sub CboEmpreendimento_DblClick(Area As Integer)
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub



Private Sub CboFornecedor_Change()
    Dim XLO_ADIANTAMENTO As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLD_SALDOADIANTAMENTO As Double
    
    If CboFornecedor.Text <> "" Then
        CboBancoForn1.Text = ""
        
        If (XInserir <> "A") Then CboTipoPagto1.ListIndex = 0
        
        DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
        If Not IsNull(DatFornecedor.Recordset.Fields("cere_cd_estruturado")) Then
            CboPcr.BoundText = DatFornecedor.Recordset.Fields("cere_cd_estruturado")
        End If
        
        If Not IsNull(DatFornecedor.Recordset.Fields("CodRedRecDesp")) Then
            CboPlanoContas.BoundText = DatFornecedor.Recordset.Fields("CodRedRecDesp")
        End If
        
        If (XInserir = "I") Then
            If (EntSaida = "S") And (PCliente = "QUALIDADOS") Then
                If Not IsNull(DatFornecedor.Recordset.Fields("CodRedCusto")) Then
                    cboPlanoContasCredito.BoundText = DatFornecedor.Recordset.Fields("CodRedCusto")
                End If
            End If
        End If
        
        If (XInserir <> "A") Then
            If Tipo = "1" And DatFornecedor.Recordset.Fields("focl_tx_classe") = "F" Then
                
                If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_tipopag")) Then
                    CboTipoPagto1.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_tipopag")
                Else
                    CboTipoPagto1.ListIndex = 0
                End If
                
            ElseIf Tipo = "0" And DatFornecedor.Recordset.Fields("focl_tx_classe") = "F" Then
                If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_tipopag")) Then
                    CboTipoPagto2.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_tipopag")
                Else
                    CboTipoPagto2.ListIndex = 0
                End If
            End If
            
            If XGB_AlertarSaldoAdiantamento = True Then
            
                If DatFornecedor.Recordset.Fields("focl_tx_classe") = "F" Then
                
                    XLT_SQL = "exec spFINAdiantamento @idEmpresa = " & PCodEmpresa & ", @idFornecedor  = " & CboFornecedor.BoundText & ", @dataSaldo  = '" & Format(Now, "MM/DD/YYYY") & "'"
                    SubQOpenRecordset XLO_ADIANTAMENTO, XLT_SQL, Estatico
                    
                    If Not XLO_ADIANTAMENTO.EOF Then
                       
                       XLD_SALDOADIANTAMENTO = CDbl(XLO_ADIANTAMENTO.Fields("Saldo").Value)
                       
                       If XLD_SALDOADIANTAMENTO > 0 Then
                           
                           If MsgBox("Este fornecedor possui um saldo de adiantamento de R$ " & Format(XLD_SALDOADIANTAMENTO, "standard") & ". Deseja realizar a prestação de contas agora?", vbInformation + vbYesNo, "INFORMAÇÃO") = vbYes Then
                                
                                XGV_FornNotaFiscal = CboFornecedor.BoundText
'                                FrmPrestacaoContas.Show 1
                                  
                           End If
                           
                       End If
                       
                    End If
                    
                    XLO_ADIANTAMENTO.Close
                    Set XLO_ADIANTAMENTO = Nothing
             
                End If
                
            End If

        End If
        
        If EntSaida = "S" Then
            subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
        End If
    End If
End Sub

Private Sub CboFornecedor_Click(Area As Integer)

'    If CboFornecedor.Text <> "" And EntSaida = "S" Then
'        subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn WHERE focl_cd_forncli = " & CboFornecedor.BoundText, Estatico
'    End If

    
End Sub

Private Sub CboFornecedor_DblClick(Area As Integer)
    Dim ChaveTemp As Variant
    Dim XInserirTemp As String
    Dim XLO_ChecarPermissao As ADODB.Recordset
    
    ChaveTemp = Chave
        
    If EntSaida = "E" Then
        FornClie = "C" 'Cliente
    Else
        FornClie = "F"  'Fornecedor
    End If
    'Pessoa = "J" 'Pessoa Jurídica
    XInserirTemp = XInserir
    
    If FrmNotasFiscais.Caption = "Cadastro de Contas a Receber" Or FrmNotasFiscais.Caption = "Alteração de Contas a Receber" Then
            SubQOpenRecordset XLO_ChecarPermissao, "SELECT Habilitacoes.func_cd_Operador, Habilitacoes.roti_tx_Programa, Habilitacoes.roti_tx_Ordem, Habilitacoes.habi_bl_Habilitado " & _
                                               "FROM Rotinas RIGHT JOIN Habilitacoes " & _
                                               "ON (Rotinas.roti_tx_Programa = Habilitacoes.roti_tx_Programa AND " & _
                                               "Rotinas.roti_tx_Ordem = Habilitacoes.roti_tx_Ordem) " & _
                                               "WHERE Rotinas.roti_tx_Programa = 'TelaFornecedores' AND func_cd_Operador = " & func_cd_operador & _
                                               " AND Rotinas.roti_tx_Descricao='Clientes: Cadastrar'", Estatico
                                               
            If Not (XLO_ChecarPermissao.EOF) Then
                If XLO_ChecarPermissao!habi_bl_habilitado = "S" Then
                        XFormForn = "FrmNotasFiscais"
                        FrmFornecedores.Show 1
                        Chave = ChaveTemp
                        XInserir = XInserirTemp
                        FrmNotasFiscais.DatFornecedor.Refresh
                        FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
                Else
                    MsgBox "Você não possui permissão para cadastrar Clientes.", 16, "AVISO"
                End If
            Else
                MsgBox "Você não possui permissão para cadastrar Clientes.", 16, "AVISO"
            End If

    ElseIf FrmNotasFiscais.Caption = "Cadastro de Contas a Pagar" Or FrmNotasFiscais.Caption = "Alteração de Contas a Pagar" Then
                SubQOpenRecordset XLO_ChecarPermissao, "SELECT Habilitacoes.func_cd_Operador, Habilitacoes.roti_tx_Programa, Habilitacoes.roti_tx_Ordem, Habilitacoes.habi_bl_Habilitado " & _
                                               "FROM Rotinas RIGHT JOIN Habilitacoes " & _
                                               "ON (Rotinas.roti_tx_Programa = Habilitacoes.roti_tx_Programa AND " & _
                                               "Rotinas.roti_tx_Ordem = Habilitacoes.roti_tx_Ordem) " & _
                                               "WHERE Rotinas.roti_tx_Programa = 'TelaFornecedores' AND func_cd_Operador = " & func_cd_operador & _
                                               " AND Rotinas.roti_tx_Descricao='Fornecedores: Cadastrar'", Estatico
            If Not (XLO_ChecarPermissao.EOF) Then
                If XLO_ChecarPermissao!habi_bl_habilitado = "S" Then
                        XFormForn = "FrmNotasFiscais"
                        FrmFornecedores.Show 1
                        Chave = ChaveTemp
                        XInserir = XInserirTemp
                        FrmNotasFiscais.DatFornecedor.Refresh
                        FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
                Else
                    MsgBox "Você não possui permissão para cadastrar Fornecedores.", 16, "AVISO"
                End If
            Else
                MsgBox "Você não possui permissão para cadastrar Fornecedores.", 16, "AVISO"
            End If
    End If
    
End Sub

Private Sub VerificaImpostosAReter()
    
    Dim XLO_RS As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLD_TOTALMES As Double
    Dim XLD_INICIO As Date
    Dim XLD_FINAL As Date
    Dim XLI_FORNECEDOR As Integer
    
    If CboFornecedor.Text <> "" Then
        
        If (DatFornecedor.Recordset("focl_tx_tipo").Value = "F") Then
             
            XLD_INICIO = CDate("01/" & Format(DtpEntrada.Value, "mm/yyyy"))
            XLD_FINAL = DateAdd("d", -1, DateAdd("m", 1, XLD_INICIO))
            XLI_FORNECEDOR = DatFornecedor.Recordset("focl_cd_forncli").Value
            
            XLT_SQL = "SELECT ISNULL(SUM(nofi_vl_valor), 0) AS Valor FROM NotasFiscais " & _
                        "WHERE " & _
                        "   empr_cd_empresa = " & CStr(PCodEmpresa) & " " & _
                        "AND nofi_tx_TipoMov = 'S' " & _
                        "AND (nofi_tx_status = 'C' OR nofi_tx_status = 'R') " & _
                        "AND nofi_dt_emissao BETWEEN " & FunNuloData(XLD_INICIO, NomeSgbd) & " AND " & FunNuloData(XLD_FINAL, NomeSgbd) & " " & _
                        "AND focl_cd_FornCli = " & XLI_FORNECEDOR
            SubQOpenRecordset XLO_RS, XLT_SQL, Estatico
            While Not XLO_RS.EOF And Not XLO_RS.BOF
                XLD_TOTALMES = XLO_RS!VALOR
                
                XLO_RS.MoveNext
            Wend
            XLO_RS.Close
            Set XLO_RS = Nothing
            
            If (XLD_TOTALMES >= 5000) Then
                MsgBox "Existem outros lançamentos que ultrapassam R$ 5.000,00. " & vbCrLf & vbCrLf & "Atente para as retenções dos Impostos Federais das notas lançadas", vbInformation, "Aviso"
            End If
            
        End If
        
    End If

End Sub

Private Sub CboFornecedor_LostFocus()
    If CboFornecedor.Text <> "" Then
        MskNumdocumento.Enabled = True
        If MskNumdocumento.Enabled = True Then
            MskNumdocumento.SetFocus
        End If
        
        VerificaImpostosAReter
    
    End If
End Sub

Private Sub CboHistorico_Change()
    If CboHistorico.Text <> "" And Not IsNull(CboHistorico.SelectedItem) Then
        DatHistorico.Recordset.Bookmark = CboHistorico.SelectedItem
        TxtHistorico.Text = CboHistorico.BoundText
    End If
End Sub

Private Sub CboHistorico_KeyPress(KeyAscii As Integer)
    If Len(CboHistorico) = 100 Then
        If KeyAscii <> 8 And CboHistorico.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub CboIdentificador2_LostFocus()
    If TabAuxiliar.TabEnabled(1) = True Then
        TabAuxiliar.Tab = 1
    End If
End Sub

Private Sub CboImovel_Change()
    Dim XLT_SQL As String
    If CboImovel.Text <> "" Then
        DatImovel.Recordset.Bookmark = CboImovel.SelectedItem
        
        '
        ' TODO:
        ' -
        
'        If Not IsNull(DatImovel.Recordset.Fields("cont_dt_base")) Then
'            LblDtContrato2.Caption = Format(DatImovel.Recordset.Fields("cont_dt_base"), "dd/MM/yy")
'        End If
        
'        LblContrato2.Caption = DatImovel.Recordset.Fields!cont_cd_Contrato
'        LblDtVenda2.Caption = Format(DatImovel.Recordset.Fields!cont_dt_Venda, "dd/MM/yy")
'        LblTipoImovel2.Caption = DatImovel.Recordset.Fields!tiim_tx_Descricao
        
        If XFormulario = "TelaFinTitulos" Then
          'Trazer todos os títulos já pagos
          XLT_SQL = "SELECT (CONVERT(varchar(8), titu_dt_vencimento, 3) " & _
             "+ '  -  ' + titu_cd_Plano + '.' + titu_cd_Parcela + '.' + titu_cd_Residuo " & _
             "+ '  -  ' + left(tipl_tx_descricao,17) + space (15 -len(left(tipl_tx_descricao,17))) + ' - ' + natureza+space(18 - len(natureza))) as Parcela, * " & _
             "FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " " & _
             "AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
             "AND not titu_dt_Pagamento is null " & _
             "AND imov_cd_Imovel = '" & CboImovel.BoundText & "' ORDER BY titu_dt_vencimento, Titulo"
        Else
            'Trazer todos os títulos desse contrato que estão a vencer
            XLT_SQL = "SELECT (CONVERT(varchar(8), titu_dt_vencimento, 3) " & _
                "+ '  -  ' + titu_cd_Plano + '.' + titu_cd_Parcela + '.' + titu_cd_Residuo " & _
                "+ '  -  ' + left(tipl_tx_descricao,17) + space (17 -len(left(tipl_tx_descricao,17))) + ' - ' + natureza+space(18 - len(natureza))) as Parcela, * " & _
                "FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " " & _
                "AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
                "AND titu_dt_Pagamento is null " & _
                "AND imov_cd_Imovel = '" & CboImovel.BoundText & "' " & _
                "AND status='ATIVO' ORDER BY titu_dt_vencimento, Titulo"
        End If
'        subConectarControleDados DatTitulos, XLT_SQL, Estatico
'        TabFinanceiro.Enabled = False
'        FraDatas.Enabled = False
'        CboTitulos.Text = ""
'        CboTitulos.Enabled = True
'        TxtTitulo.Text = ""

        CboStatus.Enabled = True
        
    End If
    
End Sub


'Private Sub CboPcr_Change()
'
'    Dim XLO_CHECAR As ADODB.Recordset
'
'    If CboPcr.Text <> "" Then
'        DatPcr.Recordset.Bookmark = CboPcr.SelectedItem
'        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
'            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
'            CboPcr.BoundText = MskPcr.Text
'            CboPcr.SetFocus
'            'Cássio 08/07/2016
'            'Se o usuário informou um CR sintético através da máscara, não permite a vinculação.
'            If Not IsNull(CboPcr.SelectedItem) Then DatPcr.Recordset.Bookmark = CboPcr.SelectedItem
'
'            If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
'                MskPcr.Mask = ""
'                MskPcr.Text = ""
'                MskPcr.Mask = pPCRmascara
'                CboPcr.Text = ""
'            End If
'
'            Exit Sub
'        ElseIf DatPcr.Recordset.Fields("cere_tx_status") = "I" Then
'            'Cássio 21/06/2013
'            'Se o centro de custo for inativo não permite selecioná-lo
'            CboPcr.BoundText = MskPcr.Text
'            Exit Sub
'        End If
'
'        MskPcr.Text = CboPcr.BoundText
'        Verifica_DataEmissao 'Inserido em 31/07/08 - Larissa
'
''        'inserido em 07/07/10 - Patrícia
'        SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr"), Estatico
'        If Not (XLO_CHECAR.EOF) Then
'            If CboPlanoContas.Text = "" Then
'                If EntSaida = "E" Then
'                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'                ElseIf EntSaida = "S" Then
'                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'                End If
'            End If
'            XLO_CHECAR.Close
'            Set XLO_CHECAR = Nothing
'        End If
'
'    End If
'End Sub

'
' Subrotina para carregar/manipular o compbo de Centros de Custo (tela de Dados Gerais NF)
' ----------------------------------------------------------------------------------------
Private Sub CboPcr_Change()

    Dim XLO_CHECAR As ADODB.Recordset
    
    If CboPcr.Text <> "" Then
    
        DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
        
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            
            'Cássio 08/07/2016
            'Se o usuário informou um CR sintético através da máscara, não permite a vinculação.
            If Not IsNull(CboPcr.SelectedItem) Then DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
            
            If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
                MskPcr.Mask = pPCRmascara
                MskPcr.Text = ""
                CboPcr.Text = ""
            End If
            
            Exit Sub
            
        ElseIf DatPCR.Recordset.Fields("cere_tx_status") = "I" Then
            'Cássio 21/06/2013
            'Se o centro de custo for inativo não permite selecioná-lo
            CboPcr.BoundText = MskPcr.Text
            Exit Sub
        End If
            
        MskPcr.Text = CboPcr.BoundText
        Verifica_DataEmissao 'Inserido em 31/07/08 - Larissa
        
       ' Se informação de vinculação a nível de empresa for verdadeira
'       If pVinculaCrCc = True Then
        
            ' Alterado em 21/03/22 - Paulo Garcia
            ' Força a (re)carga do CboPlanoContas sempre que haja mudança na opção do centro de custo,
            ' pelo usuário, obrigando a escolha de nova conta que esteja associada ao centro de custo
            ' escolhido.
            ' Testa se o formulário está carregado (visivel e habilitado), única garantia de que a
            ' execução desta subrotina será em função da opção do usuário
            ' ----------------------------------------------------------------------------------------
            If (CboPcr.Visible = True And CboPcr.Enabled = True) Then
            
                ' inserido em 07/07/10 - Patrícia
                SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
                
                ' Recarrega DatConta (dataset para o combo de contas) com as contas vinculadas ao C/C
                If Not (XLO_CHECAR.EOF) Then
        
                    If EntSaida = "E" Then
                        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                    ElseIf EntSaida = "S" Then
                        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                    End If
                                                     
                End If
                
                XLO_CHECAR.Close
                Set XLO_CHECAR = Nothing
                    
                ' Limpa o combo de Contas Contábeis se código CR selecionado difere do anterior
                If (XFT_CR <> MskPcr.Text) Then
                
                    MskConta.Text = ""
                    CboPlanoContas.Text = ""
                    CboPlanoContas.BoundText = 0
                        
                End If
                    
                XFT_CR = MskPcr.Text
                
            End If
            
        ' End If
                    
    End If
    
End Sub

'Private Sub CboPcr2_Change()
'    Dim XLO_CHECAR As ADODB.Recordset
'    If CboPcr2.Text <> "" Then
'        DatPcr.Recordset.Bookmark = CboPcr2.SelectedItem
'        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
'            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
'            CboPcr2.BoundText = MskPcr2.Text
'            CboPcr2.SetFocus
'            Exit Sub
'        End If
'        MskPcr2.Text = CboPcr2.BoundText
'        Verifica_DataEmissao
'        'inserido em 07/07/10 - Patrícia
'        SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr"), Estatico
'        If Not (XLO_CHECAR.EOF) Then
'            If CboPlanoContas2.Text = "" Then
'                If EntSaida = "E" Then
'                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'                ElseIf EntSaida = "S" Then
'                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPcr.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'                End If
'            End If
'            XLO_CHECAR.Close
'            Set XLO_CHECAR = Nothing
'        End If
'
'    End If
'End Sub

'
' Subrotina para carregar/manipular o compbo de Centros de Custo (tela de Distribuição NF)
' ----------------------------------------------------------------------------------------
Private Sub CboPcr2_Change()

    Dim XLO_CHECAR As ADODB.Recordset
    
    If CboPcr2.Text <> "" Then
    
        DatPCR.Recordset.Bookmark = CboPcr2.SelectedItem
        
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr2.BoundText = MskPcr2.Text
            CboPcr2.SetFocus
            Exit Sub
        End If
        
        MskPcr2.Text = CboPcr2.BoundText
        Verifica_DataEmissao
        
       ' Se informação de vinculação a nível de empresa for verdadeira
'       If pVinculaCrCc = True Then
        
            ' Alterado em 21/03/22 - Paulo Garcia
            ' Força a (re)carga do CboPlanoContas sempre que haja mudança na opção do centro de custo,
            ' pelo usuário, obrigando a escolha de nova conta que esteja associada ao centro de custo
            ' escolhido.
            ' Testa se o formulário está carregado (visivel e habilitado), única garantia de que a
            ' execução desta subrotina será em função da opção do usuário
            ' ----------------------------------------------------------------------------------------
            If (CboPcr2.Visible = True And CboPcr2.Enabled = True) Then
            
                'inserido em 07/07/10 - Patrícia
                SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
               
                ' Recarrega DatConta (dataset para o combo de contas) com as contas vinculadas ao C/C
                If Not (XLO_CHECAR.EOF) Then
        
                    If EntSaida = "E" Then
                        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                    ElseIf EntSaida = "S" Then
                        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                    End If
        
                End If
                                    
                XLO_CHECAR.Close
                Set XLO_CHECAR = Nothing
                
                ' Limpa o combo de Contas Contábeis se código CR selecionado difere do anterior
                If (XFT_CR <> MskPcr2.Text) Then
                
                    MskConta2.Text = ""
                    CboPlanoContas2.Text = ""
                    CboPlanoContas2.BoundText = 0
                        
                End If
                    
                XFT_CR = MskPcr2.Text
                
            End If
                    
        ' End If
                    
    End If
    
End Sub

'Private Sub cboPlanoContas_Change()
'    Dim XLO_CHECAR As ADODB.Recordset
'
'    If CboPlanoContas.Text <> "" Then
'
'        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
'
'        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
'            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
'            CboPlanoContas.BoundText = Mskconta.Text
'            CboPlanoContas.SetFocus
'            Exit Sub
'        End If
'        XLT_VALORDIFFNOTA = CreditoDebitoCC()
'
'        '******* Comentado para resolver problema de limpar o CR na tela *****
'       '*** 02-12-2021 | Dino                                           *****
'       '*** Solicitação de Marco Antônio, Cooddenador financeiro        *****
'       '*********************************************************************
'        'inserido em 07/07/10 - Patrícia
''        If CboPcr.Text <> "" Then
''            SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & " and cere_cd_pcr= " & DatPcr.Recordset.Fields("cere_cd_pcr"), Estatico
''            If Not (XLO_CHECAR.EOF) Then
''                subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
''
''                If (CboPcr.BoundText <> "") Then
''                        'DatPcr.Recordset.Find ("cere_tx_descricao = '" & CboPcr.BoundText & "'")
''                        'Cássio 19/02/2013
''                        DatPcr.Recordset.Find ("DESCCOMBO = '" & CboPcr.BoundText & "'")
''                        If (DatPcr.Recordset.EOF) Then
''                            CboPcr.Text = ""
''                        End If
''                End If
''            End If
''            XLO_CHECAR.Close
''            Set XLO_CHECAR = Nothing
''        Else
''            SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta"), Estatico
''            If Not (XLO_CHECAR.EOF) Then
''                subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
''            End If
''            XLO_CHECAR.Close
''            Set XLO_CHECAR = Nothing
''
''        End If
'        Mskconta.Text = CboPlanoContas.BoundText
'    End If
'End Sub

Private Sub cboPlanoContas_Change()

    Dim XLO_CHECAR As ADODB.Recordset
    
    If CboPlanoContas.Text <> "" Then
    
        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
        
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = MskConta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        
        XLT_VALORDIFFNOTA = CreditoDebitoCC()
        
        MskConta.Text = CboPlanoContas.BoundText
        
    End If
    
End Sub

Function CreditoDebitoCC() As Double
    Dim XLO_VALORENTRADA As ADODB.Recordset
    Dim XLO_VALORSAIDA As ADODB.Recordset
    Dim XLT_SQL As String
    Dim ValorEntrada As Double
    Dim ValorSaida As Double
    Dim CREstruturado As String
    
    CREstruturado = "'" & DatConta.Recordset.Fields("plco_cd_Estruturado") & "'"
    
    XLT_SQL = " select sum(distribuicao.dist_vl_valor) SOMAVALOR " & _
            " From notasfiscais inner join distribuicao on distribuicao.nofi_cd_notafiscal = notasfiscais.nofi_cd_notafiscal " & _
            " inner join planocontas on planocontas.plco_cd_Conta = distribuicao.plco_cd_Conta" & _
            " WHERE notasfiscais.nofi_tx_TipoMov = 'E' AND notasfiscais.nofi_tx_status = 'R' AND YEAR(notasfiscais.nofi_dt_emissao) = YEAR(GETDATE()) " & _
            " AND planocontas.plco_cd_Estruturado = " & CREstruturado & ""
            
    SubQOpenRecordset XLO_VALORENTRADA, XLT_SQL, Estatico
    
    If Not XLO_VALORENTRADA.EOF Then
        If Not IsNull(XLO_VALORENTRADA!SOMAVALOR) Then
            ValorEntrada = XLO_VALORENTRADA!SOMAVALOR
        Else
            ValorEntrada = 0
        End If
    End If
    
    XLT_SQL = " select sum(distribuicao.dist_vl_valor) SOMAVALOR " & _
            " From notasfiscais inner join distribuicao on distribuicao.nofi_cd_notafiscal = notasfiscais.nofi_cd_notafiscal " & _
            " inner join planocontas on planocontas.plco_cd_Conta = distribuicao.plco_cd_Conta" & _
            " WHERE  notasfiscais.nofi_tx_TipoMov = 'S' AND notasfiscais.nofi_tx_status = 'R' AND YEAR(notasfiscais.nofi_dt_emissao) = YEAR(GETDATE()) " & _
            " AND planocontas.plco_cd_Estruturado = " & CREstruturado & ""
            
    SubQOpenRecordset XLO_VALORSAIDA, XLT_SQL, Estatico
    
    If Not XLO_VALORSAIDA.EOF Then
        If Not IsNull(XLO_VALORSAIDA!SOMAVALOR) Then
            ValorSaida = XLO_VALORSAIDA!SOMAVALOR
        Else
            ValorSaida = 0
        End If
    End If
    
    CreditoDebitoCC = ValorEntrada - ValorSaida
    
End Function

Private Sub CboPlanoContas_DblClick(Area As Integer)

    XFormForn = "FrmNotasFiscais"
'    FrmContaContabil.Show 1
    MskConta.Text = XGV_CodContabil
    mskConta_LostFocus
                
End Sub

Private Sub CboPlanoContas2_DblClick(Area As Integer)

    XFormForn = "FrmNotasFiscais"
'    FrmContaContabil.Show 1
    MskConta2.Text = XGV_CodContabil
    MskConta2_LostFocus
                
End Sub

Private Sub CboPlanoContas2_Change()
    Dim XLO_CHECAR As ADODB.Recordset
    If CboPlanoContas2.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas2.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas2.BoundText = MskConta2.Text
            CboPlanoContas2.SetFocus
            Exit Sub
        End If
        'inserido em 07/07/10 - Patrícia
        If CboPcr2.Text <> "" Then
            SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & " and cere_cd_pcr= " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
            If Not (XLO_CHECAR.EOF) Then
                'subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
                'inserido em 05/01/17 - Cássio Medeiros
                'Evitar que o indíce do centro de custo selecionado pelo usuário seja alterado após nova atualização do objeto "DatPcr"
                If Not DatPCR.Recordset.EOF Then
                     DatPCR.Recordset.Find ("DESCCOMBO = '" & CboPcr2.Text & "'")
                End If
                
            End If
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
        Else
            SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta"), Estatico
            If Not (XLO_CHECAR.EOF) Then
                If CboPcr2.Text = "" Then
                    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
                End If
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
            End If
        End If
        MskConta2.Text = CboPlanoContas2.BoundText
    End If
End Sub

Private Sub cboPlanoContasCredito_Change()
    Dim XLO_CHECAR As ADODB.Recordset
    
    If cboPlanoContasCredito.Text <> "" Then
    
        DatContaContabilCredito.Recordset.Bookmark = cboPlanoContasCredito.SelectedItem
        If DatContaContabilCredito.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            cboPlanoContasCredito.BoundText = MskcontaCredito.Text
            cboPlanoContasCredito.SetFocus
            Exit Sub
        End If
        
    '******* Comentado para resolver problema de limpar o CR na tela *****
    '*** 02-12-2021 | Dino                                           *****
    '*** Solicitação de Marco Antônio, Cooddenador financeiro        *****
    '*********************************************************************
    '     XLT_VALORDIFFNOTA = CreditoDebitoCC()
        
        'inserido em 07/07/10 - Patrícia
    '    If CboPcr.Text <> "" Then
    '       SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatContaContabilCredito.Recordset.Fields("plco_cd_Conta") & " and cere_cd_pcr= " & DatPcr.Recordset.Fields("cere_cd_pcr"), Estatico
    '        If Not (XLO_CHECAR.EOF) Then
    '            subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatContaContabilCredito.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico

    '            If (CboPcr.BoundText <> "") Then
    '                    'DatPcr.Recordset.Find ("cere_tx_descricao = '" & CboPcr.BoundText & "'")
    '                    'Cássio 19/02/2013
    '                    DatPcr.Recordset.Find ("DESCCOMBO = '" & CboPcr.BoundText & "'")
    '                    If (DatPcr.Recordset.EOF) Then
    '                        CboPcr.Text = ""
    '                    End If
    '            End If
    '        End If
    '        XLO_CHECAR.Close
    '        Set XLO_CHECAR = Nothing
    '    Else
    '        SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta"), Estatico
    '        If Not (XLO_CHECAR.EOF) Then
    '            subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
    '        End If
    '        XLO_CHECAR.Close
    '        Set XLO_CHECAR = Nothing

    '    End If
    
        MskcontaCredito.Text = cboPlanoContasCredito.BoundText
        
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
        TxtCodTipoDoc.Enabled = True
    Else
        HabilitaControles
    End If
End Sub
'Private Sub CboStatus_LostFocus()
'    If XGI_QTDDIASPROVISAO >= 0 And XGT_TIPOEXPORTACAO = "PC" And CboStatus.Text <> "Previsto" Then
'        XFB_PROVISIONA = True
'    Else
'        XFB_PROVISIONA = False
'    End If
'End Sub

Private Sub CboStatusduplicata_click()
    If CboStatusduplicata.ListIndex = 1 Then
        TxtDocPagto2.Enabled = True
        DtpPagamento.Enabled = True
        DtpPagamento.Value = DtpVencimento.Value
        Me.CmdConfirmaduplicata.Enabled = True
    ElseIf CboStatusduplicata.ListIndex = 2 Then
        Me.CmdConfirmaduplicata.Enabled = False
    Else
        Me.CmdConfirmaduplicata.Enabled = True
        TxtDocPagto2.Text = ""
        TxtDocPagto2.Enabled = False
        DtpPagamento.Value = Null
        DtpPagamento.Enabled = False
    End If
End Sub

Private Sub CboTipoDoc_Change()
    If CboTipoDoc.Text <> "" Then
        
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        TxtCodTipoDoc.Text = CboTipoDoc.BoundText
        HabilitaControles
        
        If Not IsNull(CboTipoDoc.SelectedItem) Then
            DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
            If Not IsNull(DatTipoDoc.Recordset.Fields("tido_tx_duplicata")) Then
                Tipo = DatTipoDoc.Recordset.Fields("tido_tx_duplicata")
            End If
        End If
    
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
        
        'Desabilita combo de contas correntes - Larissa 31/07/08
        If pVinculaCrCc = True Then
            CboPlanoContas.Enabled = False
            CboPlanoContas2.Enabled = False
        End If
        
    End If
End Sub

Private Sub CboTipoPagto1_CLICK()
    Dim Achou As Boolean
    
    Achou = False
    If EntSaida = "S" Then
        If CboTipoPagto1.ListIndex = 2 Then
            'trazer o banco padrão
            If Not (DatBancoFornecedor.Recordset Is Nothing) Then
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
            Else
                MsgBox "O Fornecedor selecionado deve possuir um Banco relacionado.", vbCritical + vbOKOnly, "ATENÇÃO"
                CboTipoPagto1.SetFocus
            End If
            CboBancoForn1.Enabled = True
        ElseIf CboTipoPagto1.ListIndex = 4 Then 'DÉBITO EM C/C
            CboBancoForn1.Text = ""
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
        ElseIf CboTipoPagto2.ListIndex = 4 Then 'DÉBITO EM C/C
            CboBancoForn2.Text = ""
            CboBancoForn2.Enabled = True
        Else
            CboBancoForn2.Text = ""
            CboBancoForn2.Enabled = False
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
       
       CmdLimparIdentif_Click
       CboIdentificador2.Enabled = False
       CmdLimparIdentif.Enabled = False
       
       MskConta.Enabled = False
       CboPlanoContas.Enabled = False
       TabNotasFiscais.TabEnabled(3) = True
       PainelCriterios.Enabled = False
       PainelCriterios.Visible = False
    Else
       MskConta.Text = ""
       CboPlanoContas.Text = ""
       CboPcr.Text = ""
       TabAuxiliar.TabEnabled(0) = True
       MskPcr.Enabled = True
       CboPcr.Enabled = True
       CboIdentificador2.Enabled = True
       CmdLimparIdentif.Enabled = True
       MskConta.Enabled = True
       CboPlanoContas.Enabled = True
       'subLimpaVetor Vetor3, TDBGrid3
       TabNotasFiscais.TabEnabled(3) = False
       TabNotasFiscais.Tab = 0
    End If
    XDistribuido = "N"
End Sub

Private Sub CmdAlteraDesconto_Click()
    If TabNotasFiscais.Tab = 2 Then
        
        If TDBGrid1.Columns("NumDarf").Text <> "" Then
            MsgBox ("Este imposto já foi recolhido portanto não poderá ser alterado!"), vbExclamation, " Informação"
            Exit Sub
        End If
        FrmNotasFiscais.MousePointer = vbHourglass
        ChaveD = -1
        XDesconto = "A"
        
        PrepararInsercao
        CboDescontos.BoundText = TDBGrid1.Columns("codigo").Text
        xcod = CboDescontos.BoundText
        If IsDate(TDBGrid1.Columns("Vencimento")) Then
            DtpVencDesconto.Value = Format(TDBGrid1.Columns("Vencimento").Text, "dd/mm/yy")
        End If
        
        
        TxtValorDesconto.Text = Format(TDBGrid1.Columns("Valor").Text, "standard")
        
'        pandescontos.Visible = True
'        pandescontos.Enabled = True
        CboDescontos.SetFocus
        DesabilitaBotoes
        FrmNotasFiscais.MousePointer = vbDefault
    End If
End Sub

Private Sub CmdAlteraDistribuicao_Click()
    If TabNotasFiscais.Tab = 3 Then
        
        DesabilitaBotoes
        TDBGrid3.Enabled = False
        
'        PanDistribuicao.Visible = True
'        PanDistribuicao.Enabled = True
        XDistribuicao = "A"
        
        'Lucas Santiago - 27.01.09 - tratamento de variável nula ou vazia.
        If IsNull(TDBGrid3.Columns(1).Text) = False And TDBGrid3.Columns(1).Text <> "" Then
            MskPcr2.Text = TDBGrid3.Columns(1).Text & Space(Len(pPCRmascara) - Len(TDBGrid3.Columns(1).Text))
        End If
             
        MskConta2.Text = TDBGrid3.Columns(3).Text
        CboPlanoContas2.BoundText = MskConta2.Text
        CboPcr2.BoundText = MskPcr2.Text
        XPcr = CboPcr2.BoundText
        
        Verifica_DataEmissao 'Inserido em 28/01/09 - Larissa
        xcc = MskConta2.Text
        
        'Lucas Santiago - 27.01.09 - tratamento de variável nula ou vazia.
        If IsNull(TDBGrid3.Columns(4).Text) = False And TDBGrid3.Columns(4).Text <> "" Then
            TxtVldistribuicao.Text = Format(CDbl(TDBGrid3.Columns(4).Text), "standard")
        End If
        
        'Lucas Santiago - 27.01.09 - tratamento de variável nula ou vazia.
        If IsNull(TDBGrid3.Columns(7).Text) = False And TDBGrid3.Columns(7).Text <> "" Then
            TxtPercDistribuicao.Text = Mid(TDBGrid3.Columns(7).Text, 1, (Len(TDBGrid3.Columns(7).Text) - 1))
        End If

        CboIdentificador.BoundText = TDBGrid3.Columns(8).Text
        CboIdentificador.Text = TDBGrid3.Columns(9).Text
        
        DesabilitaBotoes
        CmdGravar.Enabled = False
        
        If MskPcr2.Enabled = True Then MskPcr2.SetFocus
        
        Exit Sub
    End If

TrataErro:

  funTrataErros (ComMensagem)

End Sub

Private Sub CmdAlteraduplicata_Click()
    Dim Status As String
    Dim x As Integer
    
    If TabNotasFiscais.Tab = 1 Then
        
        DesabilitaBotoes
        TDBGrid1.Enabled = False
'        PnlCompDuplicata.Visible = True
'        PnlCompDuplicata.Enabled = True
        
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
        ElseIf TDBGrid2.Columns("CodStatus") = "X" Then
            Status = 2
        End If
        
        CboStatusduplicata.ListIndex = Status
        
        'Desabilita o combo de status se a duplicata for realizada
        If TDBGrid2.Columns(10).Caption <> "" Then
            CboStatusduplicata.Enabled = False
        End If
       
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
            If TDBGrid2.Columns("Data Pag.") <> Null Or TDBGrid2.Columns("Data Pag.") <> "" Then
                DtpPagamento.Enabled = True
                DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Pag."), "DD/MM/YY")
                TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Pag."))
            Else
                DtpPagamento.Value = Null
                DtpPagamento.Enabled = False
                TxtDocPagto2.Text = ""
                TxtDocPagto2.Enabled = False
            End If
        Else
            If TDBGrid2.Columns("Data Rec.") <> Null Or TDBGrid2.Columns("Data Rec.") <> "" Then
                DtpPagamento.Enabled = True
                DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Rec."), "DD/MM/YY")
                TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Rec."))
            Else
                DtpPagamento.Value = Null
                DtpPagamento.Enabled = False
                TxtDocPagto2.Text = ""
                TxtDocPagto2.Enabled = False
            End If
        End If
        CboCCorrente2.BoundText = TDBGrid2.Columns("CodConta")
        TxtCCorrente2.Text = TDBGrid2.Columns("CodConta")
        CboBancoForn2.BoundText = TDBGrid2.Columns("CodBancoForn")
        XVlDuplicata = CDbl(FunNuloVal(TxtValorDuplicata.Text))
        CboFormaPagamento2.BoundText = TDBGrid2.Columns("Cód Forma Pagto")
        
        If CboStatus.Text = "Realizado" And CboStatusduplicata.Text <> "Cancelado" Then
            CboStatusduplicata.Enabled = False
            CboStatusduplicata.ListIndex = 1
        ElseIf CboStatus.Text = "Realizado" And CboStatusduplicata.Text = "Cancelado" Then
            CboStatusduplicata.Enabled = False
            CboStatusduplicata.ListIndex = 2
        ElseIf CboStatus.Text = "Cancelado" Then
            CboStatusduplicata.Enabled = False
            CboStatusduplicata.ListIndex = 2
        Else
            CboStatusduplicata.Enabled = True
        End If
    
        If TDBGrid2.Columns("CodStatus") = "R" Then
            If Not FunVerifica_Permissao("TelaNotasFiscais", "3") Then
                DesabilitaRealizado "Duplicata", False
            Else
                DesabilitaRealizado "Duplicata", True
            End If
        Else
            DesabilitaRealizado "Duplicata", True
        End If
        
        'verifica se é uma duplicata cancelada para desabilitar o botão de cancelar
        If TDBGrid2.Columns("CodStatus") = "X" Then
            Me.CmdConfirmaduplicata.Enabled = False
        Else
            Me.CmdConfirmaduplicata.Enabled = True
        End If
        
        'Se a duplicata estiver realizada
        If TDBGrid2.Columns("CodStatus") = "R" Then
            CboStatusduplicata.Enabled = False
            'Se a duplicata for consolidada, não permite alteração
            If TDBGrid2.Columns("Consolidado") = "S" Then
                CboTipoPagto2.Enabled = False
                TxtDocPagto2.Enabled = False
                TxtCCorrente2.Enabled = False
                CboCCorrente2.Enabled = False
                DtpPagamento.Enabled = False
                DtpVencimento.Enabled = False
                TxtDuplicata.Enabled = False
                TxtValorDuplicata.Enabled = False
                CmdConfirmaduplicata.Enabled = False
            End If
        End If
    End If
    
    'Preenchendo para exclusão do verso de cheque
    XLT_DOCPAGRECVERSO = Me.TxtDocPagto2.Text
    XLT_CONTAVERSO = Me.TxtCCorrente2.Text
    
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
    
        If CInt(TxtNParcelas.Text) > 199 Then
            MsgBox "O número máximo de parcelas é 199.", vbCritical, "ATENÇÃO"
            TxtNParcelas.SetFocus
            Exit Sub
        End If
        
        XFT_INSDUPLMULTIPLA = "A"
            
        If XDuplicata = "A" Then
           ValorTotalDuplicata = ValorTotalDuplicata - XVlDuplicata
        End If
    
        If CDbl(LblValorLiquido1.Caption) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
            MsgBox "A soma das duplicatas excedeu o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
            If TxtValorDuplicata.Enabled = True Then TxtValorDuplicata.SetFocus
            CmdInsereduplicata_Click
        Else
            For XLI_Y = 1 To Val(TxtNParcelas.Text)    'Grava Duplicatas múltiplas
                If Vetor2.UpperBound(1) > 0 And Vetor2.UpperBound(1) < 25 Then
                    XLT_PARCELA = Chr(66 + Vetor2.UpperBound(1))
                ElseIf Vetor2.UpperBound(1) = 0 Then
                    If Vetor2(0, 0) <> Empty Then
                        XLT_PARCELA = "B"
                    Else
                        XLT_PARCELA = "A"
                    End If
                ElseIf Vetor2.UpperBound(1) >= 25 Then
                    Select Case Vetor2.UpperBound(1)
                        Case Is <= 33
                            XLT_PARCELA = "Z" & Chr(24 + Vetor2.UpperBound(1))
                        Case Is <= 43
                            XLT_PARCELA = "Z1" & Chr(14 + Vetor2.UpperBound(1))
                        Case Is <= 53
                            XLT_PARCELA = "Z2" & Chr(4 + Vetor2.UpperBound(1))
                        Case Is <= 63
                            XLT_PARCELA = "Z3" & Chr(Vetor2.UpperBound(1) - 6)
                        Case Is <= 73
                            XLT_PARCELA = "Z4" & Chr(Vetor2.UpperBound(1) - 16)
                        Case Is <= 83
                            XLT_PARCELA = "Z5" & Chr(Vetor2.UpperBound(1) - 26)
                        Case Is <= 93
                            XLT_PARCELA = "Z6" & Chr(Vetor2.UpperBound(1) - 36)
                        Case Is <= 103
                            XLT_PARCELA = "Z7" & Chr(Vetor2.UpperBound(1) - 46)
                        Case Is <= 113
                            XLT_PARCELA = "Z8" & Chr(Vetor2.UpperBound(1) - 56)
                        Case Is <= 123
                            XLT_PARCELA = "Z9" & Chr(Vetor2.UpperBound(1) - 66)
                        Case Is <= 133
                            XLT_PARCELA = "Z10" & Chr(Vetor2.UpperBound(1) - 76)
                        Case Is <= 143
                            XLT_PARCELA = "Z11" & Chr(Vetor2.UpperBound(1) - 86)
                        Case Is <= 153
                            XLT_PARCELA = "Z12" & Chr(Vetor2.UpperBound(1) - 96)
                        Case Is <= 163
                            XLT_PARCELA = "Z13" & Chr(Vetor2.UpperBound(1) - 106)
                        Case Is <= 173
                            XLT_PARCELA = "Z14" & Chr(Vetor2.UpperBound(1) - 116)
                        Case Is <= 183
                            XLT_PARCELA = "Z15" & Chr(Vetor2.UpperBound(1) - 126)
                        Case Is <= 193
                            XLT_PARCELA = "Z16" & Chr(Vetor2.UpperBound(1) - 136)
                        Case Is <= 203
                            XLT_PARCELA = "Z17" & Chr(Vetor2.UpperBound(1) - 146)
                    End Select
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
            
            '******************REGISTRA LOG DA OPERAÇÃO**************************
            If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(10, 2) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(9, 2) As Variant
                
                'Nome dos campos referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 0) = "Doc Pagto"
                XGM_MATRIZLOG(9, 0) = "Data Pagto"
                
                'Dados referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 1) = FunNulo(TxtDocPagto2.Text)
                XGM_MATRIZLOG(9, 1) = Format$(DtpPagamento.Value, "DD/MM/YYYY")
            Else
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(8, 2) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(7, 2) As Variant
            End If
            
            'Nome dos campos da duplicata
            XGM_MATRIZLOG(0, 0) = "Nota Fiscal"
            XGM_MATRIZLOG(1, 0) = "Cod"
            XGM_MATRIZLOG(2, 0) = "Valor"
            XGM_MATRIZLOG(3, 0) = "Status"
            XGM_MATRIZLOG(4, 0) = "Data Venc"
            XGM_MATRIZLOG(5, 0) = "Tipo Pagto"
            XGM_MATRIZLOG(6, 0) = "Conta Corrente"
            If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 0) = "Banco Fornecedor"

            'Dados da duplicata
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 1) = XLT_PARCELA
            XGM_MATRIZLOG(2, 1) = Vetor2(XLI_X, 2)
            XGM_MATRIZLOG(3, 1) = "Compromissado"
            XGM_MATRIZLOG(4, 1) = Format$(XLD_VENCIMENTO, "DD/MM/YYYY")
            XGM_MATRIZLOG(5, 1) = Vetor2(XLI_X, 8)
            XGM_MATRIZLOG(6, 1) = Vetor2(XLI_X, 12)
            If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 1) = Vetor2(XLI_X, 14)
            
            'Gera descrição de log da operação
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_RECEBER)
            
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '*****************************FIM LOG - PARTE 3********************************
            
            
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
    Resume ResumoDoErro
    End If
    PanEntMultiplas.Enabled = False
    PanEntMultiplas.Visible = False
    TabNotasFiscais.Enabled = True
    CmdGravar.Enabled = True
    CmdDesistir.Enabled = True
    
End Sub

Private Sub CmdConfirmaDesconto_Click()
    Dim XLT_SQL As String
    Dim XLO_DESCONTO As Object

    If FunObrigatorioCBO(CboDescontos, "É necessário selecionar um desconto!") Then
        CboDescontos.SetFocus
        Exit Sub
    End If
    
    If (TxtValorDesconto.Text = "" Or CDbl(TxtValorDesconto.Text) = 0) Then 'Alteração em 19/05/09 - Patrícia. Anterior: If FunObrigatorioTXT(TxtValorDesconto, "Valor do desconto é obrigatório!") Then
        MsgBox "Valor do desconto tem que ser maior que zero!", vbCritical, "CAMPO OBRIGATÓRIO"
        TxtValorDesconto.SetFocus
        Exit Sub
    End If
    
    '*******Verificar se o Desconto/Imposto possui conta contábil***********
    XLT_SQL = "SELECT plco_cd_conta FROM plcDescontosImpostos " & _
            " WHERE desc_cd_desconto = " & CboDescontos.BoundText & "" & _
            " AND empr_cd_empresa = " & PCodEmpresa & ""
    SubQOpenRecordset XLO_DESCONTO, XLT_SQL, Estatico
    If Not XLO_DESCONTO.EOF Then
        If IsNull(XLO_DESCONTO!plco_cd_conta) Or XLO_DESCONTO!plco_cd_conta = "" Then
            MsgBox "O Desconto/Imposto não possui conta contábil.", vbCritical, "ATENÇÃO!"
            Exit Sub
        End If
    Else
        MsgBox "O Desconto/Imposto não possui conta contábil.", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    XLO_DESCONTO.Close
    '*************************************************************************
    
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
    If XDesconto = "I" Then
        ReDim XGM_MATRIZLOG(4, 2) As Variant
        
        'Dados do desconto informados pelo usuário
        XGM_MATRIZLOG(0, 1) = MskNumdocumento
        XGM_MATRIZLOG(1, 1) = CboDescontos.Text
        XGM_MATRIZLOG(2, 1) = Format(DtpVencDesconto, "dd/mm/yyyy")
        XGM_MATRIZLOG(3, 1) = Format$(TxtValorDesconto.Text, "##,##0.00")
        
    Else
        ReDim XGM_MATRIZLOG(4, 3) As Variant
        
        'Dados originais do desconto
        XGM_MATRIZLOG(0, 1) = MskNumdocumento
        XGM_MATRIZLOG(1, 1) = TDBGrid1.Columns("Descrição").Text
        XGM_MATRIZLOG(2, 1) = Format(TDBGrid1.Columns("Vencimento").Text, "dd/mm/yyyy")
        XGM_MATRIZLOG(3, 1) = Format$(TDBGrid1.Columns("valor").Text, "##,##0.00")
        
        'Novos dados do desconto
        XGM_MATRIZLOG(0, 2) = MskNumdocumento
        XGM_MATRIZLOG(1, 2) = CboDescontos.Text
        XGM_MATRIZLOG(2, 2) = Format(DtpVencDesconto, "dd/mm/yyyy")
        XGM_MATRIZLOG(3, 2) = Format$(TxtValorDesconto.Text, "##,##0.00")
    End If
    
    'Nome dos campos do desconto
    XGM_MATRIZLOG(0, 0) = "Nota Fiscal" 'Número de documento da nota fiscal
    XGM_MATRIZLOG(1, 0) = "Desconto" 'Nome do desconto
    XGM_MATRIZLOG(2, 0) = "Venc"     'Data de vencimento do desconto
    XGM_MATRIZLOG(3, 0) = "Valor"    'Valor do desconto
    '*****************************FIM LOG - PARTE 1********************************
       
    If (xcod = CboDescontos.BoundText) Then
        If XDesconto = "I" Then
           subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, "", "", ""), Vetor1, TDBGrid1
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
                subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, "", "", ""), Vetor1, TDBGrid1
            Else
                TDBGrid1.Columns("Descrição").Text = CboDescontos.Text
                TDBGrid1.Columns("Vencimento").Text = Format(DtpVencDesconto.Value, "dd/mm/yyyy")
                TDBGrid1.Columns("Valor").Text = TxtValorDesconto.Text
                TDBGrid1.Columns("codigo").Text = CboDescontos.BoundText
                TDBGrid1.Update
            End If
        End If
    End If
    
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
    'Gera descrição de log da operação
    If XDesconto = "I" Then
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DESCONTO, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DESCONTO, NOTAFISCAL_RECEBER)
    Else
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DESCONTO, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DESCONTO, NOTAFISCAL_RECEBER)
    End If
    
    'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
    XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
    XFT_CONTADORLOG = XFT_CONTADORLOG + 1
    '*****************************FIM LOG - PARTE 2********************************

    pandescontos.Visible = False
    pandescontos.Enabled = False
    
    If Vetor1(0, 0) <> Empty Then
       HabilitaBotoes
    Else
        CmdInsereDesconto.Enabled = True
    End If
    
    TDBGrid1.Enabled = True
   
    LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
    
    CmdExcluirTodos.Enabled = True

End Sub

Private Sub botao_conf_Click()
    Dim XLF_VALORNOTA As Double, XLF_VALORDISTRIB As Double
    Dim Xvalor
    Dim Mensagem As String
    Dim ResEstoque As Object

    'Função para verificar preenchimento dos campos obrigatórios
    If ChecarCampos Then Exit Sub
    
    'Inserido em 09/06/08 - Patrícia
    ' Verifica a Classificação "Empreemdimento\Obra" dos Centros de Custos e Contas Contábeis da Distribuição
     Mensagem = FunVerificaClassEmprObr(Str(DatPCR.Recordset.Fields("cere_cd_Pcr")), MskConta2.Text, PCodEmpresa)
     If Mensagem <> "" Then
        MsgBox Mensagem & "!", vbCritical, "AVISO"
        CboPlanoContas2.SetFocus
        Exit Sub
     End If
    '**** Fim 09/06/08 - Patrícia

    'Função para verificar preenchimento dos campos obrigatórios
    If ChecarCampos Then Exit Sub
    
    'Inserido em 06/04/2011 Andre Figueiredo
    'Verifica se o Centro de Custo não é analítico
    If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
        MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
        CboPcr2.BoundText = MskPcr2.Text
        CboPcr2.SetFocus
        Exit Sub
    End If
    
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
    If XDistribuicao = "I" Then
        ReDim XGM_MATRIZLOG(6, 2) As Variant
        
        'Dados do desconto informados pelo usuário
        XGM_MATRIZLOG(0, 1) = MskNumdocumento
        XGM_MATRIZLOG(1, 1) = CboPcr2.Text
        XGM_MATRIZLOG(2, 1) = CboPlanoContas2.Text
        XGM_MATRIZLOG(3, 1) = Format$(TxtPercDistribuicao.Text, "##,##0.00") & "%"
        XGM_MATRIZLOG(4, 1) = Format$(TxtVldistribuicao.Text, "##,##0.00")
        XGM_MATRIZLOG(5, 1) = CboIdentificador.Text
    Else
        ReDim XGM_MATRIZLOG(6, 3) As Variant
        
        'Dados originais da distribuição
        XGM_MATRIZLOG(0, 1) = MskNumdocumento
        XGM_MATRIZLOG(1, 1) = TDBGrid3.Columns("Centro Custo").Text
        XGM_MATRIZLOG(2, 1) = TDBGrid3.Columns("Conta Contábil").Text
        XGM_MATRIZLOG(3, 1) = TDBGrid3.Columns("Percentual")
        XGM_MATRIZLOG(4, 1) = Format$(TDBGrid3.Columns("Valor").Text, "##,##0.00")
        XGM_MATRIZLOG(5, 1) = TDBGrid3.Columns("Identificador").Text
        
        'Novos dados da distribuição informados pelo usuário
        XGM_MATRIZLOG(0, 2) = MskNumdocumento
        XGM_MATRIZLOG(1, 2) = MskPcr2.Text
        XGM_MATRIZLOG(2, 2) = MskConta2.Text
        XGM_MATRIZLOG(3, 2) = Format$(TxtPercDistribuicao.Text, "##,##0.00") & "%"
        XGM_MATRIZLOG(4, 2) = Format$(TxtVldistribuicao.Text, "##,##0.00")
        XGM_MATRIZLOG(5, 2) = CboIdentificador.Text
    End If
    
    'Nome dos campos da distribuição
    XGM_MATRIZLOG(0, 0) = "Nota Fiscal"
    XGM_MATRIZLOG(1, 0) = "Centro Custo"
    XGM_MATRIZLOG(2, 0) = "Conta Contábil"
    XGM_MATRIZLOG(3, 0) = "Percentual"
    XGM_MATRIZLOG(4, 0) = "Valor"
    If (CboIdentificador.Text <> "") Then XGM_MATRIZLOG(5, 0) = "Identificador"
    
    '*****************************FIM LOG - PARTE 1********************************
       
    If XDistribuicao = "A" Then
       ValorTotalDistribuicao = ValorTotalDistribuicao - CDbl(TxtVldistribuicao.Text)
    End If
    XLF_VALORNOTA = CDbl(LblValorNota.Caption)
    XLF_VALORDISTRIB = (ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text))
    If CDbl(Format(((CDbl(LblValorNota.Caption) * 100) / 100), "##,##0.00")) < CDbl(Format((((ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text)) * 100) / 100), "##,##0.00")) Then
        MsgBox "As distribuições excederam o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        TxtVldistribuicao.SetFocus
        Exit Sub
    Else
        If (XPcr = MskPcr2.Text) And (xcc = MskConta2.Text) Then
            If XDistribuicao = "I" Then
                subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, DatConta.Recordset.Fields("plco_cd_conta"), MskConta2.Text, TxtVldistribuicao.Text, 0, TxtVldistribuicao.Text, (TxtPercDistribuicao.Text / 100), CboIdentificador.BoundText, CboIdentificador.Text), Vetor3, TDBGrid3
            ElseIf XDistribuicao = "A" Then
                    TDBGrid3.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
                    TDBGrid3.Columns(1).Text = MskPcr2.Text
                    TDBGrid3.Columns(2).Text = DatConta.Recordset.Fields("plco_cd_conta")
                    TDBGrid3.Columns(3).Text = MskConta2.Text
                    TDBGrid3.Columns(4).Text = TxtVldistribuicao.Text
                    If XInserir = "A" Then
                        If EntSaida = "E" Then
                            TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Glosa").Text
                        Else
                            'Tratamento de variável nula ou vazia - Lucas Santiago - 27.01.09
                            If IsNull(TDBGrid3.Columns("Devolução").Text) = False And TDBGrid3.Columns("Devolução").Text <> "" Then
                                TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                            End If
                        End If
                    Else
                        TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text
                    End If
                    TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                    TDBGrid3.Columns(8).Text = CboIdentificador.BoundText
                    TDBGrid3.Columns(9).Text = CboIdentificador.Text
                    TDBGrid3.MoveFirst
                    TDBGrid3.Update
            End If
        Else
            If ExisteDistribuicao(2, Array(0, 2), Array(DatPCR.Recordset.Fields("cere_cd_pcr"), DatConta.Recordset.Fields("plco_cd_conta")), Vetor3, TDBGrid3) Then
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
                   subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, DatConta.Recordset.Fields("plco_cd_conta"), MskConta2.Text, TxtVldistribuicao.Text, 0, TxtVldistribuicao.Text, (TxtPercDistribuicao.Text / 100), CboIdentificador.BoundText, CboIdentificador.Text), Vetor3, TDBGrid3
                ElseIf XDistribuicao = "A" Then
                    TDBGrid3.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
                    TDBGrid3.Columns(1).Text = MskPcr2.Text
                    TDBGrid3.Columns(2).Text = DatConta.Recordset.Fields("plco_cd_conta")
                    TDBGrid3.Columns(3).Text = MskConta2.Text
                    TDBGrid3.Columns(4).Text = TxtVldistribuicao.Text
                    If EntSaida = "E" Then
                        TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Glosa").Text
                    Else
                        'Tratamento de variável nula ou vazia - Lucas Santiago - 27.01.09
                        If IsNull(TDBGrid3.Columns("Devolução").Text) = False And TDBGrid3.Columns("Devolução").Text <> "" Then
                            TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                        End If
                    End If
                    TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                    TDBGrid3.Columns(8).Text = CboIdentificador.BoundText
                    TDBGrid3.Columns(9).Text = CboIdentificador.Text
                    'TDBGrid3.SetFocus
                    TDBGrid3.Update
                End If
            End If
        End If
    End If
    
    
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
    'Gera descrição de log da operação
    If XDistribuicao = "I" Then
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
    Else
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
    End If
    
    'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
    XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
    XFT_CONTADORLOG = XFT_CONTADORLOG + 1
    '*****************************FIM LOG - PARTE 2********************************

    
    PanDistribuicao.Enabled = False
    PanDistribuicao.Visible = False
    If Vetor3(0, 0) <> Empty Then
       SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
       If (XFB_ESTOQUE = False) Or ((ResEstoque!empr_tx_AltNotasEstoque) = "S") Then
            'Não permite alterar notas do estoque, se não tiver selecionado o campo que
            'permite alterar em padrões de empresa - Lucas Santiago 23.01.09
            HabilitaBotoes
       Else
            CmdAlteraDistribuicao.Enabled = True
       End If
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
           XDistribuido = "S"
        Else
           XDistribuido = "N"
        End If
    End If
        
    '----------------------------------------------------------.
    
    MskPcr2.Mask = ""
    MskPcr2.Text = ""
    MskPcr2.Mask = pPCRmascara
    CboPcr2.Text = ""
    MskConta2.Mask = ""
    MskConta2.Text = ""
    CboPlanoContas2.Text = ""
    CmdGravar.Enabled = True
    
End Sub

Private Sub CmdConfirmaduplicata_Click()
    Dim XStatus As String
    Dim Xvalor As Double
    Dim xcod As String
    Dim XLI_ANO As Integer, XLI_MES As Integer, XLI_X As Integer
    Dim XLO_DESC As Object
    Dim XLT_DOCPAGREC As String 'Armazena o nº do DocPagRec
    Dim XLF_AUXILIAR1 As Double
    Dim XLF_AUXILIAR2 As Double
    Dim XLO_CHECAR As Object
   
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
      
    If EntSaida = "S" Then
        If CboStatusduplicata.ListIndex = 1 Then 'Realizado
            If Not FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "3") Then
                Beep
                MsgBox "Você não possui permissão para salvar um documento como Realizado.", 16, "AVISO"
                Exit Sub
            End If
        End If
    End If
    
    
    'AS VARIÁVEIS FORAM INTRODUZIDA PARA SOLUCIONAR UMA EXCEÇÃO REF. A APROX., ONDE OS 2 IFs ABAIXO NÃO RESOLVERAM
    XLF_AUXILIAR1 = Format(CDbl(LblValorLiquido1.Caption) + CDbl(TxtDevolucao.Text), "#.0000000000")
    XLF_AUXILIAR2 = Format(ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text), "#.0000000000")
    'A formatação abaixo foi implementada para tentar solucionar exceções na aproximação.
    'SÓ NÃO ATENDEU A EXCEÇÃO ACIMA->If (CDbl(LblValorLiquido1.Caption) + CDbl(TxtDevolucao.Text)) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
    'TESTE->If (CDbl(Format(LblValorLiquido1.Caption, "standard")) + CDbl(Format(TxtDevolucao.Text, "standard"))) < (CDbl(Format(ValorTotalDuplicata, "standard")) + CDbl(Format(TxtValorDuplicata.Text, "standard"))) Then
    If XLF_AUXILIAR1 < XLF_AUXILIAR2 Then
        MsgBox "A soma das duplicatas excedeu o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        TxtValorDuplicata.SetFocus
        CmdInsereduplicata_Click
    Else
        If TxtDuplicata.Text = "" Then
            xcod = "U"
        Else
            xcod = TxtDuplicata.Text
        End If
            
        XLT_DOCPAGREC = TxtDocPagto2.Text
        If CboStatusduplicata.ListIndex = 0 Then 'Compromissado
           XStatus = "C"
        ElseIf CboStatusduplicata.ListIndex = 1 Then 'Realizado
            XStatus = "R"
            
            If XLT_DOCPAGREC = "" Then 'Se Realizado sem nºDoc Pagto. gerar um automático
                XLT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
            End If
            
            '***************CONTABILIDADE*******************
            'possui o módulo de contabilidade e o tipo de documento exporta para a contabilidade
            If XGB_PQUALICONT And DatTipoDoc.Recordset.Fields("tido_tx_nec") = "0" Then
                If funMesFechado(DtpPagamento.Value) <> "A" Then
                    MsgBox "Não é possível cadastrar/alterar a duplicata com a data de Pagamento informada, pois o mês contábil já foi fechado. " _
                        + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    Exit Sub
                End If
            End If
            '***********************************************
        ElseIf CboStatusduplicata.ListIndex = 2 Then
            XStatus = "X"
        End If

        '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
        If XDuplicata = "I" Then 'Cadastro de duplicata
             If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(10, 2) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(9, 2) As Variant
                
                'Nome dos campos referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 0) = "Doc Pagto"
                XGM_MATRIZLOG(9, 0) = "Data Pagto"
                
                'Dados referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 2) = FunNulo(TxtDocPagto2.Text)
                XGM_MATRIZLOG(9, 2) = Format$(DtpPagamento.Value, "DD/MM/YYYY")
                
            Else
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(8, 2) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(7, 2) As Variant
            End If
            
            'Dados da duplicata
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            If LblValorLiquido1.Caption = TxtValorDuplicata.Text Then XGM_MATRIZLOG(1, 1) = "U"
            If LblValorLiquido1.Caption <> TxtValorDuplicata.Text Then XGM_MATRIZLOG(1, 1) = TxtDuplicata.Text
            XGM_MATRIZLOG(2, 1) = Format$(TxtValorDuplicata.Text, "##,##0.00")
            XGM_MATRIZLOG(3, 1) = CboStatusduplicata.List(CboStatusduplicata.ListIndex)
            XGM_MATRIZLOG(4, 1) = Format$(DtpVencimento.Value, "DD/MM/YYYY")
            XGM_MATRIZLOG(5, 1) = CboTipoPagto2.List(CboTipoPagto2.ListIndex)
            XGM_MATRIZLOG(6, 1) = CboCCorrente2.Text
            If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 1) = CboBancoForn2.Text
                        
        Else 'Alteração de duplicata
            If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(10, 3) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(9, 3) As Variant
                
                'Nome dos campos referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 0) = "Doc Pagto"
                XGM_MATRIZLOG(9, 0) = "Data Pagto"
                
                'Dados originais referentes ao pagamento da duplicata
                If EntSaida = "S" Then XGM_MATRIZLOG(8, 1) = TDBGrid2.Columns("Doc. Pag.").Text
                If EntSaida = "E" Then XGM_MATRIZLOG(8, 1) = TDBGrid2.Columns("Doc. Rec.").Text
                If EntSaida = "S" Then XGM_MATRIZLOG(9, 1) = Format$(TDBGrid2.Columns("Data Pag.").Text, "DD/MM/YYYY")
                If EntSaida = "E" Then XGM_MATRIZLOG(9, 1) = Format$(TDBGrid2.Columns("Data Rec.").Text, "DD/MM/YYYY")
                
                'Novos dados referentes ao pagamento da duplicata
                XGM_MATRIZLOG(8, 2) = FunNulo(TxtDocPagto2.Text)
                XGM_MATRIZLOG(9, 2) = Format$(DtpPagamento.Value, "DD/MM/YYYY")
                
            Else
                If (CboBancoForn2.Enabled = True) Then ReDim XGM_MATRIZLOG(8, 3) As Variant
                If (CboBancoForn2.Enabled = False) Then ReDim XGM_MATRIZLOG(7, 3) As Variant
            End If
            
            'Dados originais da duplicata
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 1) = TDBGrid2.Columns("Código").Text
            XGM_MATRIZLOG(2, 1) = Format$(TDBGrid2.Columns("valor").Text, "##,##0.00")
            XGM_MATRIZLOG(3, 1) = TDBGrid2.Columns("status").Text
            XGM_MATRIZLOG(4, 1) = Format$(TDBGrid2.Columns("vencimento").Text, "DD/MM/YYYY")
            If EntSaida = "S" Then XGM_MATRIZLOG(5, 1) = TDBGrid2.Columns("Tipo Pag.").Text
            If EntSaida = "E" Then XGM_MATRIZLOG(5, 1) = TDBGrid2.Columns("Tipo Rec.").Text
            XGM_MATRIZLOG(6, 1) = TDBGrid2.Columns("Conta Corrente (Empresa)").Text
            If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 1) = TDBGrid2.Columns("Conta Corrente (Fornecedor)").Text
                    
            'Novos dados da duplicata (alterados pelo usuário)
            XGM_MATRIZLOG(0, 2) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 2) = TxtDuplicata.Text
            XGM_MATRIZLOG(2, 2) = Format$(TxtValorDuplicata.Text, "##,##0.00")
            XGM_MATRIZLOG(3, 2) = CboStatusduplicata.List(CboStatusduplicata.ListIndex)
            XGM_MATRIZLOG(4, 2) = Format$(DtpVencimento.Value, "DD/MM/YYYY")
            XGM_MATRIZLOG(5, 2) = CboTipoPagto2.List(CboTipoPagto2.ListIndex)
            XGM_MATRIZLOG(6, 2) = CboCCorrente2.Text
            If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 2) = CboBancoForn2.Text
            
        End If
       
        'Nome dos campos da duplicata
        XGM_MATRIZLOG(0, 0) = "Nota Fiscal" 'Número do documento da nota fiscal
        XGM_MATRIZLOG(1, 0) = "Cod"
        XGM_MATRIZLOG(2, 0) = "Valor"
        XGM_MATRIZLOG(3, 0) = "Status"
        XGM_MATRIZLOG(4, 0) = "Data Venc"
        XGM_MATRIZLOG(5, 0) = "Tipo Pagto"
        XGM_MATRIZLOG(6, 0) = "Conta Corrente"
        If (CboBancoForn2.Enabled = True) Then XGM_MATRIZLOG(7, 0) = "Banco Fornecedor"
        
        '*****************************FIM LOG - PARTE 1********************************
                
        If XDuplicata = "I" Then
            If CboStatusduplicata.Text = "Realizado" Then
                'Checar se o nº do doc. pagto. já não está associado a um pagto nesta data e para esta conta
                If EntSaida = "S" Then ' Despesa
                    SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex) & "' and dupl_tx_docpagrec ='" & XLT_DOCPAGREC & "' and dupl_dt_pagrec = " & FunNuloData(DtpPagamento.Value, NomeSgbd) & "", "Estatico"
                    If Not XLO_CHECAR.EOF Then
                        MsgBox "Este nº de Doc. de Pagto. já foi emitido!", vbCritical, "ATENÇÃO!"
                        Exit Sub
                    End If
                    XLO_CHECAR.Close
                Else
                    SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='C' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex) & "' and dupl_tx_docpagrec ='" & XLT_DOCPAGREC & "'  and dupl_dt_pagrec = " & FunNuloData(DtpPagamento.Value, NomeSgbd) & "", "Estatico"
                    If Not XLO_CHECAR.EOF Then
                        MsgBox "Este nº doc. de crédito já foi registrado!", vbCritical, "ATENÇÃO!"
                        Exit Sub
                    End If
                    XLO_CHECAR.Close
                End If
            End If
                                                
            InsereRegistroVetorDupl Array(xcod, Format(DtpVencimento.Value, "dd/mm/yyyy"), TxtValorDuplicata.Text, 0, TxtValorDuplicata.Text, XStatus, CboStatusduplicata.Text, _
            CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex), CboTipoPagto2.Text, XLT_DOCPAGREC, Format(DtpPagamento.Value, "dd/mm/yyyy"), CboCCorrente2.BoundText, CboCCorrente2.Text, FunNuloValor(CboBancoForn2), CboBancoForn2.Text, CboFormaPagamento2.BoundText, CboFormaPagamento2.Text), Vetor2, TDBGrid2
            CmdDuplMultiplas.Enabled = True
        
        ElseIf XDuplicata = "A" Then
            If TDBGrid2.Columns("Status").Text = "Compromissado" And CboStatusduplicata.Text = "Realizado" Then
                'Checar se o nº do doc. pagto. já não está associado a um pagato nesta data e para esta conta
                If EntSaida = "S" Then ' Despesa
                    SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex) & "' and dupl_tx_docpagrec ='" & XLT_DOCPAGREC & "' and dupl_dt_pagrec = " & FunNuloData(DtpPagamento.Value, NomeSgbd) & "", "Estatico"
                    If Not XLO_CHECAR.EOF Then
                        MsgBox "Este nº de Doc. de Pagto. já foi emitido!", vbCritical, "ATENÇÃO!"
                        Exit Sub
                    End If
                    XLO_CHECAR.Close
                Else
                    SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='C' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex) & "' and dupl_tx_docpagrec ='" & XLT_DOCPAGREC & "' and dupl_dt_pagrec = " & FunNuloData(DtpPagamento.Value, NomeSgbd) & "", "Estatico"
                    If Not XLO_CHECAR.EOF Then
                        MsgBox "Este nº doc. de crédito já foi registrado!", vbCritical, "ATENÇÃO!"
                        Exit Sub
                    End If
                    XLO_CHECAR.Close
                End If
            End If
            
            TDBGrid2.Columns("Código").Text = TxtDuplicata.Text
            TDBGrid2.Columns("Vencimento").Text = Format(DtpVencimento.Value, "dd/mm/yyyy")
            TDBGrid2.Columns("Valor").Text = TxtValorDuplicata.Text
            TDBGrid2.Columns("CodStatus").Text = XStatus
            TDBGrid2.Columns("Status").Text = CboStatusduplicata.Text
            TDBGrid2.Columns("CodTipoPagto").Text = CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex)
            TDBGrid2.Columns("Valor Líquido").Text = TxtValorDuplicata.Text - TDBGrid2.Columns(3).Text
            TDBGrid2.Columns("CodConta").Text = CboCCorrente2.BoundText
            TDBGrid2.Columns("Conta Corrente (Empresa)").Text = CboCCorrente2.Text
            
            If EntSaida = "S" Then
                TDBGrid2.Columns("Tipo Pag.").Text = CboTipoPagto2.Text
                TDBGrid2.Columns("Doc. Pag.").Text = XLT_DOCPAGREC
                TDBGrid2.Columns("Data Pag.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
                TDBGrid2.Columns("CodBancoForn").Text = FunNuloValor(CboBancoForn2)
                TDBGrid2.Columns("Conta Corrente (Fornecedor)").Text = CboBancoForn2.Text
                TDBGrid2.Columns("Cód Forma Pagto").Text = CboFormaPagamento2.BoundText
                TDBGrid2.Columns("Forma Pagto").Text = CboFormaPagamento2.Text
            Else
                TDBGrid2.Columns("Tipo Rec.").Text = CboTipoPagto2.Text
                TDBGrid2.Columns("Doc. Rec.").Text = XLT_DOCPAGREC
                TDBGrid2.Columns("Data Rec.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
            End If
                        
            TDBGrid2.Update
                        
        End If
                
        '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
        'Gera descrição de log da operação
        If XDuplicata = "I" Then
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DUPLICATA, NOTAFISCAL_RECEBER)
        Else
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DUPLICATA, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR_DUPLICATA, NOTAFISCAL_RECEBER)
        End If

        'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
        XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
        XFT_CONTADORLOG = XFT_CONTADORLOG + 1
        '*****************************FIM LOG - PARTE 2********************************
    
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
            If Vetor1(XLI_X, 1) = Empty And Vetor1(XLI_X, 3) <> Empty Then
               SubQOpenRecordset XLO_DESC, "SELECT * FROM DescontosImpostos WHERE desc_cd_desconto=" & Vetor1(XLI_X, 3), "Estatico"
               If XLO_DESC!desc_tx_tipobase = "1" Or XLO_DESC!desc_tx_tipobase = "3" Then
                   XLI_ANO = Year(CDate(Vetor2(0, 1)))
                   XLI_MES = Month(CDate(Vetor2(0, 1)))
                   If XLI_MES = 12 Then
                       XLI_MES = 1
                   Else
                       XLI_MES = XLI_MES + 1
                   End If
                   TDBGrid1.Columns("Vencimento").Text = DateSerial(XLI_ANO, XLI_MES, XLO_DESC!desc_tx_DiaRecolhimento)
               Else
                   TDBGrid1.Columns("Vencimento").Text = Vetor2(0, 1)
               End If
               XLO_DESC.Close
            End If
            TDBGrid1.MoveNext
        Next
    End If
    
    '***************CONTABILIDADE*******************
    If Not IsNull(DtpPagamento) Then
        If Month(DtpPagamento.Value) = Month(DtpEntrada.Value) Then
            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                If (DtpPagamento.Value - DtpEntrada.Value >= XGI_QTDDIASPROVISAO) Then
                    XFB_PROVISIONA = True
                Else
                    XFB_PROVISIONA = False
                End If
            Else
                XFB_PROVISIONA = False
            End If
        Else
            XFB_PROVISIONA = True
        End If
    Else
        If Month(DtpVencimento.Value) = Month(DtpEntrada.Value) Then
            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                If ((DtpVencimento.Value - DtpEntrada.Value) >= XGI_QTDDIASPROVISAO) Then
                    XFB_PROVISIONA = True
                Else
                    XFB_PROVISIONA = False
                End If
            Else
                XFB_PROVISIONA = False
            End If
        Else
            XFB_PROVISIONA = True
        End If
    End If
    '***********************************************
ResumoDoErro:
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Resume ResumoDoErro
End Sub

Private Sub CmdConfirmar_Click()
    Dim XLF_VALORNOTA As Double
    Dim XLF_VALORDISTRIB As Double
    Dim Xvalor
    Dim Mensagem As String
    Dim ResEstoque As Object
    Dim Xau As Double
    ValorTotalDistribuicao = 0
    
    If CboCriterios.Text = "" Then
        MsgBox "Escolha um critério.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    XDistribuicao = "I"
    TDBGrid10.MoveFirst
    While Not TDBGrid10.EOF
    
        'Verifica a Classificação "Empreemdimento\Obra" dos Centros de Custos e Contas Contábeis da Distribuição
         Mensagem = FunVerificaClassEmprObr(Str(Adodc1.Recordset.Fields("cere_cd_Pcr")), Adodc1.Recordset.Fields("plco_cd_CodReduzido"), PCodEmpresa)
         If Mensagem <> "" Then
            MsgBox Mensagem & "!", vbCritical, "AVISO"
            CboPlanoContas2.SetFocus
            Exit Sub
         End If
         
        
        XLF_VALORNOTA = CDbl(LblValorNota.Caption)
        XLF_VALORDISTRIB = (ValorTotalDistribuicao + CDbl(LblValorNota.Caption) * (TDBGrid10.Columns(4) / 100))
        ValorTotalDistribuicao = ValorTotalDistribuicao + CDbl(LblValorNota.Caption) * (TDBGrid10.Columns(4) / 100)
            
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
        If XDistribuicao = "I" Then
            ReDim XGM_MATRIZLOG(6, 2) As Variant
            
            'Dados do desconto informados pelo usuário
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("cere_cd_pcr")
            XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("plco_cd_CodReduzido")
            XGM_MATRIZLOG(3, 1) = Format$(TDBGrid10.Columns(4), "##,##0.00") & "%"
            XGM_MATRIZLOG(4, 1) = Format$(Xau, "##,##0.00")
            XGM_MATRIZLOG(5, 1) = CboIdentificador.Text
        End If
        
        'Nome dos campos da distribuição
        XGM_MATRIZLOG(0, 0) = "Nota Fiscal"
        XGM_MATRIZLOG(1, 0) = "Centro Custo"
        XGM_MATRIZLOG(2, 0) = "Conta Contábil"
        XGM_MATRIZLOG(3, 0) = "Percentual"
        XGM_MATRIZLOG(4, 0) = "Valor"
        If (CboIdentificador.Text <> "") Then XGM_MATRIZLOG(5, 0) = "Identificador"
        
        '*****************************FIM LOG - PARTE 1********************************
    
        If XDistribuicao = "I" Then
            subInsereRegistroVetor Array(Adodc1.Recordset.Fields("cere_cd_pcr"), TDBGrid10.Columns(0).Value, Adodc1.Recordset.Fields("plco_cd_conta"), TDBGrid10.Columns(2), CDbl(LblValorNota.Caption) * (TDBGrid10.Columns(4) / 100), 0, CDbl(LblValorNota.Caption) * (TDBGrid10.Columns(4) / 100), (TDBGrid10.Columns(4).Value / 100), CboIdentificador.BoundText, CboIdentificador.Text), Vetor3, TDBGrid3
        End If
        
            
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
        'Gera descrição de log da operação
        If XDistribuicao = "I" Then
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
        End If
        
        'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
        XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
        XFT_CONTADORLOG = XFT_CONTADORLOG + 1
        '*****************************FIM LOG - PARTE 2********************************
        
        If XInserir = "I" Or XDistribuido = "N" Then
            If LblValorNota.Caption = CDbl(TxtValor.Text) Then
                XDistribuido = "S"
            Else
                XDistribuido = "N"
            End If
        End If
        
        
        TDBGrid10.MoveNext
    Wend
    
    LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
    
    PainelCriterios.Enabled = False
    PainelCriterios.Visible = False
        
    CmdInsereDistribuicao.Enabled = False
    CmdGravar.Enabled = True

End Sub

Private Sub CmdCriterio_Click()
    If TabNotasFiscais.Tab = 3 Then
        FrmNotasFiscais.MousePointer = vbHourglass
'        PainelCriterios.Visible = True
'        PainelCriterios.Enabled = True
        FrmNotasFiscais.MousePointer = vbDefault
    End If
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

Private Sub CmdDuplMultiplas_Click()
    If TabNotasFiscais.Tab = 1 Then
'        PanEntMultiplas.Visible = True
'        PanEntMultiplas.Enabled = True
        If XFT_INSDUPLMULTIPLA = "I" Then
            TxtNParcelas.Text = ""
            TxtNParcelas.SetFocus
            TxtPeriodicidade.Text = ""
            CmdDesiste.Visible = False
        Else
            CmdDesiste.Visible = True
        End If
        TabNotasFiscais.Enabled = False
        XFT_NOTADUPLICATA = "D" ' Gerar várias duplicatas
    End If
End Sub

Private Sub CmdEntMultiplas_Click()
    If ChecarCamposOrelhaPrincipal Then Exit Sub
'    PanEntMultiplas.Visible = True
'    PanEntMultiplas.Enabled = True
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
    
    If TabNotasFiscais.Tab = 2 Then
        
        If TDBGrid1.Columns("NumDarf").Text <> "" Then
            MsgBox ("Este imposto já foi recolhido portanto não poderá ser excluído!"), vbExclamation, " Informação"
            Exit Sub
        End If
        
        Houve_Erro = False
            
        DesabilitaBotoes
        TDBGrid1.Enabled = False
        If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
            
            '*********REGISTRA LOG DA OPERAÇÃO - PARTE 1**********
            ReDim XGM_MATRIZLOG(4, 2) As Variant
        
            XGM_MATRIZLOG(0, 0) = "Nota Fiscal" 'Nome do desconto
            XGM_MATRIZLOG(1, 0) = "Desconto" 'Nome do desconto
            XGM_MATRIZLOG(2, 0) = "Venc"     'Data de vencimento do desconto
            XGM_MATRIZLOG(3, 0) = "Valor"    'Valor do desconto
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 1) = TDBGrid1.Columns("Descrição").Text
            XGM_MATRIZLOG(2, 1) = Format(TDBGrid1.Columns("Vencimento").Text, "dd/mm/yyyy")
            XGM_MATRIZLOG(3, 1) = Format$(TDBGrid1.Columns("valor").Text, "##,##0.00")
            '****************FIM LOG - PARTE 1********************
            
           subRemoveRegistroVetor Vetor1, TDBGrid1
           
           
           '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
            'Gera descrição de log da operação
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DESCONTO, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DESCONTO, NOTAFISCAL_RECEBER)
                   
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '*****************************FIM LOG - PARTE 2********************************
           
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
    End If
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub CmdExcluirTodos_Click()

Dim Houve_Erro As Integer
Dim x As Integer

    If TabNotasFiscais.Tab = 2 Then
    
        DesabilitaBotoes
        TDBGrid1.Enabled = False
        If MsgBox("Confirma Remoção de Todos os Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        
            Houve_Erro = False
            
            For x = 0 To TDBGrid1.ApproxCount - 1
                'TDBGrid1.Delete
                subRemoveRegistroVetor Vetor1, TDBGrid1
            Next
        
        HabilitaBotoes
        CmdExcluirTodos.Enabled = False
        
        End If
        
    End If
            
End Sub

Private Sub CmdExcluiduplicata_Click()
    Dim Houve_Erro As Integer
    Dim ValorExcluido As Double

    If TabNotasFiscais.Tab = 1 Then
        Houve_Erro = False
         
        DesabilitaBotoes
        
        TDBGrid2.Enabled = False
        
        If XInserir = "A" And TDBGrid2.Columns("Status") = "Realizado" Then
             MsgBox "A duplicata não pode ser apagada porque já foi realizada.", vbCritical + vbOKOnly, "ATENÇÃO"
             HabilitaBotoes
             Exit Sub
        End If
        
        If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
             
            '*********REGISTRA LOG DA OPERAÇÃO - PARTE 1**********
            ReDim XGM_MATRIZLOG(3, 2) As Variant
            
            XGM_MATRIZLOG(0, 0) = "Nota Fiscal"
            XGM_MATRIZLOG(1, 0) = "Código"
            XGM_MATRIZLOG(2, 0) = "Valor"
            XGM_MATRIZLOG(0, 1) = MskNumdocumento.Text
            XGM_MATRIZLOG(1, 1) = TDBGrid2.Columns("Código").Text
            XGM_MATRIZLOG(2, 1) = Format$(TDBGrid2.Columns("valor").Text, "##,##0.00")
            '****************FIM LOG - PARTE 1********************
             
             ValorExcluido = TDBGrid2.Columns("Valor").Value
             subRemoveRegistroVetor Vetor2, TDBGrid2
             
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
            'Gera descrição de log da operação
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DUPLICATA, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DUPLICATA, NOTAFISCAL_RECEBER)
                   
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '*****************************FIM LOG - PARTE 2********************************
             
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
    End If

RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub


Public Sub Grava()
    
    Dim Sql As String
    Dim XStatus As String
    Dim XSerie As String
    Dim ResNF As Object
    Dim x As Byte
    Dim XNumDoc As String
    Dim XCodDupl As String
    Dim ResFat As Object
    Dim ResFormulario As Object
    Dim ResEstoque As Object
    Dim ResEmpObr As Object 'Declaração em 05/06/08 - Patrícia
    Dim Desc_CREmpObr As String 'Declaração em 05/06/08 - Patrícia
    Dim Desc_CCEmpObr As String 'Declaração em 05/06/08 - Patrícia
    Dim DataVencimento As Date
    Dim XLT_DOCPAGREC As String
    Dim ChaveTemp As Variant
    Dim XLT_INSERIRTEMP As String
    Dim XLB_REGISTROAFETADO As Boolean
    Dim XLT_SQL As String
    Dim RESFAT2 As Object
    Dim XLD_DATAENTRADA As Date
    Dim XLT_TIPODOC As Object 'Criado para verificar se o tipo de documento é prestação de contas - Larissa Moura 02/06/2008
    Dim teste As String
    Dim TelaAnterior As Integer ' Declarado em 09/06/08 - Patrícia
    Dim Cont As Integer 'Larissa 30/06/2008
    Dim ResVencDuplicatas As New ADODB.Recordset 'Patrícia 05/08/2010
    Dim XLD_DATAVENCDUPL As Variant 'Patrícia 05/08/2010
    Dim XLD_DATAPAGDUPL As Variant 'Patrícia 05/08/2010
    Dim XLO_RS As ADODB.Recordset
    Dim XLB_TIPODOCINATIVO As Boolean
    If EntSaida = "S" Then
        If (CboStatus.Text = "Realizado") Then
            If Not FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "3") Then
                Beep
                MsgBox "Você não possui permissão para salvar um documento como Realizado.", 16, "AVISO"
                Exit Sub
            End If
        End If
    End If
    
    
    XLB_TIPODOCINATIVO = False
    
    SubQOpenRecordset XLO_RS, "SELECT * FROM Tipodocumentos WHERE tido_bl_ativo = 0 AND tido_cd_TipoDoc = " & DatTipoDoc.Recordset("tido_cd_TipoDoc").Value, Estatico
    While Not XLO_RS.EOF And Not XLO_RS.BOF
        XLB_TIPODOCINATIVO = True
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
    Set XLO_RS = Nothing
    
    If (XLB_TIPODOCINATIVO) Then
        MsgBox "O tipo de documento escolhido está inativo. Escolha outro.", vbInformation, "ATENCAO"
        Exit Sub
    End If
    
    If XInserir = "A" And Tipo = "0" Then
    SubQOpenRecordset ResVencDuplicatas, "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura & "order by dupl_dt_vencimento", Dinamico
    'Pega a data de vencimento da primeira duplicata
    If Not ResVencDuplicatas.EOF Then
        XLD_DATAVENCDUPL = ResVencDuplicatas!dupl_dt_vencimento
        XLD_DATAPAGDUPL = ResVencDuplicatas!dupl_dt_pagrec
    End If
    ResVencDuplicatas.Close
    Set ResVencDuplicatas = Nothing
    End If
    
    'Larissa, 28/01/09
    'Verifica se o usuário cadastrou distribuição
    If ChkDistribuicao.Value = 1 And EntrouDist = True And (LblValorTotalDistrib = "" Or LblValorTotalDistrib = "0,00") Then
        MsgBox "É necessário cadastrar a distribuição da nota!", vbInformation, "ATENCAO"
        Exit Sub
    End If
    
    'Inserido em 09/06/08 - Patrícia
    ' Verifica a Classificação "Empreendimento\Obra" dos Centros de Custos e Contas Contábeis da Distribuição
    If (DtpEmissao) >= CDate(Data_EmpObr) Then
                If ChkDistribuicao.Value = 0 Then 'Verifica se houve distribuição
                    XMENSAGEM = FunVerificaClassEmprObr(Str(DatPCR.Recordset.Fields("cere_cd_Pcr")), MskConta.Text, PCodEmpresa)
                     If XMENSAGEM <> "" Then
                        MsgBox XMENSAGEM & "!", vbCritical, "AVISO"
                        CboPlanoContas.SetFocus
                        Exit Sub
                     End If
                Else
                    'Se houver distribuição, faz a verificação da classificação "Empreendimento\Obra"
                     TDBGrid3.MoveFirst
                     Do While Not TDBGrid3.EOF
                          XMENSAGEM = FunVerificaClassEmprObr(TDBGrid3.Columns("CodPcr").Value, TDBGrid3.Columns("Conta Contábil").Value, PCodEmpresa)
                          If XMENSAGEM <> "" Then
                             'Mostra a Tela de Distribuição
                             MsgBox XMENSAGEM & "!", vbCritical, "AVISO"
                             TelaAnterior = TabNotasFiscais.Tab
                             TabNotasFiscais.Tab = 3
                             TabNotasFiscais_Click (TelaAnterior)
                             Exit Sub
                          End If
                          TDBGrid3.MoveNext
                      Loop
               
                End If
    End If
    'Fim 09/06/08 - Patrícia
    
    Reload = True
    
    CmdGravar.Enabled = False
    
    Me.MousePointer = vbHourglass
    
    TabNotasFiscais.Tab = 0
    If ChecarCamposOrelhaPrincipal Then
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    'verificar se possui duplicatas canceladas para essa nota
    If FunVerificaDuplicataCancelada = False Then
        MsgBox "Notas Fiscais com duplicatas canceladas não podem ser alteradas!", vbCritical, "AVISO"
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    
    '/////////////////////////////////////////////////////////////////////////////////
    'SE A NOTA É DO ESTOQUE O QUALIFIN SÓ PODE ALTERAR O IDENTIFICADOR E SAI DA ROTINA
    SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
    If (XFB_ESTOQUE = True) And (ResEstoque!empr_tx_AltNotasEstoque = "N") Then
        'Não permite alterar notas do estoque, se não tiver selecionado o campo que
        'permite alterar em padrões de empresa - Lucas Santiago 23.01.09
        If ChkDistribuicao.Value = 1 Then
            If Vetor3(0, 0) <> Empty Then
                For x = 0 To Vetor3.UpperBound(1)
                   Conexao.Execute ("UPDATE Distribuicao SET iden_cd_Projeto = " & FunNuloBancoVal(Vetor3(x, 8)) & "" & _
                                    " WHERE nofi_cd_NotaFiscal = " & Chave & " AND cere_cd_Pcr = '" & Vetor3(x, 0) & "'" & _
                                    " AND plco_cd_conta = '" & Vetor3(x, 2) & "'")
                Next
            End If
        Else
            Conexao.Execute ("UPDATE Distribuicao SET iden_cd_Projeto = " & FunNuloBancoVal(CboIdentificador2.BoundText) & "" & _
                                " WHERE nofi_cd_NotaFiscal = " & Chave & "")
        End If
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Unload Me
        Exit Sub
    End If
    
    '//////////////////////////////////////////////////////////////////////////////////
    '***************CONTABILIDADE - COMPOSIÇÃO DE DUPLICATA - QUANDO DA ALTERAÇÃO ***********
    'Se o usuário abrisse esse form para consultar a nota e gravasse sem confirma as duplicatas
    'o valor da parcela era lançado na contabilidade. Deve-se então, varrer o vertor para confirmar a contabilização
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If XInserir = "A" And Tipo = "0" Then 'Só para cadastros na orelha de composição de duplicatas
            For x = 0 To Vetor2.UpperBound(1)
                If XFB_PROVISIONA = False Then 'se já foi definido para provisionar não precisa continuar a pesquisa
                    If IsDate(Vetor2(x, 10)) Then
                        If Month(CDate(Vetor2(x, 10))) = Month(DtpEntrada.Value) Then
                            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                                If ((CDate(Vetor2(x, 10)) - DtpEntrada.Value) >= XGI_QTDDIASPROVISAO) Then
                                    XFB_PROVISIONA = True
                                Else
                                    XFB_PROVISIONA = False
                                    'Verificar se o mês está fechado com base na data de pagto das parcelas
                                    If XGB_PQUALICONT Then  'possui o módulo de contabilidade
                                        If funMesFechado(Vetor2(x, 10)) <> "A" Then
                                            MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                                                + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Else
                                XFB_PROVISIONA = False
                                'Verificar se o mês está fechado com base na data de pagto das parcelas
                                If XGB_PQUALICONT Then  'possui o módulo de contabilidade
                                    If funMesFechado(Vetor2(x, 10)) <> "A" Then
                                        MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                                            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                                        Me.MousePointer = vbDefault
                                        CmdGravar.Enabled = True
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            XFB_PROVISIONA = True
                        End If
                    Else
                        If Month(CDate(Vetor2(x, 1))) = Month(DtpEntrada.Value) Then
                            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                                If ((CDate(Vetor2(x, 1)) - DtpEntrada.Value) >= XGI_QTDDIASPROVISAO) Then
                                    XFB_PROVISIONA = True
                                Else
                                End If
                            Else
                                XFB_PROVISIONA = False
                            End If
                        Else
                            XFB_PROVISIONA = True
                        End If
                    End If
                Else
                    x = Vetor2.UpperBound(1)
                End If
            Next
        End If
    End If
    
    '****************************************************************************************
    'Em vista do que foi acordado para a realização da manutenção permanente do sistema,informo que a maneira pela qual o sistema
    '***************CONTABILIDADE - REALIZADO*******************
    'possui o módulo de contabilidade
    If XGB_PQUALICONT And DatTipoDoc.Recordset.Fields("tido_tx_nec") = "0" Then
        'Só entra na checagem os que estiverem sendo realizados através da alteração
        If Not IsNull(DtpPagtoParcelaUnica) Then
            If funMesFechado(DtpPagtoParcelaUnica.Value) <> "A" Then
                MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                CmdGravar.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    '**********************************************************
    If Tipo = "1" Or CboStatus.Text = "Previsto" Then 'Gera duplicata automaticamente
        XCodDupl = "U" 'Parcela única
        XFB_PROVISIONA = funVerificaProvisao(DtpPagtoParcelaUnica, DtpEntrada, DtpVencParcelaUnica, XGI_QTDDIASPROVISAO, CboStatus.Text)
    End If
    
    '***************CONTABILIDADE*******************
    'Comp ou Realizado, possui o módulo de contabilidade, exporta base provisão, e se pela data é preciso provisionar
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" And XFB_PROVISIONA Then
            'A variável XLB_PROVISIONA foi incluida p/ verificar se (DtPagRec-DtEntrada) >= XGI_QTDDIASPROVISAO
            If XFB_PROVISIONA Then
                If funMesFechado(DtpEntrada.Value) <> "A" Then
                    MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                        + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                Else
                    If Tipo = 1 Then ' Se gera duplicata automaticamente entao Tipo =1
                        XFB_PROVISIONA = funVerificaProvisao(IIf(IsNull(DtpPagtoParcelaUnica.Value), "", DtpPagtoParcelaUnica.Value), IIf(IsNull(DtpEntrada.Value), "", DtpEntrada.Value), IIf(IsNull(DtpVencParcelaUnica.Value), "", DtpVencParcelaUnica.Value), XGI_QTDDIASPROVISAO, CboStatus.Text)
'                    Else
'                        XFB_PROVISIONA = funVerificaProvisao(IIf(IsNull(XLD_DATAPAGDUPL), "", XLD_DATAPAGDUPL), IIf(IsNull(DtpEntrada.Value), "", DtpEntrada.Value), IIf(IsNull(XLD_DATAVENCDUPL), "", XLD_DATAVENCDUPL), XGI_QTDDIASPROVISAO, CboStatus.Text)
                    End If
                    
                    If XFB_PROVISIONA = True Then
                            DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
                            'Se for adiantamento, não precisa checar a conta de adiantamento
                            If DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") <> "1" Then
                                While IsNull(DatFornecedor.Recordset.Fields("plco_cd_ContaCusto")) Or DatFornecedor.Recordset.Fields("plco_cd_ContaCusto") = ""
                                    'mas q falta de vontade de trabalhar viu!!!! essa semana demorou pra porra de passar
                                    ChaveTemp = Chave
                                    If EntSaida = "S" Then
                                        If MsgBox("O fornecedor não possui Conta Contábil para provisão." _
                                            + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro da nota.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            GoTo FimGrava
                                        End If
                                    Else
                                        If MsgBox("O Cliente não possui Conta Contábil para provisão." _
                                            + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro da nota.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            Exit Sub
                                        End If
                                    End If
                                    Pessoa = DatFornecedor.Recordset.Fields("focl_tx_Tipo")
                                    FornClie = DatFornecedor.Recordset.Fields("focl_tx_classe")
                                    
                                    XFormForn = "FrmNotasFiscais"
                                    XLT_INSERIRTEMP = XInserir
                                    XInserir = "A"
                                    Chave = DatFornecedor.Recordset.Fields("focl_cd_forncli")
                                    FrmFornecedores.TxtCgc = DatFornecedor.Recordset.Fields("focl_tx_cgccpf")
                                    FrmFornecedores.Show 1
                                    Chave = ChaveTemp
                                    XInserir = XLT_INSERIRTEMP
                                    FrmNotasFiscais.DatFornecedor.Refresh
                                    FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
                                    DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
                                Wend
                            End If
                    End If
                    
                End If
            End If
        End If
    End If
    
    DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
    
    'Verifica para os casos de adiantamento, se o fornecedor possui conta de adiantamento cadastrada
    If DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") = "1" Then
        If XInserir = "A" Then
            'Foi necessário incluir a exclusão da provisão por que o usuário
            'pode estar alterando uma nota que foi provisionada para um adiantamento
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                            " (lanc_tx_Origem = 'FN') AND " & _
                            "empr_cd_empresa = " & PCodEmpresa
        End If
        
        While IsNull(DatFornecedor.Recordset.Fields("plco_cd_ContaProvAdiant")) Or DatFornecedor.Recordset.Fields("plco_cd_ContaProvAdiant") = ""
            ChaveTemp = Chave
            If EntSaida = "S" Then
                If MsgBox("O fornecedor não possui Conta Contábil para Adiantamento." _
                    + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro do Adiantamento.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                End If
            Else
                If MsgBox("O Cliente não possui Conta Contábil para Adiantamento." _
                    + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro do Adiantamento.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                End If
            End If
            Pessoa = DatFornecedor.Recordset.Fields("focl_tx_Tipo")
            FornClie = DatFornecedor.Recordset.Fields("focl_tx_classe")
            
            XFormForn = "FrmNotasFiscais"
            XLT_INSERIRTEMP = XInserir
            XInserir = "A"
            Chave = DatFornecedor.Recordset.Fields("focl_cd_forncli")
            FrmFornecedores.TxtCgc = DatFornecedor.Recordset.Fields("focl_tx_cgccpf")
            FrmFornecedores.Show 1
            Chave = ChaveTemp
            XInserir = XLT_INSERIRTEMP
            FrmNotasFiscais.DatFornecedor.Refresh
            FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
            DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
        Wend
    End If
    '***********************************************
            
    If XDistribuido = "N" And (Trim(MskPcr.Text) = "" And MskConta.Text = "") Then 'Alterado em 31/07/08 - Larissa
        MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    ElseIf ChkDistribuicao.Value = 1 Then
        If Vetor3(0, 0) <> Empty Then
            If TxtValor.Text <> funSomaColuna(Vetor3, 4) Then
                MsgBox "A soma das distribuições é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                CmdGravar.Enabled = True
                Exit Sub
            End If
        Else
            MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
        
    If CboStatus.ListIndex <> 0 And Tipo = 0 And XTipoFatura <> "3" Then 'Não é previsto e não gera duplicata automaticamente.
        'Verifica se o total da nota é igual a soma das duplicatas com os descontos ou se alterou o valor da nota
        If VerificaValores Then
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
        
'    XStatus = funVerificaStatus(CboStatus.ListIndex)
    
    If Not ConsistenciaStatus Then
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    'Inserido em 27/10/08 - Larissa
    '***************CONTABILIDADE - PREVISTO*******************
    If XInserir = "A" And CboStatus.Text = "Previsto" Then
        If funMesFechado(DtpEntrada.Value) = "A" Then
            'Apagar os Lançamentos contábeis relacionados com a provisão desta NF, caso exista
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                            " (lanc_tx_Origem = 'FN') AND empr_cd_empresa = " & PCodEmpresa
        Else
            MsgBox "Não é possível alterar o status da Nota de compromissado para previsto, pois o mês contábil da provisão já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
    
    '**********************************************************
      
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
    
    If TxtDocPagto1.Text = "" And Tipo = "1" And XStatus = "R" Then 'Se Realizado sem nºDoc Pagto. gerar um automático
        XLT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    Else
        XLT_DOCPAGREC = TxtDocPagto1.Text
    End If
        
    '-------------------------------------------------------------------
    
    'Definição de XTipoFatura
    If XTipoFatura <> "3" Then
        If IsNull(DtpVencParcelaUnica.Value) Then
            If EntrouDupl Then
                If Vetor2.UpperBound(1) > 0 Then
                    XTipoFatura = 2                     '2 - Uma nota para várias duplicatas
                Else
                    If Vetor2(0, 0) = Empty Then
                       If EntrouDesc Then
                            If Vetor1(0, 0) <> Empty Then
                                XTipoFatura = 1         '1 - Não compôs duplicata porque o total de descontos é igual ao valor total da nota fiscal
                                XFB_PROVISIONA = True
                            Else
                                XTipoFatura = 4         '4 - Falta compor a fatura
                            End If
                        Else
                            XTipoFatura = 4             '4 - Falta compor a fatura
                        End If
                    Else
                        XTipoFatura = 1                 '1 - Uma nota para uma duplicata
                    End If
                End If
            Else
                If EntrouDesc Then
                    If Vetor1(0, 0) <> Empty Then
                        XTipoFatura = 1                 '1 - Não compôs duplicata porque o total de descontos é igual ao valor total da nota fiscal
                        XFB_PROVISIONA = True
                    Else
                        XTipoFatura = 4                 '4 - Falta compor a fatura
                    End If
                Else
                    If XInserir = "I" Then
                       XTipoFatura = 4
                    ElseIf Vetor2(0, 0) = Empty And Vetor1(0, 0) = Empty Then  'Só entra se for Alteração e o vetor de Duplic.estiver vazio
                        'Só entra se for Alteração e o vetor de Duplic. e o vetor de desc. estiverem vazio
                        XTipoFatura = 4
                    End If
                End If
            End If
        Else
            XTipoFatura = 1
        End If
    End If
        
    'O XTipoFatura = 3 será atribuído pela rotina de composição de fatura que significa: Várias notas para várias duplicatas
        
    'Definir o código da duplicata como "U" para caso de unica e "A" para caso de ser a primeira de uma série
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
        XLD_DATAENTRADA = CDate(DtpEntrada.Value)
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
                        
            If (chkPendente.Value = 1) Then
                ResFormulario!nofi_tx_pendestoque = "P"
            Else
                ResFormulario!nofi_tx_pendestoque = Null
            End If
        
            ResFormulario!nofi_tx_tipomov = Xtipomov
            ResFormulario!nofi_tx_status = XStatus
            ResFormulario!nofi_nr_documento = XNumDoc
            ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
            ResFormulario!nofi_tx_serie = XSerie
            ResFormulario!tido_cd_TipoDoc = CboTipoDoc.BoundText
            ResFormulario!nofi_dt_entrada = XLD_DATAENTRADA
            ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
            ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
            ResFormulario!nofi_tx_distribuir = XDistribuido
            ResFormulario!nofi_nr_fatura = xNumFat
            ResFormulario!nofi_tx_tipofatura = XTipoFatura
            ResFormulario!empr_cd_empresa = PCodEmpresa
'            ResFormulario!nofi_nr_chamado = FunNulo(TxtChamado.Text)
            
'            If chkSigiloso.Value = Checked Then
'                ResFormulario!nofi_bl_Sigiloso = 1
'            Else
'                ResFormulario!nofi_bl_Sigiloso = 0
'            End If
            
            ResFormulario.Update

            Chave = ResFormulario!nofi_cd_notafiscal
                
            If OptMes.Value Then
                XLD_DATAENTRADA = DateAdd("m", CInt(TxtPeriodicidade.Text), XLD_DATAENTRADA)
            ElseIf OptDia.Value Then
                XLD_DATAENTRADA = DateAdd("d", CInt(TxtPeriodicidade.Text), XLD_DATAENTRADA)
            End If
            
            'Grava distribuição p/ parcelas múltiplas
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
            " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
            Conexao.Execute Sql
            
            'Grava Duplicata da Nota de Parcelas Múltiplas
            Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                  " VALUES ('U'," & FunNuloData(DataVencimento, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & xNumFat & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
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
            If XTipoFatura <> 4 Then
                If ResFormulario!nofi_nr_fatura = 0 Then
                    xNumFat = FunPegaGuardaUltimo("F")
                Else
                    xNumFat = ResFormulario!nofi_nr_fatura
                End If
            Else
                xNumFat = 0
            End If
            
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
        
        If (chkPendente.Value = 1) Then
            ResFormulario!nofi_tx_pendestoque = "P"
        Else
            ResFormulario!nofi_tx_pendestoque = Null
        End If
            
        ResFormulario!nofi_tx_tipomov = Xtipomov
        ResFormulario!nofi_tx_status = XStatus
        ResFormulario!nofi_nr_documento = XNumDoc
        ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
        ResFormulario!nofi_tx_serie = XSerie
        ResFormulario!tido_cd_TipoDoc = CboTipoDoc.BoundText
        ResFormulario!nofi_dt_entrada = DtpEntrada.Value
        ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
        ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
        ResFormulario!nofi_tx_distribuir = XDistribuido
        ResFormulario!nofi_nr_fatura = xNumFat
        ResFormulario!nofi_tx_tipofatura = XTipoFatura
        ResFormulario!empr_cd_empresa = PCodEmpresa
'        ResFormulario!nofi_nr_chamado = FunNuloVal(TxtChamado.Text)
        
'        If chkSigiloso.Value = Checked Then
'            ResFormulario!nofi_bl_Sigiloso = 1
'        Else
'            ResFormulario!nofi_bl_Sigiloso = 0
'        End If
        
        ResFormulario.Update

        If XInserir = "A" And XStatus = "P" Then
           Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & xNumFat
        End If
        
        If XInserir = "I" Then
            Chave = ResFormulario!nofi_cd_notafiscal
        End If
        
    End If
        
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
    
    'Gera descrição de log da operação
    If XInserir = "I" Then
        'Nº de Doc e Nº de Fatura que foi gerado pelo sistema:
        XFM_MATRIZNOTAFISCAL(0, 1) = XNumDoc
        XFM_MATRIZNOTAFISCAL(1, 1) = xNumFat
        
        XFM_MATRIZNOTAFISCAL(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
        XFM_MATRIZNOTAFISCAL(3, 1) = TxtSerie.Text
        XFM_MATRIZNOTAFISCAL(4, 1) = CboStatus.Text
        XFM_MATRIZNOTAFISCAL(5, 1) = CboFornecedor.Text
        XFM_MATRIZNOTAFISCAL(6, 1) = CboTipoDoc.Text
        XFM_MATRIZNOTAFISCAL(7, 1) = Format$(DtpEntrada.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(8, 1) = Format$(DtpEmissao.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(9, 1) = FunNulo(CboHistorico.Text)
        
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, CADASTRAR, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, CADASTRAR, NOTAFISCAL_RECEBER)
    Else
        'Número de fatura original:
        XFM_MATRIZNOTAFISCAL(1, 1) = ResFormulario!nofi_nr_fatura
        XFM_MATRIZPAGAMENTO(0, 1) = ResFormulario!nofi_nr_fatura
        
        'Novos dados da nota:
        XFM_MATRIZNOTAFISCAL(0, 2) = MskNumdocumento.Text
        XFM_MATRIZNOTAFISCAL(2, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
        XFM_MATRIZNOTAFISCAL(3, 2) = TxtSerie.Text
        XFM_MATRIZNOTAFISCAL(4, 2) = CboStatus.Text
        XFM_MATRIZNOTAFISCAL(5, 2) = CboFornecedor.Text
        XFM_MATRIZNOTAFISCAL(6, 2) = CboTipoDoc.Text
        XFM_MATRIZNOTAFISCAL(7, 2) = Format$(DtpEntrada.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(8, 2) = Format$(DtpEmissao.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(9, 2) = FunNulo(CboHistorico.Text)
               
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, ALTERAR, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, ALTERAR, NOTAFISCAL_RECEBER)
    End If
    
    'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
    XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
    XFT_CONTADORLOG = XFT_CONTADORLOG + 1
    '**********************FIM REGISTRA LOG - PARTE 2 ***************************

    ResFormulario.Close
    
    'Descontos
    If EntrouDesc And XStatus <> "P" And Tipo = "0" Then
        If XInserir = "A" Then
              Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & xNumFat
        End If
        
        If Vetor1(0, 0) <> Empty Then
            subGravaVetor "DescontosNF", "deen_dt_prevrecolhimento,deen_vl_valor,desc_cd_desconto, deen_dt_recolhimento, deen_nr_DarfDam, deen_nr_fatura", Array("D", "N", "N", "D", "T", "N"), 5, Array(1, 2, 3, 5, 6), "" & xNumFat & "", Vetor1, TDBGrid1
        End If
    End If
        
    '***************CONTABILIDADE*******************
    SubQOpenRecordset ResFormulario, "SELECT dupl_vl_Valor,dupl_tx_status,dupl_cd_Duplicata, " & _
            " dupl_tx_DocPagRec, dupl_tx_tipopagrec, dupl_dt_PagRec, coco_cd_codigo, dupl_tx_estorno FROM Duplicatas " & _
            " WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
                
    While Not ResFormulario.EOF
        If ResFormulario!dupl_tx_status = "R" And (IsNull(ResFormulario!dupl_tx_Estorno) Or ResFormulario!dupl_tx_Estorno = "N" Or ResFormulario!dupl_tx_Estorno = "") Then
            'Como o código da duplicata será apagado qdo da gravacao da duplicata, foi necessário apagar os registros na tabela de lançamento para gravar novos no futuro
            
            'Início - Larissa Moura 02/06/2008
            
            'Verifica se a nota é uma Prestação de Contas, pois ela não possui conta corrente.
            SubQOpenRecordset XLT_TIPODOC, "SELECT tido_tx_PrestContas FROM TipoDocumentos WHERE tido_cd_TipoDoc = " & CboTipoDoc.BoundText, Estatico
            If XLT_TIPODOC!tido_tx_PrestContas <> "1" Then
            
            'Fim - Larissa Moura 02/06/2008
                    XLT_SQL = "SELECT duplicatas.dupl_cd_Duplicata FROM (tipodocumentos " & _
                                       " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                       " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                       " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                       " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura " & _
                                       " WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                       " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                       " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N' or dupl_tx_estorno = '')"
            Else
                    XLT_SQL = "SELECT duplicatas.dupl_cd_Duplicata FROM (tipodocumentos " & _
                                      " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                      " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                      " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                      " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura " & _
                                      " WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo is null " & _
                                      " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                      " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N' or dupl_tx_estorno = '')"

            End If
            XLT_TIPODOC.Close
            SubQOpenRecordset RESFAT2, XLT_SQL, Estatico
            While Not RESFAT2.EOF
                    Sql = "DELETE FROM Lancamentos " & _
                        "WHERE lanc_cd_OrigemFin = " & RESFAT2!dupl_cd_Duplicata & " AND " & _
                                "lanc_tx_Origem = 'FD' AND " & _
                                "empr_cd_empresa = " & PCodEmpresa
                    Conexao.Execute Sql
                RESFAT2.MoveNext
            Wend
        End If
        ResFormulario.MoveNext
    Wend
    ResFormulario.Close
    
    '***********************************************
    
    'Duplicatas
    If Not ParcelasMultiplas Then
        If EntrouDupl Then 'composição de duplicata
            If XInserir = "A" Then
                If Vetor4(0, 0) <> Empty Then
                    GravaVetorDuplDescAcresc
                Else
                
                    Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & xNumFat & " and (dupl_tx_estorno <> 'S' or dupl_tx_estorno is null)" 'Não excluir os lançamentos de estorno
                    
                    'Exclui os versos de cheques que existem associados a este cheque
                    Conexao.Execute "DELETE FROM VersoCheques where vech_tx_numcheque= '" & FunNulo(XLT_DOCPAGRECVERSO) & "' and coco_cd_codigo = " & XLT_CONTAVERSO
                    
                    subGravaVetorDupl "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_estorno,dupl_tx_Consolidado,dupl_dt_concilia,dupl_dt_ProcPagamento,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "TN", "T", "D", "D", "N", "T"), 15, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13, 16, 17, 18, 19), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                End If
            End If
            
            If XInserir = "I" Then
                If Vetor2(0, 0) <> Empty Then
                    subGravaVetorDupl "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_estorno,dupl_tx_Consolidado,dupl_dt_concilia,dupl_dt_ProcPagamento,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "TN", "T", "D", "D", "N", "T"), 15, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13, 16, 17, 18, 19), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                Else
                    'O usuário pode ter entrado na orelha de duplicatas, mas ter voltado e decidido fazer parcela única
                    If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente ou tem status = previsto
                         Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                               " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                         Conexao.Execute Sql
                    End If
                End If
            End If
        Else 'geração automática de duplicata
            If XInserir = "I" Then
               If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente ou tem status = previsto
                    Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                          " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                    Conexao.Execute Sql
                    
                    'Dados do pagamento
                    XFM_MATRIZPAGAMENTO(0, 1) = Str(xNumFat)
                    XFM_MATRIZPAGAMENTO(1, 1) = "U"
                    XFM_MATRIZPAGAMENTO(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                    XFM_MATRIZPAGAMENTO(3, 1) = CboStatus.Text
                    XFM_MATRIZPAGAMENTO(4, 1) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
                    XFM_MATRIZPAGAMENTO(5, 1) = CboTipoPagto1.Text
                    XFM_MATRIZPAGAMENTO(6, 1) = CboCCorrente.Text
                    XFM_MATRIZPAGAMENTO(7, 1) = FunNulo(CboBancoForn1.Text)
                    
                    If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                        XFM_MATRIZPAGAMENTO(8, 1) = TxtDocPagto1.Text
                        XFM_MATRIZPAGAMENTO(9, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
                    End If
               End If
            ElseIf XInserir = "A" And (Tipo = "1" Or XStatus = "P") Then
                'A rotina foi modificada para atender às alterações de Dupl que possuam Desc/Acresc
                Conexao.Execute "UPDATE Duplicatas " & _
                                    "SET dupl_nr_duplicata ='" & XCodDupl & "', " & _
                                    "dupl_dt_vencimento = " & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & _
                                    "dupl_vl_GlosaDevolucao = " & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & _
                                    "dupl_vl_valor = " & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & "," & _
                                    "dupl_tx_status = '" & XStatus & "'," & _
                                    "dupl_tx_tipopagrec = '" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "'," & _
                                    "dupl_tx_docpagrec = '" & FunNulo(XLT_DOCPAGREC) & "'," & _
                                    "coco_cd_codigo= " & FunNuloValor(CboCCorrente) & "," & _
                                    "bafo_cd_codigo = " & FunNuloValor(CboBancoForn1) & "," & _
                                    "dupl_tx_debcred ='" & Xdebcred & "'," & _
                                    "dupl_nr_fatura = " & Str(xNumFat) & "," & _
                                    "dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & " " & _
                                    "WHERE dupl_nr_fatura = " & xNumFat & " AND (dupl_tx_estorno is null  OR  dupl_tx_estorno = '' OR dupl_tx_estorno = 'N')", XLB_REGISTROAFETADO
                
                If Not XLB_REGISTROAFETADO Then
                    Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                          " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                    Conexao.Execute Sql
                End If
            End If
            
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
            
            'Gera descrição de log da operação
            If XInserir = "I" Then
                If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, CADASTRAR_DUPLICATA, NOTAFISCAL_PAGAR)
                If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, CADASTRAR_DUPLICATA, NOTAFISCAL_RECEBER)
            Else
                'Novos dados do pagamento:
                XFM_MATRIZPAGAMENTO(0, 2) = Str(xNumFat)
                XFM_MATRIZPAGAMENTO(1, 2) = "U"
                XFM_MATRIZPAGAMENTO(2, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                XFM_MATRIZPAGAMENTO(3, 2) = CboStatus.Text
                XFM_MATRIZPAGAMENTO(4, 2) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
                XFM_MATRIZPAGAMENTO(5, 2) = CboTipoPagto1.Text
                XFM_MATRIZPAGAMENTO(6, 2) = CboCCorrente.Text
                XFM_MATRIZPAGAMENTO(7, 2) = FunNulo(CboBancoForn1.Text)
                                           
                If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                    XFM_MATRIZPAGAMENTO(8, 1) = TxtDocPagto1.Text
                    XFM_MATRIZPAGAMENTO(9, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
                Else
                    XFM_MATRIZPAGAMENTO(8, 0) = ""
                    XFM_MATRIZPAGAMENTO(9, 0) = ""
                End If
                                           
                If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, ALTERAR_DUPLICATA, NOTAFISCAL_PAGAR)
                If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, ALTERAR_DUPLICATA, NOTAFISCAL_RECEBER)
            End If
        
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '******************FIM REGISTRA LOG DA OPERAÇÃO - PARTE 2 *******************
        End If
    End If
    
    'Distribuição
    If EntrouDist Then
        If XInserir = "A" Then
              Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
        End If
        
        If Vetor3(0, 0) <> Empty Then
            subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N", "N"), 5, Array(0, 2, 5, 6, 8), "" & Chave & "", Vetor3, TDBGrid3
        Else
            'se entrou na orelha de distribuicao excluiu o rateio e cadastrou na orelha de apropriação
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
                " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
            Conexao.Execute Sql
        End If
    Else
        If ParcelasMultiplas = False Then
            If XInserir = "A" Then
                Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
            End If
            
            If ChkDistribuicao Then
                If Vetor3(0, 0) <> Empty Then
                    subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N", "N"), 5, Array(0, 2, 5, 6, 8), "" & Chave & "", Vetor3, TDBGrid3
                End If
            Else
                            
                Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
                      " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
                Conexao.Execute Sql
                
                'Dados da apropriação
                XFM_MATRIZAPROPRIACAO(0, 1) = XNumDoc
                XFM_MATRIZAPROPRIACAO(1, 1) = CboPcr.Text
                XFM_MATRIZAPROPRIACAO(2, 1) = CboPlanoContas.Text
                XFM_MATRIZAPROPRIACAO(3, 1) = "100%"
                XFM_MATRIZAPROPRIACAO(4, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                XFM_MATRIZAPROPRIACAO(5, 1) = CboIdentificador2.Text
                
                '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
                'Gera descrição de log da operação
                If XInserir = "I" Then
                    If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
                    If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
                Else
                    'Novos dados da apropriação:
                    XFM_MATRIZAPROPRIACAO(0, 2) = MskNumdocumento.Text
                    XFM_MATRIZAPROPRIACAO(1, 2) = CboPcr.Text
                    XFM_MATRIZAPROPRIACAO(2, 2) = CboPlanoContas.Text
                    XFM_MATRIZAPROPRIACAO(3, 2) = "100%"
                    XFM_MATRIZAPROPRIACAO(4, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                    XFM_MATRIZAPROPRIACAO(5, 2) = FunNulo(CboIdentificador2.Text)
                                               
                    If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, ALTERAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
                    If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, ALTERAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
                End If
            
                'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
                XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
                XFT_CONTADORLOG = XFT_CONTADORLOG + 1
                '******************FIM REGISTRA LOG DA OPERAÇÃO - PARTE 2 *******************
            End If
        End If
    End If
    
    'Verificar se é preciso atualizar composição de fatura
    If XValorBruto <> CDbl(TxtValor.Text) And XTipoFatura = "3" And XInserir = "A" Then
        If MsgBox("É necessário atualizar a Composição da Fatura. Deseja atualizar agora ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
            'Chamar tela de composição de fatura
            NFatura = xNumFat
            XFormulario = "FrmNotasFiscais"
            FrmFatura.Show 1
        End If
    End If
    
    '***************CONTABILIDADE*******************
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If XGT_TIPOEXPORTACAO = "PC" And _
            (CboStatus.Text <> "Previsto" And _
            (IsNull(DatTipoDoc.Recordset.Fields("tido_tx_sigla")) Or _
            (DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") <> "1" And _
            DatTipoDoc.Recordset.Fields("tido_tx_sigla") <> "NP"))) Then  'Exporta base provisão
            
            'Parte do Caixa para registros provisionados
            SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata, dupl_dt_pagrec, " & _
                            " dupl_tx_Status, dupl_tx_TipoPagRec, coco_cd_codigo, " & _
                            " dupl_tx_docpagrec, dupl_vl_valor, dupl_tx_estorno, dupl_tx_debcred " & _
                            " FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
            
            If XFB_PROVISIONA And xNumFat <> 0 Then
                'xNumFat <> 0 significa que não tem parcela e a fatura será composta futuramente
                'nesses casos não é possível provisionar
'                SubProvisao xNumFat, DtpEntrada.Value
            ElseIf Not XFB_PROVISIONA And XInserir = "A" Then
                'Apagar os Lançamentos contábeis relacionados com a provisão desta NF, caso exista
                Conexao.Execute "DELETE FROM Lancamentos " & _
                                "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                                    " (lanc_tx_Origem = 'FN') AND " & _
                                    "empr_cd_empresa = " & PCodEmpresa
            End If
            
            While Not ResFormulario.EOF
                'Em se tratando de parcela, só contabiliza qdo. realizado
                'E quando for estorno só considerar o valor > 0
                If ResFormulario!dupl_tx_status = "R" And ResFormulario!dupl_vl_valor > 0 Then
                    'se a duplicata está sendo cadastrada como realizada ou se alterou para realizada,
                    'Se alterou, deverá ser verificado se foi consolidado.
                    'se a duplicata já era realizada e está sendo alterada
                    
                    'RECONTABILIZAR O ESTORNO SOMENTE SE A NOTA QUE ESTÁ SENDO ALTERADA POSSUIR UM
                    If ResFormulario!dupl_tx_Estorno = "S" Then
                    Else
                        'VERIFICA SE FOI CONSOLIDADO
                          XLT_SQL = "SELECT Count(*) as ContReg FROM (tipodocumentos " & _
                                    " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                    " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                    " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                    " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                    " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                    " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N'  or dupl_tx_estorno = '')"
    
                        SubQOpenRecordset ResFat, XLT_SQL, Estatico
                        If ResFat!ContReg > 1 Then
                            'SE ESTÁ ALTERANDO UM CONSOLIDADO NÃO DEVE APAGAR NEM RECADASTRAR A PARTE DO VALOR CONSOLIDADO
'                            SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, True, False
                        Else
'                            SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, False, False
                        End If
                    End If
                End If
                ResFormulario.MoveNext
            Wend
            ResFormulario.Close
        Else 'Exporta base caixa
            SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata, dupl_dt_pagrec, " & _
                        " dupl_tx_Status, dupl_tx_TipoPagRec, coco_cd_codigo, dupl_tx_docpagrec " & _
                        " FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
            
            While Not ResFormulario.EOF
                If ResFormulario!dupl_tx_status = "R" Then 'Em se tratando de parcela, só contabiliza qdo. realizado
                    'se a duplicata está sendo cadastrada como realizada ou se alterou para realizada,
                    'Se alterou, deverá ser verificado se foi consolidado.
                    'se a duplicata já era realizada e está sendo alterada

                    XLT_SQL = "SELECT Count(*) as ContReg FROM (tipodocumentos " & _
                                " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                " and dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & ""

                    SubQOpenRecordset ResFat, XLT_SQL, Estatico
                    
                    If ResFat!ContReg > 1 Then
                        'SE ESTÁ ALTERANDO UM CONSOLIDADO NÃO DEVE APAGAR NEM RECADASTRAR
                        'A PARTE DO VALOR CONSOLIDADO
'                        SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, True, False
                    Else
'                        SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, False, False
                    End If
                End If
                ResFormulario.MoveNext
            Wend
            ResFormulario.Close
        End If
    End If
    
    '***********************************************
    Conexao.CommitTrans
        
    '***************************** GRAVA LOG NO BANCO **********************************
    Dim aux As Integer
    Dim inicioLog As String
    Dim fimLog As String
    
    For Cont = 0 To (XFT_CONTADORLOG - 1)
        XFT_LOG = XFV_VETORLOG(Cont)
        
        'No caso de cadastro de nota fiscal, insere o número da nota fiscal gerado
        If XInserir = "I" Then
            aux = InStr(1, XFT_LOG, "NOTA FISCAL:", vbTextCompare)
            If aux > 0 Then
                aux = aux + 12
                inicioLog = Mid(XFT_LOG, 1, aux) + XNumDoc
                fimLog = Mid(XFT_LOG, aux, Len(XFT_LOG) - aux)
                XFT_LOG = inicioLog + fimLog
            End If
        End If
        
        If EntSaida = "E" Then Call subRegistraLog("TelaNotasFiscais", "2", XFT_LOG)
        If EntSaida = "S" Then Call subRegistraLog("TelaNotasFiscais", "1", XFT_LOG)
    Next
    
    '************************** FIM GRAVA LOG NO BANCO **********************************
    
    If XInserir = "A" Then
        If cmdNovo.Enabled = False Then
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            
        Else
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = False '26/11/09 - Patrícia
        End If
    Else
        XInserir = "A"
        NumDocAnt = MskNumdocumento.Text
        cmdNovo.Enabled = True
        
        If Not EntrouDesc Then
            Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
            subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura", "deen_dt_recolhimento", "deen_nr_DarfDam"), Vetor1, TDBGrid1
        End If
        Sql = "SELECT DuplDescAcresc.* from DuplDescAcresc LEFT JOIN Duplicatas ON Duplicatas.dupl_cd_duplicata = DuplDescAcresc.dupl_cd_duplicata Where Duplicatas.dupl_nr_fatura = " & NFatura
        CarregaVetorDescAcresc Sql, Array("dupl_cd_duplicata", "deac_cd_descacresc", "dude_vl_descacresc"), Vetor4, 3
            
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = False 'Daniel Coelho 09/02/2012
    End If
        
    XFT_NOVO = 0 '13/11/09 - Patrícia

FimGrava:

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
    Dim ResEstoque As Object
    Dim ResEmpObr As Object 'Declaração em 05/06/08 - Patrícia
    Dim Desc_CREmpObr As String 'Declaração em 05/06/08 - Patrícia
    Dim Desc_CCEmpObr As String 'Declaração em 05/06/08 - Patrícia
    Dim DataVencimento As Date
    Dim XLT_DOCPAGREC As String
    Dim ChaveTemp As Variant
    Dim XLT_INSERIRTEMP As String
    Dim XLB_REGISTROAFETADO As Boolean
    Dim XLT_SQL As String
    Dim RESFAT2 As Object
    Dim XLD_DATAENTRADA As Date
    Dim XLT_TIPODOC As Object 'Criado para verificar se o tipo de documento é prestação de contas - Larissa Moura 02/06/2008
    Dim teste As String
    Dim TelaAnterior As Integer ' Declarado em 09/06/08 - Patrícia
    Dim Cont As Integer 'Larissa 30/06/2008
    Dim ResVencDuplicatas As New ADODB.Recordset 'Patrícia 05/08/2010
    Dim XLD_DATAVENCDUPL As Variant 'Patrícia 05/08/2010
    Dim XLD_DATAPAGDUPL As Variant 'Patrícia 05/08/2010
    Dim XLO_RS As ADODB.Recordset
    Dim XLB_TIPODOCINATIVO As Boolean
    
    If Mid(PEmpresa, 1, 5) = "APAE" Then
        If XLT_VALORDIFFNOTA <= 0 Then
            Beep
            MsgBox "A soma das NF recebidas são inferiores a soma de NF de pagemento.", vbInformation, "AVISO"
        End If
    End If
    
    If EntSaida = "S" Then
        If (CboStatus.Text = "Realizado") Then
            If Not FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "3") Then
                Beep
                MsgBox "Você não possui permissão para salvar um documento como Realizado.", 16, "AVISO"
                Exit Sub
            End If
        End If
    End If
        
    XLB_TIPODOCINATIVO = False
    
    SubQOpenRecordset XLO_RS, "SELECT * FROM Tipodocumentos WHERE tido_bl_ativo = 0 AND tido_cd_TipoDoc = " & DatTipoDoc.Recordset("tido_cd_TipoDoc").Value, Estatico
    While Not XLO_RS.EOF And Not XLO_RS.BOF
        XLB_TIPODOCINATIVO = True
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
    Set XLO_RS = Nothing
    
    If (XLB_TIPODOCINATIVO) Then
        MsgBox "O tipo de documento escolhido está inativo. Escolha outro.", vbInformation, "ATENCAO"
        Exit Sub
    End If
     
    If XInserir = "A" And Tipo = "0" Then
    SubQOpenRecordset ResVencDuplicatas, "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura & "order by dupl_dt_vencimento", Dinamico
    'Pega a data de vencimento da primeira duplicata
    If Not ResVencDuplicatas.EOF Then
        XLD_DATAVENCDUPL = ResVencDuplicatas!dupl_dt_vencimento
        XLD_DATAPAGDUPL = ResVencDuplicatas!dupl_dt_pagrec
    End If
    ResVencDuplicatas.Close
    Set ResVencDuplicatas = Nothing
    End If
    
    'Larissa, 28/01/09
    'Verifica se o usuário cadastrou distribuição
    If ChkDistribuicao.Value = 1 And EntrouDist = True And (LblValorTotalDistrib = "" Or LblValorTotalDistrib = "0,00") Then
        MsgBox "É necessário cadastrar a distribuição da nota!", vbInformation, "ATENCAO"
        Exit Sub
    End If
        
    'Inserido em 09/06/08 - Patrícia
    ' Verifica a Classificação "Empreendimento\Obra" dos Centros de Custos e Contas Contábeis da Distribuição
    If (DtpEmissao) >= CDate(Data_EmpObr) Then
                If ChkDistribuicao.Value = 0 Then 'Verifica se houve distribuição
                    XMENSAGEM = FunVerificaClassEmprObr(Str(DatPCR.Recordset.Fields("cere_cd_Pcr")), MskConta.Text, PCodEmpresa)
                     If XMENSAGEM <> "" Then
                        MsgBox XMENSAGEM & "!", vbCritical, "AVISO"
                        CboPlanoContas.SetFocus
                        Exit Sub
                     End If
                Else
                    'Se houver distribuição, faz a verificação da classificação "Empreendimento\Obra"
                     TDBGrid3.MoveFirst
                     Do While Not TDBGrid3.EOF
                          XMENSAGEM = FunVerificaClassEmprObr(TDBGrid3.Columns("CodPcr").Value, TDBGrid3.Columns("Conta Contábil").Value, PCodEmpresa)
                          If XMENSAGEM <> "" Then
                             'Mostra a Tela de Distribuição
                             MsgBox XMENSAGEM & "!", vbCritical, "AVISO"
                             TelaAnterior = TabNotasFiscais.Tab
                             TabNotasFiscais.Tab = 3
                             TabNotasFiscais_Click (TelaAnterior)
                             Exit Sub
                          End If
                          TDBGrid3.MoveNext
                      Loop
               
                End If
    End If
    'Fim 09/06/08 - Patrícia
    
    Reload = True
    
    CmdGravar.Enabled = False
    
    Me.MousePointer = vbHourglass
    
    TabNotasFiscais.Tab = 0
    If ChecarCamposOrelhaPrincipal Then
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    'verificar se possui duplicatas canceladas para essa nota
    If FunVerificaDuplicataCancelada = False Then
        MsgBox "Notas Fiscais com duplicatas canceladas não podem ser alteradas!", vbCritical, "AVISO"
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
        
    '/////////////////////////////////////////////////////////////////////////////////
    'SE A NOTA É DO ESTOQUE O QUALIFIN SÓ PODE ALTERAR O IDENTIFICADOR E SAI DA ROTINA
    SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
    If (XFB_ESTOQUE = True) And (ResEstoque!empr_tx_AltNotasEstoque = "N") Then
        'Não permite alterar notas do estoque, se não tiver selecionado o campo que
        'permite alterar em padrões de empresa - Lucas Santiago 23.01.09
        If ChkDistribuicao.Value = 1 Then
            If Vetor3(0, 0) <> Empty Then
                For x = 0 To Vetor3.UpperBound(1)
                   Conexao.Execute ("UPDATE Distribuicao SET iden_cd_Projeto = " & FunNuloBancoVal(Vetor3(x, 8)) & "" & _
                                    " WHERE nofi_cd_NotaFiscal = " & Chave & " AND cere_cd_Pcr = '" & Vetor3(x, 0) & "'" & _
                                    " AND plco_cd_conta = '" & Vetor3(x, 2) & "'")
                Next
            End If
        Else
            Conexao.Execute ("UPDATE Distribuicao SET iden_cd_Projeto = " & FunNuloBancoVal(CboIdentificador2.BoundText) & "" & _
                                " WHERE nofi_cd_NotaFiscal = " & Chave & "")
        End If
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Unload Me
        Exit Sub
    End If
    '//////////////////////////////////////////////////////////////////////////////////
    
    '***************CONTABILIDADE - COMPOSIÇÃO DE DUPLICATA - QUANDO DA ALTERAÇÃO ***********
    'Se o usuário abrisse esse form para consultar a nota e gravasse sem confirma as duplicatas
    'o valor da parcela era lançado na contabilidade. Deve-se então, varrer o vertor para confirmar a contabilização
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If XInserir = "A" Then 'Só para cadastros na orelha de composição de duplicatas
            For x = 0 To Vetor2.UpperBound(1)
                If XFB_PROVISIONA = False Then 'se já foi definido para provisionar não precisa continuar a pesquisa
                    If IsDate(Vetor2(x, 10)) Then
                        If Month(CDate(Vetor2(x, 10))) = Month(DtpEntrada.Value) Then
                            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                                If ((CDate(Vetor2(x, 10)) - DtpEntrada.Value) >= XGI_QTDDIASPROVISAO) Then
                                    XFB_PROVISIONA = True
                                Else
                                    XFB_PROVISIONA = False
                                    'Verificar se o mês está fechado com base na data de pagto das parcelas
                                    If XGB_PQUALICONT Then  'possui o módulo de contabilidade
                                        If funMesFechado(Vetor2(x, 10)) <> "A" Then
                                            MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                                                + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Else
                                XFB_PROVISIONA = False
                                'Verificar se o mês está fechado com base na data de pagto das parcelas
                                If XGB_PQUALICONT Then  'possui o módulo de contabilidade
                                    If funMesFechado(Vetor2(x, 10)) <> "A" Then
                                        MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                                            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                                        Me.MousePointer = vbDefault
                                        CmdGravar.Enabled = True
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            XFB_PROVISIONA = True
                        End If
                    Else
                        If Month(CDate(Vetor2(x, 1))) = Month(DtpEntrada.Value) Then
                            If XGI_QTDDIASPROVISAO >= 0 And CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" Then  '13/04/10
                                If ((CDate(Vetor2(x, 1)) - DtpEntrada.Value) >= XGI_QTDDIASPROVISAO) Then
                                    XFB_PROVISIONA = True
                                'Else
                                End If
                            Else
                                XFB_PROVISIONA = False
                            End If
                        Else
                            XFB_PROVISIONA = True
                        End If
                    End If
                Else
                    x = Vetor2.UpperBound(1)
                End If
            Next
        End If
    End If
    
    '****************************************************************************************
    'Em vista do que foi acordado para a realização da manutenção permanente do sistema,informo que a maneira pela qual o sistema
    '***************CONTABILIDADE - REALIZADO*******************
    'possui o módulo de contabilidade
    If XGB_PQUALICONT And DatTipoDoc.Recordset.Fields("tido_tx_nec") = "0" Then
        'Só entra na checagem os que estiverem sendo realizados através da alteração
        If Not IsNull(DtpPagtoParcelaUnica) Then
            If funMesFechado(DtpPagtoParcelaUnica.Value) <> "A" Then
                MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                CmdGravar.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    '**********************************************************
    If Tipo = "1" Or CboStatus.Text = "Previsto" Then 'Gera duplicata automaticamente
        XCodDupl = "U" 'Parcela única
        XFB_PROVISIONA = funVerificaProvisao(DtpPagtoParcelaUnica, DtpEntrada, DtpVencParcelaUnica, XGI_QTDDIASPROVISAO, CboStatus.Text)
    End If
    
    '***************CONTABILIDADE*******************
    'Comp ou Realizado, possui o módulo de contabilidade, exporta base provisão, e se pela data é preciso provisionar
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If CboStatus.Text <> "Previsto" And XGT_TIPOEXPORTACAO = "PC" And XFB_PROVISIONA Then
            'A variável XLB_PROVISIONA foi incluida p/ verificar se (DtPagRec-DtEntrada) >= XGI_QTDDIASPROVISAO
            If XFB_PROVISIONA Then
                If funMesFechado(DtpEntrada.Value) <> "A" Then
                    MsgBox "Não é possível cadastrar/alterar a Nota nessa data, pois o mês contábil já foi fechado. " _
                        + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                Else
                    If Tipo = 1 Then ' Se gera duplicata automaticamente entao Tipo =1
                        XFB_PROVISIONA = funVerificaProvisao(IIf(IsNull(DtpPagtoParcelaUnica.Value), "", DtpPagtoParcelaUnica.Value), IIf(IsNull(DtpEntrada.Value), "", DtpEntrada.Value), IIf(IsNull(DtpVencParcelaUnica.Value), "", DtpVencParcelaUnica.Value), XGI_QTDDIASPROVISAO, CboStatus.Text)
'                    Else
'                        XFB_PROVISIONA = funVerificaProvisao(IIf(IsNull(XLD_DATAPAGDUPL), "", XLD_DATAPAGDUPL), IIf(IsNull(DtpEntrada.Value), "", DtpEntrada.Value), IIf(IsNull(XLD_DATAVENCDUPL), "", XLD_DATAVENCDUPL), XGI_QTDDIASPROVISAO, CboStatus.Text)
                    End If
                    
                    If XFB_PROVISIONA = True Then
                            DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
                            'Se for adiantamento, não precisa checar a conta de adiantamento
                            If DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") <> "1" Then
                                While IsNull(DatFornecedor.Recordset.Fields("plco_cd_ContaCusto")) Or DatFornecedor.Recordset.Fields("plco_cd_ContaCusto") = ""
                                    'mas q falta de vontade de trabalhar viu!!!! essa semana demorou pra porra de passar
                                    ChaveTemp = Chave
                                    If EntSaida = "S" Then
                                        If MsgBox("O fornecedor não possui Conta Contábil para provisão." _
                                            + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro da nota.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            Exit Sub
                                        End If
                                    Else
                                        If MsgBox("O Cliente não possui Conta Contábil para provisão." _
                                            + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro da nota.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                                            Me.MousePointer = vbDefault
                                            CmdGravar.Enabled = True
                                            Exit Sub
                                        End If
                                    End If
                                    Pessoa = DatFornecedor.Recordset.Fields("focl_tx_Tipo")
                                    FornClie = DatFornecedor.Recordset.Fields("focl_tx_classe")
                                    
                                    XFormForn = "FrmNotasFiscais"
                                    XLT_INSERIRTEMP = XInserir
                                    XInserir = "A"
                                    Chave = DatFornecedor.Recordset.Fields("focl_cd_forncli")
                                    FrmFornecedores.TxtCgc = DatFornecedor.Recordset.Fields("focl_tx_cgccpf")
                                    FrmFornecedores.Show 1
                                    Chave = ChaveTemp
                                    XInserir = XLT_INSERIRTEMP
                                    FrmNotasFiscais.DatFornecedor.Refresh
                                    FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
                                    DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
                                Wend
                            End If
                    End If
                End If
            End If
        End If
    End If
    
    DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
    
    'Verifica para os casos de adiantamento, se o fornecedor possui conta de adiantamento cadastrada
    If DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") = "1" Then
        If XInserir = "A" Then
            'Foi necessário incluir a exclusão da provisão por que o usuário
            'pode estar alterando uma nota que foi provisionada para um adiantamento
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                            " (lanc_tx_Origem = 'FN') AND " & _
                            "empr_cd_empresa = " & PCodEmpresa
        End If
        
        While IsNull(DatFornecedor.Recordset.Fields("plco_cd_ContaProvAdiant")) Or DatFornecedor.Recordset.Fields("plco_cd_ContaProvAdiant") = ""
            ChaveTemp = Chave
            If EntSaida = "S" Then
                If MsgBox("O fornecedor não possui Conta Contábil para Adiantamento." _
                    + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro do Adiantamento.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                End If
            Else
                If MsgBox("O Cliente não possui Conta Contábil para Adiantamento." _
                    + Chr(13) + "Clique em 'Sim' para cadastrá-la ou 'Não' para sair do cadastro do Adiantamento.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
                    Me.MousePointer = vbDefault
                    CmdGravar.Enabled = True
                    Exit Sub
                End If
            End If
            Pessoa = DatFornecedor.Recordset.Fields("focl_tx_Tipo")
            FornClie = DatFornecedor.Recordset.Fields("focl_tx_classe")
            
            XFormForn = "FrmNotasFiscais"
            XLT_INSERIRTEMP = XInserir
            XInserir = "A"
            Chave = DatFornecedor.Recordset.Fields("focl_cd_forncli")
            FrmFornecedores.TxtCgc = DatFornecedor.Recordset.Fields("focl_tx_cgccpf")
            FrmFornecedores.Show 1
            Chave = ChaveTemp
            XInserir = XLT_INSERIRTEMP
            FrmNotasFiscais.DatFornecedor.Refresh
            FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
            DatFornecedor.Recordset.Bookmark = CboFornecedor.SelectedItem
        Wend
    End If
    
    '***********************************************
            
    If XDistribuido = "N" And (Trim(MskPcr.Text) = "" And MskConta.Text = "") Then 'Alterado em 31/07/08 - Larissa
        MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    ElseIf ChkDistribuicao.Value = 1 Then
        If Vetor3(0, 0) <> Empty Then
            If TxtValor.Text <> funSomaColuna(Vetor3, 4) Then
                MsgBox "A soma das distribuições é diferente do valor bruto da Nota!", vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                CmdGravar.Enabled = True
                Exit Sub
            End If
        Else
            MsgBox "É necessário realizar Distribuição", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
        
    If CboStatus.ListIndex <> 0 And Tipo = 0 And XTipoFatura <> "3" Then 'Não é previsto e não gera duplicata automaticamente.
        'Verifica se o total da nota é igual a soma das duplicatas com os descontos ou se alterou o valor da nota
        If VerificaValores Then
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
        
'    XStatus = funVerificaStatus(CboStatus.ListIndex)
    
    If Not ConsistenciaStatus Then
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    'Inserido em 27/10/08 - Larissa
    '***************CONTABILIDADE - PREVISTO*******************
    If XInserir = "A" And CboStatus.Text = "Previsto" Then
        If funMesFechado(DtpEntrada.Value) = "A" Then
            'Apagar os Lançamentos contábeis relacionados com a provisão desta NF, caso exista
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                            " (lanc_tx_Origem = 'FN') AND empr_cd_empresa = " & PCodEmpresa
        Else
            MsgBox "Não é possível alterar o status da Nota de compromissado para previsto, pois o mês contábil da provisão já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
    
    '**********************************************************
      
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
    
    If TxtDocPagto1.Text = "" And Tipo = "1" And XStatus = "R" Then 'Se Realizado sem nºDoc Pagto. gerar um automático
        XLT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    Else
        XLT_DOCPAGREC = TxtDocPagto1.Text
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
                       If EntrouDesc Then
                            If Vetor1(0, 0) <> Empty Then
                                XTipoFatura = 1        '1 - Não compôs duplicata porque o total de descontos é igual ao valor total da nota fiscal
                                XFB_PROVISIONA = True
                            Else
                                XTipoFatura = 4        '4 - Falta compor a fatura
                            End If
                        Else
                            XTipoFatura = 4        '4 - Falta compor a fatura
                        End If
                    Else
                        XTipoFatura = 1    '1 - Uma nota para uma duplicata
                    End If
                End If
            Else
                If EntrouDesc Then
                    If Vetor1(0, 0) <> Empty Then
                        XTipoFatura = 1        '1 - Não compôs duplicata porque o total de descontos é igual ao valor total da nota fiscal
                        XFB_PROVISIONA = True
                    Else
                        XTipoFatura = 4        '4 - Falta compor a fatura
                    End If
                Else
                    If XInserir = "I" Then
                       XTipoFatura = 4
                    ElseIf Vetor2(0, 0) = Empty And Vetor1(0, 0) = Empty Then  'Só entra se for Alteração e o vetor de Duplic.estiver vazio
                        'Só entra se for Alteração e o vetor de Duplic. e o vetor de desc. estiverem vazio
                        XTipoFatura = 4
                    End If
                End If
            End If
        Else
            XTipoFatura = 1
        End If
    End If
        
    'O XTipoFatura = 3 será atribuído pela rotina de composição de fatura que significa: Várias notas para várias duplicatas
    
    'Definir o código da duplicata como "U" para caso de unica e "A" para caso de ser a primeira de uma série
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
    
    If ChkDistribuicao.Value = 1 Then
        XDistribuido = "S"
    Else
        XDistribuido = "N"
    End If
        
    If ParcelasMultiplas = True Then
        DataVencimento = CDate(DtpVencParcelaUnica.Value)
        XLD_DATAENTRADA = CDate(DtpEntrada.Value)
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
                        
            If (chkPendente.Value = 1) Then
                ResFormulario!nofi_tx_pendestoque = "P"
            Else
                ResFormulario!nofi_tx_pendestoque = Null
            End If
                    
            ResFormulario!nofi_tx_tipomov = Xtipomov
            ResFormulario!nofi_tx_status = XStatus
            ResFormulario!nofi_nr_documento = XNumDoc
            ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
            ResFormulario!nofi_tx_serie = XSerie
            ResFormulario!tido_cd_TipoDoc = CboTipoDoc.BoundText
            ResFormulario!nofi_dt_entrada = XLD_DATAENTRADA
            ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
            ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
            ResFormulario!nofi_tx_distribuir = XDistribuido
            ResFormulario!nofi_nr_fatura = xNumFat
            ResFormulario!nofi_tx_tipofatura = XTipoFatura
            ResFormulario!empr_cd_empresa = PCodEmpresa
'            ResFormulario!nofi_nr_chamado = FunNuloVal(TxtChamado.Text)
            
            If (EntSaida = "S") And (PCliente = "QUALIDADOS") Then
                ResFormulario!plco_cd_contacredito = DatContaContabilCredito.Recordset.Fields("plco_cd_conta")
            End If
            
'            If chkSigiloso.Value = Checked Then
'                ResFormulario!nofi_bl_Sigiloso = 1
'            Else
'                ResFormulario!nofi_bl_Sigiloso = 0
'            End If
            
            ResFormulario.Update

            Chave = ResFormulario!nofi_cd_notafiscal
                
            If OptMes.Value Then
                XLD_DATAENTRADA = DateAdd("m", CInt(TxtPeriodicidade.Text), XLD_DATAENTRADA)
            ElseIf OptDia.Value Then
                XLD_DATAENTRADA = DateAdd("d", CInt(TxtPeriodicidade.Text), XLD_DATAENTRADA)
            End If
            
            'Grava distribuição p/ parcelas múltiplas
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
            " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
            
            Conexao.Execute Sql
            
            'Grava Duplicata da Nota de Parcelas Múltiplas
            Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                  " VALUES ('U'," & FunNuloData(DataVencimento, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & xNumFat & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
            
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
            If XTipoFatura <> 4 Then
                If ResFormulario!nofi_nr_fatura = 0 Then
                    xNumFat = FunPegaGuardaUltimo("F")
                Else
                    xNumFat = ResFormulario!nofi_nr_fatura
                End If
            Else
                xNumFat = 0
            End If
            
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
        
        If (chkPendente.Value = 1) Then
            ResFormulario!nofi_tx_pendestoque = "P"
        Else
            ResFormulario!nofi_tx_pendestoque = Null
        End If
            
        ResFormulario!nofi_tx_tipomov = Xtipomov
        ResFormulario!nofi_tx_status = XStatus
        ResFormulario!nofi_nr_documento = XNumDoc
        ResFormulario!focl_cd_forncli = CboFornecedor.BoundText
        ResFormulario!nofi_tx_serie = XSerie
        ResFormulario!tido_cd_TipoDoc = CboTipoDoc.BoundText
        ResFormulario!nofi_dt_entrada = DtpEntrada.Value
        ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
        ResFormulario!nofi_vl_valor = FunNuloVal(TxtValor.Text)
        ResFormulario!nofi_tx_distribuir = XDistribuido
        ResFormulario!nofi_nr_fatura = xNumFat
        ResFormulario!nofi_tx_tipofatura = XTipoFatura
        ResFormulario!empr_cd_empresa = PCodEmpresa
'        ResFormulario!nofi_nr_chamado = FunNuloVal(TxtChamado.Text)
        
        If (EntSaida = "S") And (PCliente = "QUALIDADOS") Then
            ResFormulario!plco_cd_contacredito = DatContaContabilCredito.Recordset.Fields("plco_cd_conta")
        End If
       
'        If chkSigiloso.Value = Checked Then
'            ResFormulario!nofi_bl_Sigiloso = 1
'        Else
'            ResFormulario!nofi_bl_Sigiloso = 0
'        End If
        
        ResFormulario.Update

        If XInserir = "A" And XStatus = "P" Then
           Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & xNumFat
        End If
        
        If XInserir = "I" Then
            Chave = ResFormulario!nofi_cd_notafiscal
        End If
    End If
        
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
    
    'Gera descrição de log da operação
    If XInserir = "I" Then
        'Nº de Doc e Nº de Fatura que foi gerado pelo sistema:
        XFM_MATRIZNOTAFISCAL(0, 1) = XNumDoc
        XFM_MATRIZNOTAFISCAL(1, 1) = xNumFat
        
        XFM_MATRIZNOTAFISCAL(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
        XFM_MATRIZNOTAFISCAL(3, 1) = TxtSerie.Text
        XFM_MATRIZNOTAFISCAL(4, 1) = CboStatus.Text
        XFM_MATRIZNOTAFISCAL(5, 1) = CboFornecedor.Text
        XFM_MATRIZNOTAFISCAL(6, 1) = CboTipoDoc.Text
        XFM_MATRIZNOTAFISCAL(7, 1) = Format$(DtpEntrada.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(8, 1) = Format$(DtpEmissao.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(9, 1) = FunNulo(CboHistorico.Text)
        
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, CADASTRAR, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, CADASTRAR, NOTAFISCAL_RECEBER)
    Else
        'Número de fatura original:
        XFM_MATRIZNOTAFISCAL(1, 1) = ResFormulario!nofi_nr_fatura
        XFM_MATRIZPAGAMENTO(0, 1) = ResFormulario!nofi_nr_fatura
        
        'Novos dados da nota:
        XFM_MATRIZNOTAFISCAL(0, 2) = MskNumdocumento.Text
        XFM_MATRIZNOTAFISCAL(2, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
        XFM_MATRIZNOTAFISCAL(3, 2) = TxtSerie.Text
        XFM_MATRIZNOTAFISCAL(4, 2) = CboStatus.Text
        XFM_MATRIZNOTAFISCAL(5, 2) = CboFornecedor.Text
        XFM_MATRIZNOTAFISCAL(6, 2) = CboTipoDoc.Text
        XFM_MATRIZNOTAFISCAL(7, 2) = Format$(DtpEntrada.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(8, 2) = Format$(DtpEmissao.Value, "DD/MM/YY")
        XFM_MATRIZNOTAFISCAL(9, 2) = FunNulo(CboHistorico.Text)
               
        If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, ALTERAR, NOTAFISCAL_PAGAR)
        If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZNOTAFISCAL, ALTERAR, NOTAFISCAL_RECEBER)
    End If
    
    'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
    XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
    XFT_CONTADORLOG = XFT_CONTADORLOG + 1
    
    '**********************FIM REGISTRA LOG - PARTE 2 ***************************
        
    ResFormulario.Close
                
    'Descontos
    If EntrouDesc And XStatus <> "P" And Tipo = "0" Then
        If XInserir = "A" Then
              Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & xNumFat
        End If
        
        If Vetor1(0, 0) <> Empty Then
            subGravaVetor "DescontosNF", "deen_dt_prevrecolhimento,deen_vl_valor,desc_cd_desconto, deen_dt_recolhimento, deen_nr_DarfDam, deen_nr_fatura", Array("D", "N", "N", "D", "T", "N"), 5, Array(1, 2, 3, 5, 6), "" & xNumFat & "", Vetor1, TDBGrid1
        End If
    End If
        
    '***************CONTABILIDADE*******************
    SubQOpenRecordset ResFormulario, "SELECT dupl_vl_Valor,dupl_tx_status,dupl_cd_Duplicata, " & _
            " dupl_tx_DocPagRec, dupl_tx_tipopagrec, dupl_dt_PagRec, coco_cd_codigo, dupl_tx_estorno FROM Duplicatas " & _
            " WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
                
    While Not ResFormulario.EOF
        If ResFormulario!dupl_tx_status = "R" And (IsNull(ResFormulario!dupl_tx_Estorno) Or ResFormulario!dupl_tx_Estorno = "N" Or ResFormulario!dupl_tx_Estorno = "") Then
            'Como o código da duplicata será apagado qdo da gravacao da duplicata, foi necessário apagar os registros na tabela de lançamento para gravar novos no futuro
            
            'Início - Larissa Moura 02/06/2008
            
            'Verifica se a nota é uma Prestação de Contas, pois ela não possui conta corrente.
            SubQOpenRecordset XLT_TIPODOC, "SELECT tido_tx_PrestContas FROM TipoDocumentos WHERE tido_cd_TipoDoc = " & CboTipoDoc.BoundText, Estatico
            If XLT_TIPODOC!tido_tx_PrestContas <> "1" Then
            
            'Fim - Larissa Moura 02/06/2008
                    XLT_SQL = "SELECT duplicatas.dupl_cd_Duplicata FROM (tipodocumentos " & _
                                       " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                       " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                       " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                       " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura " & _
                                       " WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                       " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                       " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N' or dupl_tx_estorno = '')"
            Else
                    XLT_SQL = "SELECT duplicatas.dupl_cd_Duplicata FROM (tipodocumentos " & _
                                      " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                      " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                      " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                      " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura " & _
                                      " WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo is null " & _
                                      " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                      " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N' or dupl_tx_estorno = '')"

            End If
            
            XLT_TIPODOC.Close
            SubQOpenRecordset RESFAT2, XLT_SQL, Estatico
            
            While Not RESFAT2.EOF
                    Sql = "DELETE FROM Lancamentos " & _
                        "WHERE lanc_cd_OrigemFin = " & RESFAT2!dupl_cd_Duplicata & " AND " & _
                                "lanc_tx_Origem = 'FD' AND " & _
                                "empr_cd_empresa = " & PCodEmpresa
                    Conexao.Execute Sql
                RESFAT2.MoveNext
            Wend
        End If
        ResFormulario.MoveNext
    Wend
    ResFormulario.Close
    
    '***********************************************
    
    'Duplicatas
    If Not ParcelasMultiplas Then
        If EntrouDupl Then 'composição de duplicata
            If XInserir = "A" Then
                If Vetor4(0, 0) <> Empty Then
                    GravaVetorDuplDescAcresc
                Else
                    Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & xNumFat & " and (dupl_tx_estorno <> 'S' or dupl_tx_estorno is null)" 'Não excluir os lançamentos de estorno
                    
                    'Exclui os versos de cheques que existem associados a este cheque
                    Conexao.Execute "DELETE FROM VersoCheques where vech_tx_numcheque= '" & FunNulo(XLT_DOCPAGRECVERSO) & "' and coco_cd_codigo = " & XLT_CONTAVERSO
                    
                    subGravaVetorDupl "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_estorno,dupl_tx_Consolidado,dupl_dt_concilia,dupl_dt_ProcPagamento,fopa_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "TN", "T", "D", "D", "N", "T"), 15, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13, 16, 17, 18, 19, 20), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                End If
            End If
            If XInserir = "I" Then
                If Vetor2(0, 0) <> Empty Then
                    subGravaVetorDupl "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_estorno,dupl_tx_Consolidado,dupl_dt_concilia,dupl_dt_ProcPagamento,fopa_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "N", "T", "T", "T", "D", "N", "N", "TN", "T", "D", "D", "N", "T"), 15, Array(0, 1, 3, 4, 5, 7, 9, 10, 11, 13, 16, 17, 18, 19, 20), Str(xNumFat) & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
                Else
                    'O usuário pode ter entrado na orelha de duplicatas, mas ter voltado e decidido fazer parcela única
                    If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente ou tem status = previsto
                         Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec,fopa_cd_codigo)" & _
                               " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "," & FunNuloValor(CboFormaPagamento1) & ")"
                         Conexao.Execute Sql
                    End If
                End If
            End If
        Else 'geração automática de duplicata
            If XInserir = "I" Then
               If Tipo = "1" Or XStatus = "P" Then 'Gera duplicata automaticamente ou tem status = previsto
                    Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec,fopa_cd_codigo)" & _
                          " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "," & FunNuloValor(CboFormaPagamento1) & ")"
                    Conexao.Execute Sql
                    
                    'Dados do pagamento
                    XFM_MATRIZPAGAMENTO(0, 1) = Str(xNumFat)
                    XFM_MATRIZPAGAMENTO(1, 1) = "U"
                    XFM_MATRIZPAGAMENTO(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                    XFM_MATRIZPAGAMENTO(3, 1) = CboStatus.Text
                    XFM_MATRIZPAGAMENTO(4, 1) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
                    XFM_MATRIZPAGAMENTO(5, 1) = CboTipoPagto1.Text
                    XFM_MATRIZPAGAMENTO(6, 1) = CboCCorrente.Text
                    XFM_MATRIZPAGAMENTO(7, 1) = FunNulo(CboBancoForn1.Text)
                    
                    If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                        XFM_MATRIZPAGAMENTO(8, 1) = TxtDocPagto1.Text
                        XFM_MATRIZPAGAMENTO(9, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
                    End If
               End If
            ElseIf XInserir = "A" And (Tipo = "1" Or XStatus = "P") Then
                'A rotina foi modificada para atender às alterações de Dupl que possuam Desc/Acresc
                Conexao.Execute "UPDATE Duplicatas " & _
                                    "SET dupl_nr_duplicata ='" & XCodDupl & "', " & _
                                    "dupl_dt_vencimento = " & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & _
                                    "dupl_vl_GlosaDevolucao = " & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & _
                                    "dupl_vl_valor = " & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & "," & _
                                    "dupl_tx_status = '" & XStatus & "'," & _
                                    "dupl_tx_tipopagrec = '" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "'," & _
                                    "dupl_tx_docpagrec = '" & FunNulo(XLT_DOCPAGREC) & "'," & _
                                    "coco_cd_codigo= " & FunNuloValor(CboCCorrente) & "," & _
                                    "bafo_cd_codigo = " & FunNuloValor(CboBancoForn1) & "," & _
                                    "dupl_tx_debcred ='" & Xdebcred & "'," & _
                                    "dupl_nr_fatura = " & Str(xNumFat) & "," & _
                                    "dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "," & _
                                    "fopa_cd_codigo = " & FunNuloValor(CboFormaPagamento1) & " " & _
                                    "WHERE dupl_nr_fatura = " & xNumFat & " AND (dupl_tx_estorno is null  OR  dupl_tx_estorno = '' OR dupl_tx_estorno = 'N')", XLB_REGISTROAFETADO
                
                If Not XLB_REGISTROAFETADO Then
                    Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_GlosaDevolucao,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec)" & _
                          " VALUES ('" & XCodDupl & "'," & FunNuloData(DtpVencParcelaUnica.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & FunNuloVal(FunTrataFloat(LblValorLiquido1.Caption)) & ",'" & XStatus & "','" & CStr(CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)) & "','" & FunNulo(XLT_DOCPAGREC) & "'," & FunNuloValor(CboCCorrente) & "," & FunNuloValor(CboBancoForn1) & ",'" & Xdebcred & "'," & Str(xNumFat) & "," & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & ")"
                    Conexao.Execute Sql
                End If
            End If
            
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
            'Gera descrição de log da operação
            If XInserir = "I" Then
                If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, CADASTRAR_DUPLICATA, NOTAFISCAL_PAGAR)
                If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, CADASTRAR_DUPLICATA, NOTAFISCAL_RECEBER)
            Else
                'Novos dados do pagamento:
                XFM_MATRIZPAGAMENTO(0, 2) = Str(xNumFat)
                XFM_MATRIZPAGAMENTO(1, 2) = "U"
                XFM_MATRIZPAGAMENTO(2, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                XFM_MATRIZPAGAMENTO(3, 2) = CboStatus.Text
                XFM_MATRIZPAGAMENTO(4, 2) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
                XFM_MATRIZPAGAMENTO(5, 2) = CboTipoPagto1.Text
                XFM_MATRIZPAGAMENTO(6, 2) = CboCCorrente.Text
                XFM_MATRIZPAGAMENTO(7, 2) = FunNulo(CboBancoForn1.Text)
                                           
                If UCase(CboStatusduplicata.List(CboStatusduplicata.ListIndex)) = "REALIZADO" Then
                    XFM_MATRIZPAGAMENTO(8, 1) = TxtDocPagto1.Text
                    XFM_MATRIZPAGAMENTO(9, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
                Else
                    XFM_MATRIZPAGAMENTO(8, 0) = ""
                    XFM_MATRIZPAGAMENTO(9, 0) = ""
                End If
                                           
                If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, ALTERAR_DUPLICATA, NOTAFISCAL_PAGAR)
                If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZPAGAMENTO, ALTERAR_DUPLICATA, NOTAFISCAL_RECEBER)
            End If
        
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '******************FIM REGISTRA LOG DA OPERAÇÃO - PARTE 2 *******************
        End If
    End If
    
    'Distribuição
    If EntrouDist Then
        If XInserir = "A" Then
              Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
        End If
        
        If Vetor3(0, 0) <> Empty Then
            subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N", "N"), 5, Array(0, 2, 5, 6, 8), "" & Chave & "", Vetor3, TDBGrid3
        Else
            'se entrou na orelha de distribuicao excluiu o rateio e cadastrou na orelha de apropriação
            Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
                " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
            Conexao.Execute Sql
        End If
    Else
        If ParcelasMultiplas = False Then
            If XInserir = "A" Then
                Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
            End If
            
            If ChkDistribuicao Then
                If Vetor3(0, 0) <> Empty Then
                    subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N", "N"), 5, Array(0, 2, 5, 6, 8), "" & Chave & "", Vetor3, TDBGrid3
                End If
            Else
                            
                Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor,dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
                      " VALUES ('" & DatPCR.Recordset.Fields("cere_cd_pcr") & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & FunNuloVal(FunTrataFloat((TxtValor.Text - CDbl(TxtDevolucao.Text)))) & "," & FunNuloVal(FunTrataFloat(TxtDevolucao.Text)) & "," & Chave & ", " & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
                Conexao.Execute Sql
                
                'Dados da apropriação
                XFM_MATRIZAPROPRIACAO(0, 1) = XNumDoc
                XFM_MATRIZAPROPRIACAO(1, 1) = CboPcr.Text
                XFM_MATRIZAPROPRIACAO(2, 1) = CboPlanoContas.Text
                XFM_MATRIZAPROPRIACAO(3, 1) = "100%"
                XFM_MATRIZAPROPRIACAO(4, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                XFM_MATRIZAPROPRIACAO(5, 1) = CboIdentificador2.Text
                
                '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2 ***********************
                
                'Gera descrição de log da operação
                If XInserir = "I" Then
                    If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
                    If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, CADASTRAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
                Else
                    'Novos dados da apropriação:
                    XFM_MATRIZAPROPRIACAO(0, 2) = MskNumdocumento.Text
                    XFM_MATRIZAPROPRIACAO(1, 2) = CboPcr.Text
                    XFM_MATRIZAPROPRIACAO(2, 2) = CboPlanoContas.Text
                    XFM_MATRIZAPROPRIACAO(3, 2) = "100%"
                    XFM_MATRIZAPROPRIACAO(4, 2) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
                    XFM_MATRIZAPROPRIACAO(5, 2) = FunNulo(CboIdentificador2.Text)
                                               
                    If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, ALTERAR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
                    If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XFM_MATRIZAPROPRIACAO, ALTERAR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
                End If
            
                'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
                XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
                XFT_CONTADORLOG = XFT_CONTADORLOG + 1
                
                '******************FIM REGISTRA LOG DA OPERAÇÃO - PARTE 2 *******************
                
            End If
        End If
    End If
    
    'Verificar se é preciso atualizar composição de fatura
    If XValorBruto <> CDbl(TxtValor.Text) And XTipoFatura = "3" And XInserir = "A" Then
        If MsgBox("É necessário atualizar a Composição da Fatura. Deseja atualizar agora ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
            'Chamar tela de composição de fatura
            NFatura = xNumFat
            XFormulario = "FrmNotasFiscais"
            FrmFatura.Show 1
        End If
    End If
    
    '***************CONTABILIDADE*******************
    
    If DatTipoDoc.Recordset.Fields("tido_tx_NEC") = "0" Then
        If XGT_TIPOEXPORTACAO = "PC" And _
            (CboStatus.Text <> "Previsto" And _
            (IsNull(DatTipoDoc.Recordset.Fields("tido_tx_sigla")) Or _
            (DatTipoDoc.Recordset.Fields("tido_tx_adiantamento") <> "1" And _
            DatTipoDoc.Recordset.Fields("tido_tx_sigla") <> "NP"))) Then  'Exporta base provisão
            
            'Parte do Caixa para registros provisionados
            SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata, dupl_dt_pagrec, " & _
                            " dupl_tx_Status, dupl_tx_TipoPagRec, coco_cd_codigo, " & _
                            " dupl_tx_docpagrec, dupl_vl_valor, dupl_tx_estorno, dupl_tx_debcred " & _
                            " FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
            
            If XFB_PROVISIONA And xNumFat <> 0 Then
                'xNumFat <> 0 significa que não tem parcela e a fatura será composta futuramente
                'nesses casos não é possível provisionar
'                SubProvisao xNumFat, DtpEntrada.Value
            ElseIf Not XFB_PROVISIONA And XInserir = "A" Then
                'Apagar os Lançamentos contábeis relacionados com a provisão desta NF, caso exista
                Conexao.Execute "DELETE FROM Lancamentos " & _
                                "WHERE lanc_cd_OrigemFin = " & Chave & " AND " & _
                                    " (lanc_tx_Origem = 'FN') AND " & _
                                    "empr_cd_empresa = " & PCodEmpresa
            End If
            
            While Not ResFormulario.EOF
                'Em se tratando de parcela, só contabiliza qdo. realizado
                'E quando for estorno só considerar o valor > 0
                If ResFormulario!dupl_tx_status = "R" And ResFormulario!dupl_vl_valor > 0 Then
                    'se a duplicata está sendo cadastrada como realizada ou se alterou para realizada,
                    'Se alterou, deverá ser verificado se foi consolidado.
                    'se a duplicata já era realizada e está sendo alterada
                    
                    'RECONTABILIZAR O ESTORNO SOMENTE SE A NOTA QUE ESTÁ SENDO ALTERADA POSSUIR UM
                    If ResFormulario!dupl_tx_Estorno = "S" Then
                    Else
                        'VERIFICA SE FOI CONSOLIDADO
                          XLT_SQL = "SELECT Count(*) as ContReg FROM (tipodocumentos " & _
                                    " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                    " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                    " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                    " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                    " AND dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & " " & _
                                    " AND (dupl_tx_estorno is null or dupl_tx_estorno = 'N'  or dupl_tx_estorno = '')"
    
                        SubQOpenRecordset ResFat, XLT_SQL, Estatico
                        
                        If ResFat!ContReg > 1 Then
                            'SE ESTÁ ALTERANDO UM CONSOLIDADO NÃO DEVE APAGAR NEM RECADASTRAR A PARTE DO VALOR CONSOLIDADO
'                            SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, True, False
                        Else
'                            SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, False, False
                        End If
                    End If
                End If
                ResFormulario.MoveNext
            Wend
            ResFormulario.Close
        Else 'Exporta base caixa
            SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata, dupl_dt_pagrec, " & _
                        " dupl_tx_Status, dupl_tx_TipoPagRec, coco_cd_codigo, dupl_tx_docpagrec " & _
                        " FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
            
            While Not ResFormulario.EOF
                If ResFormulario!dupl_tx_status = "R" Then 'Em se tratando de parcela, só contabiliza qdo. realizado
                    'se a duplicata está sendo cadastrada como realizada ou se alterou para realizada,
                    'Se alterou, deverá ser verificado se foi consolidado.
                    'se a duplicata já era realizada e está sendo alterada

                    XLT_SQL = "SELECT Count(*) as ContReg FROM (tipodocumentos " & _
                                " RIGHT JOIN notasfiscais ON tipodocumentos.tido_cd_tipodoc = notasfiscais.tido_cd_tipodoc) " & _
                                " RIGHT JOIN (duplicatas LEFT JOIN consfinrelfinanceirodescacresc ON " & _
                                " duplicatas.dupl_cd_duplicata = consfinrelfinanceirodescacresc.dupl_cd_duplicata) ON " & _
                                " notasfiscais.nofi_nr_fatura = duplicatas.dupl_nr_fatura WHERE dupl_tx_tipopagrec ='" & ResFormulario!dupl_tx_tipopagrec & "' and coco_cd_codigo = " & ResFormulario!coco_cd_codigo & " " & _
                                " and dupl_tx_docpagrec= '" & ResFormulario!dupl_tx_docpagrec & "' and tido_tx_nec='0' and dupl_dt_PagRec = " & FunNuloData(ResFormulario!dupl_dt_pagrec, NomeSgbd) & ""

                    SubQOpenRecordset ResFat, XLT_SQL, Estatico
                    If ResFat!ContReg > 1 Then
                        'SE ESTÁ ALTERANDO UM CONSOLIDADO NÃO DEVE APAGAR NEM RECADASTRAR
                        'A PARTE DO VALOR CONSOLIDADO
'                        SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, True, False
                    Else
'                        SubContabilizarParcela ResFormulario!dupl_cd_Duplicata, ResFormulario!dupl_dt_pagrec, False, False
                    End If
                End If
                ResFormulario.MoveNext
            Wend
            ResFormulario.Close
        End If
    End If
    
    '***********************************************
    
    Conexao.CommitTrans
        
    '***************************** GRAVA LOG NO BANCO **********************************
    Dim aux As Integer
    Dim inicioLog As String
    Dim fimLog As String
    
    For Cont = 0 To (XFT_CONTADORLOG - 1)
        XFT_LOG = XFV_VETORLOG(Cont)
        
        'No caso de cadastro de nota fiscal, insere o número da nota fiscal gerado
        If XInserir = "I" Then
            aux = InStr(1, XFT_LOG, "NOTA FISCAL:", vbTextCompare)
            If aux > 0 Then
                aux = aux + 12
                inicioLog = Mid(XFT_LOG, 1, aux) + XNumDoc
                fimLog = Mid(XFT_LOG, aux, Len(XFT_LOG) - aux)
                XFT_LOG = inicioLog + fimLog
            End If
        End If
        
        If EntSaida = "E" Then Call subRegistraLog("TelaNotasFiscais", "2", XFT_LOG)
        If EntSaida = "S" Then Call subRegistraLog("TelaNotasFiscais", "1", XFT_LOG)
    Next
    
    '************************** FIM GRAVA LOG NO BANCO **********************************
    
    '////Imprime processo de pagamento/////
    If XInserir = "I" And XTipoFatura <> "4" And EntSaida = "S" Then
        If XGT_IMPRIMEPP = "S" Then
            If MsgBox("Deseja Imprimir Processo de Pagamento?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
'                subImprimePP xNumFat, False
            End If
        End If
    End If
    
    If XInserir = "A" Then
        If cmdNovo.Enabled = False Then
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Call CmdDesistir_Click
        Else
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = False '26/11/09 - Patrícia
            Call CmdDesistir_Click
        End If
    Else
        XInserir = "A"
        NumDocAnt = MskNumdocumento.Text
        cmdNovo.Enabled = True
        
        If Not EntrouDesc Then
            Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
            subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura", "deen_dt_recolhimento", "deen_nr_DarfDam"), Vetor1, TDBGrid1
        End If
        
        Sql = "SELECT DuplDescAcresc.* from DuplDescAcresc LEFT JOIN Duplicatas ON Duplicatas.dupl_cd_duplicata = DuplDescAcresc.dupl_cd_duplicata Where Duplicatas.dupl_nr_fatura = " & NFatura
        CarregaVetorDescAcresc Sql, Array("dupl_cd_duplicata", "deac_cd_descacresc", "dude_vl_descacresc"), Vetor4, 3
            
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = False 'Daniel Coelho 09/02/2012
    End If
        
    XFT_NOVO = 0 '13/11/09 - Patrícia
    
''''    CmdGravar.Enabled = False '13/11/09 - Patrícia
''''    If XInserir = "A" Then
''''        CmdDesistir_Click
''''    End If

    Exit Sub
    Unload Me
    
TrataErro:
    'MsgBox Err.Description, vbCritical, "Erro"
End Sub

Sub GravaVetorDuplDescAcresc()
    Dim ResFormulario As Object
    Dim x As Integer, y As Integer
    Dim AchouDupl As Boolean
              
    'Quando excluir duplicatas do grid, é preciso excluí-las da tabela.
    'quando o código de uma duplicata ainda estiver no vetor, esta não deverá ser apagada.
    SubQOpenRecordset ResFormulario, "SELECT dupl_cd_Duplicata FROM Duplicatas " & _
                " WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
    While Not ResFormulario.EOF
        AchouDupl = False
        For x = 0 To Vetor2.UpperBound(1)
            If Vetor2(x, 15) <> Empty Then
                If ResFormulario!dupl_cd_Duplicata = CLng(Vetor2(x, 15)) Then
                    AchouDupl = True
                    x = Vetor2.UpperBound(1)
                End If
            End If
        Next
        
        If AchouDupl = False Then
            Conexao.Execute "DELETE from Duplicatas WHERE dupl_cd_duplicata = " & ResFormulario!dupl_cd_Duplicata
        End If
        
        ResFormulario.MoveNext
    Wend
    ResFormulario.Close
         
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
         
         If Vetor2(x, 3) <> Empty Then
            ResFormulario!dupl_vl_GlosaDevolucao = Vetor2(x, 3)
         Else
            ResFormulario!dupl_vl_GlosaDevolucao = 0
         End If
         
         ResFormulario!dupl_vl_valor = Vetor2(x, 4)
         ResFormulario!dupl_tx_status = Vetor2(x, 5)
         ResFormulario!dupl_tx_tipopagrec = Vetor2(x, 7)
         ResFormulario!dupl_tx_docpagrec = Vetor2(x, 9)
         
         If Not IsDate(Vetor2(x, 10)) Then
             ResFormulario!dupl_dt_pagrec = Null
         Else
             ResFormulario!dupl_dt_pagrec = Format$(Vetor2(x, 10), "DD/MM/YY")
         End If
         
         ResFormulario!coco_cd_codigo = Vetor2(x, 11)
         ResFormulario!bafo_cd_codigo = funNull(Vetor2(x, 13))
         ResFormulario!dupl_nr_fatura = xNumFat
         ResFormulario!dupl_tx_debcred = Xdebcred
                     
         ResFormulario.Update
    Next
    ResFormulario.Close
    
End Sub

Function ConsistenciaDatas() As Boolean
'Verifica se as datas seguem a ordem de preenchimento
'Se ConsistenciaDatas= TRUE    >> Datas consistentes
'Se ConsistenciaDatas= FALSE  >> Datas não consistentes
    ConsistenciaDatas = False
    If CboStatus.ListIndex <> 0 Then
        If TabNotasFiscais.Tab = 0 Then
            If Not IsNull(DtpVencParcelaUnica.Value) And Not IsNull(DtpEmissao.Value) Then
                If DateDiff("d", DtpEmissao.Value, DtpVencParcelaUnica.Value) < 0 Then
                    MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    DtpVencParcelaUnica.SetFocus
                    Exit Function
                End If
            End If
'            If Not IsNull(DtpPagtoParcelaUnica.Value) Then
'                If CboStatus.ListIndex = 2 Then
'                    If DateDiff("d", DtpEmissao.Value, DtpPagtoParcelaUnica.Value) < 0 Then
'                        MsgBox "A data de pagamento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
'                        DtpPagtoParcelaUnica.SetFocus
'                        Exit Function
'                    End If,
'                End If
'            End If
        ElseIf TabNotasFiscais.Tab = 2 Then
           If DateDiff("d", DtpEmissao.Value, DtpVencDesconto.Value) < 0 Then
                MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                If DtpVencDesconto.Enabled = True Then DtpVencDesconto.SetFocus
                Exit Function
           End If
        ElseIf TabNotasFiscais.Tab = 1 Then
           If Not IsNull(DtpVencimento.Value) Then
                If DateDiff("d", DtpEmissao.Value, DtpVencimento.Value) < 0 Then
                    MsgBox "A data de vencimento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                    If DtpVencimento.Enabled = True Then DtpVencimento.SetFocus
                    Exit Function
                End If
           End If
           
           If CboStatus.ListIndex = 2 Then
                If Not IsNull(DtpPagamento.Value) Then
                    If DateDiff("d", DtpEmissao.Value, DtpPagamento.Value) < 0 Then
                        MsgBox "A data de pagamento é anterior a data de emissão !", vbCritical, "ATENÇÃO"
                        If DtpPagamento.Enabled = True Then DtpPagamento.SetFocus
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

Private Sub CmdImportarPlanilha_Click()

On Error GoTo TrataErro

    Dim XLO_EXCEL As Excel.Application
    Dim XLO_WB As Excel.Workbook
    Dim XLO_WS As Excel.Worksheet
    Dim XLO_RS As ADODB.Recordset
    
    Dim XLT_CAMINHOARQUIVOEXCEL As String
    Dim XLI_LINHAARQUIVOEXCEL As Integer

    Dim CR_CODIGO As Integer
    Dim CR_ESTRUTURADO As String
    Dim CC_CODIGO As Integer
    Dim CC_REDUZIDO As String
    Dim CC_ESTRUTURADO As String
    Dim VALOR_DISTRIBUICAO As Double
    Dim PERCENT_DISTRIBUICAO As Double
    Dim TOTAL_DISTRIBUICAO As Double
    Dim Achou As Boolean
    
    Dim XLB_FLAGIMPORTAR As Boolean
    Dim XLI_REGISTROSIMPORTADOS As Integer

    Me.MousePointer = vbHourglass
    
    XLB_FLAGIMPORTAR = True
    
    XLI_REGISTROSIMPORTADOS = 0
    
    TOTAL_DISTRIBUICAO = 0
            
    With CommonDialog1
        .FileName = ""
        .Filter = "Arquivo Microsoft Excel (*.xls) |*.xls|" 'Sets the filter
        .ShowOpen
    End With
    
    XLT_CAMINHOARQUIVOEXCEL = CommonDialog1.FileName
    
    If XLT_CAMINHOARQUIVOEXCEL <> "" Then
                
        Set XLO_EXCEL = New Excel.Application
        
        Set XLO_WB = XLO_EXCEL.Workbooks.Open(XLT_CAMINHOARQUIVOEXCEL)
             
        If Not XLO_EXCEL.Worksheets(1) Is Nothing Then
                    
            XLO_WB.Sheets(1).Select
            XLI_LINHAARQUIVOEXCEL = 1
            
            
            'Percorre as linhas da planilha para capturar os valores
            'A soma destes valores (TOTAL_DISTRIBUICAO) será utilizado para realizar os testes das operações seguintes.
            While XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 3) <> ""
            
                  VALOR_DISTRIBUICAO = XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 3)
                  TOTAL_DISTRIBUICAO = TOTAL_DISTRIBUICAO + VALOR_DISTRIBUICAO
                  
                  XLI_LINHAARQUIVOEXCEL = XLI_LINHAARQUIVOEXCEL + 1
                  
            Wend
            
            'Verifica se a soma das distribuições NÃO é maior do que o valor da nota
            'Se for maior, ele encerra a operação e informa ao usuário
            If TOTAL_DISTRIBUICAO > CDbl(TxtValor.Text) Then
            
                 MsgBox "A soma das distribuições é maior que o total da nota.", vbInformation, "ATENÇÃO"
                 Me.MousePointer = vbDefault
                 Exit Sub
                 
            End If
            
            'Verifica se a soma das distribuições é maior que zero
            'Se não houver valor, ele encerra a operação e informa ao usuário
            If TOTAL_DISTRIBUICAO = 0 Then
            
                 MsgBox "Não foram definidos valores para as distribuições.", vbInformation, "ATENÇÃO"
                 Me.MousePointer = vbDefault
                 Exit Sub
                 
            End If
                    
            XLI_LINHAARQUIVOEXCEL = 1
        
            'Percorre as linhas da planilha para capturar as informações
            While XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 1) <> "" And _
                    XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 2) <> "" And _
                    XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 3) <> ""
                    XLB_FLAGIMPORTAR = True
                    
                    CR_CODIGO = 0
                    CC_CODIGO = 0
                    
                    'Novos dados da distribuição informados pelo usuário
                    CR_ESTRUTURADO = XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 1)
                    CC_REDUZIDO = XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 2)
                    VALOR_DISTRIBUICAO = XLO_WB.Application.Cells(XLI_LINHAARQUIVOEXCEL, 3)
                             
                    If Not DatPCR.Recordset.EOF Then
                        DatPCR.Recordset.MoveFirst
                    End If
                    
                    Achou = False
                    
                    'Verifica se existe o CÓDIGO ESTRUTURADO do centro de resultados informado pelo usuário
                    'Se existir, ele captura o CÓDIGO (ID) deste centro de resultados
                    While Not DatPCR.Recordset.EOF And Achou = False
                    
                        If DatPCR.Recordset.Fields("cere_cd_Estruturado").Value = CR_ESTRUTURADO Then
                        
                            CR_CODIGO = DatPCR.Recordset.Fields("cere_cd_Pcr").Value
                            Achou = True
                        
                        End If
                    
                        DatPCR.Recordset.MoveNext
                    
                    Wend
                    
                    DatPCR.Recordset.MoveFirst
                    
                    If Not DatConta.Recordset.EOF Then
                        DatConta.Recordset.MoveFirst
                    End If
                    
                    Achou = False
                    
                    'Verifica se existe o CÓDIGO REDUZIDO da conta contábil informado pelo usuário
                    'Se existir, ele captura o CÓDIGO (ID) desta conta contábil
                    While Not DatConta.Recordset.EOF And Achou = False
                    
                        If DatConta.Recordset.Fields("plco_cd_CodReduzido").Value = CC_REDUZIDO Then
                        
                            CC_CODIGO = DatConta.Recordset.Fields("plco_cd_Conta").Value
                            Achou = True
                        
                        End If
                    
                        DatConta.Recordset.MoveNext
                    
                    Wend
                    
                    DatConta.Recordset.MoveFirst
                    
                    If CC_CODIGO > 0 And CR_CODIGO > 0 And VALOR_DISTRIBUICAO > 0 Then
                       PERCENT_DISTRIBUICAO = VALOR_DISTRIBUICAO / CDbl(TxtValor.Text)
                        
                       If Not ExisteDistribuicao(2, Array(0, 2), Array(CR_CODIGO, CC_CODIGO), Vetor3, TDBGrid3) Then
                        'Insere o registro da distribuição no grid
                        subInsereRegistroVetor Array(CR_CODIGO, CR_ESTRUTURADO, CC_CODIGO, CC_REDUZIDO, VALOR_DISTRIBUICAO, 0, VALOR_DISTRIBUICAO, PERCENT_DISTRIBUICAO, Null, ""), Vetor3, TDBGrid3
                       End If
                       
                    End If
                    XLI_LINHAARQUIVOEXCEL = XLI_LINHAARQUIVOEXCEL + 1

            Wend
         End If
         
         LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
         ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
         XLO_EXCEL.Workbooks.Close
         XLO_EXCEL.Quit
         Me.MousePointer = vbDefault
         Exit Sub
         
    End If
                                
TrataErro:
    If Not XLO_WB Is Nothing Then
        XLO_WB.Close False
    End If
    
    If Not XLO_EXCEL Is Nothing Then
        Set XLO_EXCEL = Nothing
    End If
    funTrataErros (ComMensagem)
    
End Sub

Private Sub CmdInsereDesconto_Click()
    If TabNotasFiscais.Tab = 2 Then
        FrmNotasFiscais.MousePointer = vbHourglass
        'If ExisteChave(1, Array("4"), Array(Fatura), Vetor1, TDBGrid1) Then Exit Sub
        ChaveD = -1
        XDesconto = "I"
        PrepararInsercao
'        pandescontos.Visible = True
'        pandescontos.Enabled = True
        
        CboDescontos.SetFocus
        DesabilitaBotoes
        FrmNotasFiscais.MousePointer = vbDefault
    End If
End Sub

Private Sub CmdInsereDistribuicao_Click()
    If TabNotasFiscais.Tab = 3 Then
        FrmNotasFiscais.MousePointer = vbHourglass
        XDistribuicao = "I"
        PrepararInsercao
        If LblValorNota.Caption = LblValorTotalDistrib Then
           MsgBox "As distribuições já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        Else
'            PanDistribuicao.Visible = True
'            PanDistribuicao.Enabled = True
            MskPcr2.SetFocus
            DesabilitaBotoes
        End If
        
        FrmNotasFiscais.MousePointer = vbDefault
    End If
End Sub

Private Sub CmdInsereduplicata_Click()
    If TabNotasFiscais.Tab = 1 Then
        If Vetor2.UpperBound(1) >= 75 Then
           MsgBox "O número de máximo de parcelas é 75.", vbCritical, "ATENÇÃO"
           Exit Sub
        End If
        
        FrmNotasFiscais.MousePointer = vbHourglass
        XDuplicata = "I"
        ChaveDp = -1
        If (LblValorLiquido1.Caption + CDbl(TxtDevolucao.Text)) = LblValorTotal.Caption Then
           MsgBox "As duplicatas já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        ElseIf CDbl((LblValorLiquido1.Caption + CDbl(TxtDevolucao.Text)) - LblValorTotal.Caption) < 0 Then
            MsgBox "O valor da duplicata não pode ser < 0.", vbCritical, "ATENÇÃO"
        Else
'            PnlCompDuplicata.Visible = True
'            PnlCompDuplicata.Enabled = True
            TxtDocPagto2.Enabled = False
            DtpPagamento.Enabled = False
            
            PrepararInsercao
    
            'DesabilitaBotoes
            
            If TxtValorDuplicata.Enabled = True Then TxtValorDuplicata.SetFocus
        End If
        If CboStatus.Text = "Realizado" Or CboStatus.Text = "Cancelado" Then
            CboStatusduplicata.ListIndex = 1
            CboStatusduplicata.Enabled = False
        Else
            CboStatusduplicata.Enabled = True
        End If
    
        FrmNotasFiscais.MousePointer = vbDefault
    End If
End Sub


Private Sub cmdLimparCC_Click()
    MskConta.Text = ""
    CboPlanoContas.BoundText = 0
    CboPlanoContas.Text = ""
    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
End Sub

Private Sub CmdLimparCC2_Click()
    MskConta2.Text = ""
    CboPlanoContas2.BoundText = 0
    CboPlanoContas2.Text = ""
    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
End Sub

Private Sub CmdLimparCR_Click()
Dim XLO_CHECAR As ADODB.Recordset
MskPcr.Mask = ""
    MskPcr.Text = ""
    CboPcr.BoundText = 0
    CboPcr.Text = ""
    MskPcr.Mask = pPCRmascara
    
    SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & " and cere_cd_pcr= " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
    If Not (XLO_CHECAR.EOF) Then
            If MskConta.Text <> "" Then
                subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
            End If
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
    End If

    
    If EntSaida = "E" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
    ElseIf EntSaida = "S" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
    End If
End Sub

Private Sub CmdLimparCR2_Click()
Dim XLO_CHECAR As ADODB.Recordset
    MskPcr2.Mask = ""
    MskPcr2.Text = ""
    CboPcr2.BoundText = 0
    CboPcr2.Text = ""
    MskPcr2.Mask = pPCRmascara
    
    SubQOpenRecordset XLO_CHECAR, "SELECT plco_cd_Conta FROM CR_CC where plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & " and cere_cd_pcr= " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
    If Not (XLO_CHECAR.EOF) Then
            If MskConta2.Text <> "" Then
                subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr IN (SELECT cere_cd_Pcr FROM CR_CC WHERE plco_cd_Conta = " & DatConta.Recordset.Fields("plco_cd_Conta") & ")" & " ORDER BY cere_cd_estruturado", Estatico
            End If
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
    End If
    
    If EntSaida = "E" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
    ElseIf EntSaida = "S" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
    End If
End Sub

Private Sub CmdLimparHistorico_Click()
    CboHistorico.BoundText = 0
    CboHistorico.Text = ""
End Sub


Private Sub CmdLimparIdentCcredito_Click()
    MskcontaCredito.Text = ""
    cboPlanoContasCredito.BoundText = 0
    cboPlanoContasCredito.Text = ""
End Sub

Private Sub CmdLimparIdentif_Click()
    CboIdentificador2.BoundText = 0
    CboIdentificador2.Text = ""
End Sub

Private Sub cmdNovo_Click()
    If XFT_NOVO <> 1 Then
        CmdGravar.Enabled = True
        XFT_NOVO = 1
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
        TxtCodTipoDoc.Enabled = False
        ParcelasMultiplas = False
        XFB_ESTOQUE = False
        XFB_QUALIFAT = False
        XFT_INSPARCMULTIPLA = "I"
        CboFornecedor.Enabled = False
        MskNumdocumento.Enabled = False
        CmdEntMultiplas.Enabled = True
        DtpEntrada.Enabled = False
        DtpEmissao.Enabled = False
        TxtValor.Enabled = False
        CboHistorico.Enabled = False
        TxtHistorico.Enabled = False
        CmdLimparHistorico.Enabled = False
        LblValorLiquido1.Enabled = False
        ChkDistribuicao.Value = 0
    
        TabAuxiliar.Enabled = False
        
        chkPendente.Value = False
        ValorTotalDuplicata = 0
        ValorTotalDistribuicao = 0
        
        TabNotasFiscais.TabEnabled(1) = False
        TabNotasFiscais.TabEnabled(2) = False
        TabNotasFiscais.TabEnabled(3) = False
    
        EntrouDesc = False
        EntrouDupl = False
        EntrouDist = False
    
        cmdNovo.Enabled = False
        CboStatus.SetFocus
            
        '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
        'Modificado Por Daniel Coelho - 18/06/2012 -  antes XFV_VETORLOG(50)
        ReDim XFV_VETORLOG(100) As Variant 'Linha de código alterada em 29/01/09 - Patrícia. Anterior:ReDim XFV_VETORLOG(50) As Variant
        XFT_CONTADORLOG = 0
        
        'NOTA FISCAL
        ReDim XFM_MATRIZNOTAFISCAL(10, 2) As Variant
        
        XFM_MATRIZNOTAFISCAL(0, 0) = "Nº Doc"
        XFM_MATRIZNOTAFISCAL(1, 0) = "Num Fatura"
        XFM_MATRIZNOTAFISCAL(2, 0) = "Valor"
        XFM_MATRIZNOTAFISCAL(3, 0) = "Serie"
        XFM_MATRIZNOTAFISCAL(4, 0) = "Status"
        If EntSaida = "E" Then XFM_MATRIZNOTAFISCAL(5, 0) = "Cliente"
        If EntSaida = "S" Then XFM_MATRIZNOTAFISCAL(5, 0) = "Fornecedor"
        XFM_MATRIZNOTAFISCAL(6, 0) = "Tipo Doc."
        XFM_MATRIZNOTAFISCAL(7, 0) = "Data Entrada"
        XFM_MATRIZNOTAFISCAL(8, 0) = "Data Emissão"
        XFM_MATRIZNOTAFISCAL(9, 0) = "Historico"
        
        'PAGAMENTO
        ReDim XFM_MATRIZPAGAMENTO(10, 2) As Variant
                    
        XFM_MATRIZPAGAMENTO(0, 0) = "Num Fatura"
        XFM_MATRIZPAGAMENTO(1, 0) = "Cod"
        XFM_MATRIZPAGAMENTO(2, 0) = "Valor"
        XFM_MATRIZPAGAMENTO(3, 0) = "Status"
        XFM_MATRIZPAGAMENTO(4, 0) = "Data Venc"
        XFM_MATRIZPAGAMENTO(5, 0) = "Tipo Pagto"
        XFM_MATRIZPAGAMENTO(6, 0) = "Conta Corrente"
        If (CboBancoForn1.Enabled = True) Then XFM_MATRIZPAGAMENTO(7, 0) = "Banco Fornecedor"
        XFM_MATRIZPAGAMENTO(8, 0) = "Doc Pagto"
        XFM_MATRIZPAGAMENTO(9, 0) = "Data Pagto"
    
        'APROPRIAÇÃO
        ReDim XFM_MATRIZAPROPRIACAO(6, 2) As Variant
             
        XFM_MATRIZAPROPRIACAO(0, 0) = "Nota Fiscal"
        XFM_MATRIZAPROPRIACAO(1, 0) = "Centro Custo"
        XFM_MATRIZAPROPRIACAO(2, 0) = "Conta Contábil"
        XFM_MATRIZAPROPRIACAO(3, 0) = "Percentual"
        XFM_MATRIZAPROPRIACAO(4, 0) = "Valor"
        XFM_MATRIZAPROPRIACAO(5, 0) = "Identificador"
        
        '******************FIM REGISTRA LOG DA OPERAÇÃO - PARTE 1***********************
    End If
        
End Sub

Private Sub CmdRemoverDistribuicao_Click()
    Dim Houve_Erro As Integer

    If TabNotasFiscais.Tab = 3 Then
        Houve_Erro = False
       
        DesabilitaBotoes
       
        TDBGrid3.Enabled = False
       
        If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
            
            '*********REGISTRA LOG DA OPERAÇÃO - PARTE 1**********
            ReDim XGM_MATRIZLOG(6, 2) As Variant
        
            XGM_MATRIZLOG(0, 0) = "Nota Fiscal"
            XGM_MATRIZLOG(1, 0) = "Centro Custo"
            XGM_MATRIZLOG(2, 0) = "Conta Contábil"
            XGM_MATRIZLOG(3, 0) = "Percentual"
            XGM_MATRIZLOG(4, 0) = "Valor"
            XGM_MATRIZLOG(5, 0) = "Identificador"
            XGM_MATRIZLOG(0, 1) = MskNumdocumento
            XGM_MATRIZLOG(1, 1) = TDBGrid3.Columns("Centro Custo").Text
            XGM_MATRIZLOG(2, 1) = TDBGrid3.Columns("Conta Contábil").Text
            XGM_MATRIZLOG(3, 1) = TDBGrid3.Columns("Percentual").Text
            XGM_MATRIZLOG(4, 1) = Format$(TDBGrid3.Columns("Valor").Text, "##,##0.00")
            XGM_MATRIZLOG(5, 1) = TDBGrid3.Columns("Identificador").Text
            '****************FIM LOG - PARTE 1********************

            subRemoveRegistroVetor Vetor3, TDBGrid3
                        
            '******************REGISTRA LOG DA OPERAÇÃO - PARTE 2**************************
            'Gera descrição de log da operação
            If EntSaida = "S" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DISTRIBUICAO, NOTAFISCAL_PAGAR)
            If EntSaida = "E" Then XFT_LOG = funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR_DISTRIBUICAO, NOTAFISCAL_RECEBER)
                   
            'Guarda o log no vetor XFV_VETORLOG para que ele seja gravado na tabela de Log quando o usuário clicar em Gravar
            XFV_VETORLOG(XFT_CONTADORLOG) = XFT_LOG
            XFT_CONTADORLOG = XFT_CONTADORLOG + 1
            '*****************************FIM LOG - PARTE 2********************************
            
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
    End If
    
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
    Dim ResEstoque As Object

    PanDistribuicao.Visible = False
    PanDistribuicao.Enabled = False
       
    XDistribuicao = ""
    TDBGrid3.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
       If XFB_ESTOQUE = True Then ' A nota não é do estoque
            CmdAlteraDistribuicao.Enabled = True
       End If
    Else
       CmdInsereDistribuicao.Enabled = True
    End If
    
    HabilitaBotoes
    CmdGravar.Enabled = True
    
    MskPcr2.Mask = ""
    MskPcr2.Text = ""
    MskPcr2.Mask = pPCRmascara
    CboPcr2.Text = ""
    MskConta2.Mask = ""
    MskConta2.Text = ""
    CboPlanoContas2.Text = ""
    
'    'Inserido em 07/07/10 - Patrícia--------------------------
'    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
'
'    If EntSaida = "E" Then
'        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
'    ElseIf EntSaida = "S" Then
'        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
'    End If
'    '----------------------------------------------------------

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
    'Set Formulario = FrmNotasFiscais
    FrmNotasFiscais.DatFornecedor.Refresh
    FrmNotasFiscais.CboFornecedor.BoundText = XGV_FornNotaFiscal
    
    PanFornecedor.Visible = False
    PanFornecedor.Enabled = False
End Sub


Private Sub Verifica_DataEmissao()
    
        If pVinculaCrCc = True Then
        CboPlanoContas.Enabled = True '07/05/09 - Patrícia
        If (DtpEmissao) >= CDate(Data_VinculoCRCC) Then
            If CboPcr.Text <> "" Then
                DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
                
                If EntSaida = "E" Then
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCVinculadaCombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') and cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_Pcr") & " ORDER BY plco_tx_desccontabil", Estatico
                Else
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCVinculadaCombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') and cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_Pcr") & " ORDER BY plco_tx_desccontabil", Estatico
                End If
                
                If (MskConta.Text <> "") Then
                    DatConta.Recordset.Find ("plco_cd_CodReduzido = " & MskConta.Text)
                    
                    If DatConta.Recordset.EOF Then
                        MskConta.Mask = ""
                        MskConta.Text = ""
                        CboPlanoContas.Text = ""
                    End If
                End If
                
            ElseIf CboPcr2.Text <> "" Then
                DatPCR.Recordset.Bookmark = CboPcr2.SelectedItem
                
                If EntSaida = "E" Then
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCVinculadaCombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') and cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_Pcr") & " ORDER BY plco_tx_desccontabil", Estatico
                Else
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCVinculadaCombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') and cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_Pcr") & " ORDER BY plco_tx_desccontabil", Estatico
                End If
                
                If (MskConta2.Text <> "") Then
                    DatConta.Recordset.Find ("plco_cd_CodReduzido = " & MskConta2.Text)
                    
                    If DatConta.Recordset.EOF Then
                        MskConta2.Mask = ""
                        MskConta2.Text = ""
                        CboPlanoContas2.Text = ""
                    End If
                End If
                
            Else
                Exit Sub
            End If
        End If
        
        CboPlanoContas2.Enabled = True
    End If
End Sub

Private Sub Command4_Click()
    PainelCriterios.Visible = False
    PainelCriterios.Enabled = False
    CboCriterios.Text = ""
End Sub

Private Sub DtpEmissao_Click()
    Verifica_DataEmissao
End Sub

Private Sub Form_Activate()
    'Set Formulario = FrmNotasFiscais

    If (XGB_CONTABILIZA) Then
        Unload Me
    Else
        Reload = False
    End If
    
End Sub

Private Sub Form_Load()
    'Set Formulario = FrmNotasFiscais
    Data_EmpObr = "10/06/2008" ' Alteração em 09/06/08 - Patrícia
    XFT_NOVO = 0 '13/11/09 - Patrícia
    
'    If FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "12") Then
'        chkSigiloso.Visible = True
'    Else
'        chkSigiloso.Visible = False
'    End If
        
    Reload = False
    
    Dim ResFat As Object
    Dim ResEstoque As Object
    Dim ResCentroResultado As ADODB.Recordset
    Dim Sql As String
    
    subManutencaoJanelasAtivas "I", "FrmNotasFiscais"
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If

    XFT_STATUSANT = ""
    
'    If XInserir = "I" Then
'        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE (tido_tx_sigla IS NULL or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') and tido_bl_ativo = 1  ORDER BY tido_tx_descricao", Estatico
'    Else
'        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos ORDER BY tido_tx_descricao", Estatico
'    End If
        
    subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE (tido_tx_sigla IS NULL or tido_tx_sigla='NF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') and tido_bl_ativo = 1  ORDER BY tido_tx_descricao", Estatico
        
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico
    
    'Cássio 08/08/2012
    subConectarControleDadosNV DatCriterios, "SELECT * FROM CriteriosDistribuicao WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY crdi_tx_descricao", Estatico
    
    'Comum às orelhas de dados gerais e duplicatas
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    Else
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    End If
    
    'Comum às orelhas de dados gerais e distribuição
    If XInserir = "A" Then
        subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcomboAtivoInativo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    Else
        subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_tx_status = 'A' ORDER BY cere_cd_estruturado", Estatico
    End If
        
    'Cássio Medeiros 25/09/2013
    If XInserir = "A" Then
        Sql = "SELECT CentroResultados.cere_cd_Pcr, CentroResultados.cere_tx_status FROM Distribuicao INNER JOIN CentroResultados ON CentroResultados.cere_cd_Pcr = Distribuicao.cere_cd_Pcr" & _
              " INNER JOIN NotasFiscais ON NotasFiscais.nofi_cd_notafiscal = Distribuicao.nofi_cd_notafiscal" & _
              " WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & ""
              
        SubQOpenRecordset ResCentroResultado, Sql, "Estatico"
        
        If Not (ResCentroResultado.BOF And ResCentroResultado.EOF) And ResCentroResultado.RecordCount = 1 Then
            'Verifica se o CR vinculado à nota está inativo
            If ResCentroResultado!cere_tx_status = "I" And CboStatus.ListIndex = 2 Then
                subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcomboAtivoInativo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr = " & ResCentroResultado!cere_cd_Pcr & " ORDER BY cere_cd_estruturado", Estatico
            ElseIf ResCentroResultado!cere_tx_status = "I" And CboStatus.ListIndex <> 2 Then
                subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcomboAtivoInativo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
            End If
        End If
    End If
    
    cboPlanoContasCredito.Visible = False
    CmdLimparIdentCcredito.Visible = False
    MskcontaCredito.Visible = False
    lblContaCredito.Visible = False
    
    '
    ' Carga inicial de DatConta (combo Contas contábeis), conforme as opções de Inserção ou Alteração e Entrada ou Saida
    ' Alterado em 21/03/2022 - Paulo Garcia
    ' -------------------------------------------------------------------------------------------------------------------
    XFT_CR = ""
    XFT_CT = ""
    
    If XInserir = "A" Then
                                
'       ' Se informação de vinculação a nível de empresa for verdadeira
'        If pVinculaCrCc = True Then
        
            If EntSaida = "E" Then
                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'R' OR plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & ResCentroResultado!cere_cd_Pcr & ") ORDER BY plco_tx_desccontabil", Estatico
            ElseIf EntSaida = "S" Then
                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'D' OR plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & ResCentroResultado!cere_cd_Pcr & ") ORDER BY plco_tx_desccontabil", Estatico
            End If
        
'        Else
'
'            If EntSaida = "E" Then
'                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
'            ElseIf EntSaida = "S" Then
'                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
'            End If
'
'        End If
        
    ElseIf XInserir = "I" Then
    
        If EntSaida = "E" Then
            subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        ElseIf EntSaida = "S" Then
            subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND plco_tx_tipo = 'A' AND plco_tx_status = 'A' AND (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        End If
        
    End If

    
'    If EntSaida = "E" Then
'        FrmNotasFiscais.Caption = "Cadastro de Contas a Receber"
'        LblNomeFantasia.Caption = "Cliente:"
'        LblValorNota2.Caption = "Valor a Receber:"
'        subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
'        subConectarControleDadosNV DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'C' and empr_cd_empresa = " & PCodEmpresa & " ORDER BY focl_tx_razaosocial ", Estatico
''        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'        subConectarControleDadosNV DatContaContabilCredito, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'
'        cboPlanoContasCredito.Visible = False
'        MskcontaCredito.Visible = False
'        lblContaCredito.Visible = False
'
'        LblTipoPagto.Caption = "Tipo Recebimento:"
'        LblDocPagto.Caption = "Doc. Recebimento:"
'        LblBcForn.Visible = False
'        CboBancoForn1.Visible = False
'        LblPagto.Caption = "Recebimento:"
'        PanPagto1.Top = 570
'        CboTipoPagto1.RemoveItem (1)
'        CboTipoPagto1.RemoveItem (1)
'        CboTipoPagto1.RemoveItem (1)
'        CboTipoPagto1.RemoveItem (1)
'
'        CboTipoPagto2.RemoveItem (1)
'        CboTipoPagto2.RemoveItem (1)
'        CboTipoPagto2.RemoveItem (1)
'        CboTipoPagto2.RemoveItem (1)
'
'        TDBGrid2.Columns("Tipo Pag.").Caption = "Tipo Rec."
'        TDBGrid2.Columns("Doc. Pag.").Caption = "Doc. Rec."
'        TDBGrid2.Columns("Data Pag.").Caption = "Data Rec."
'        TDBGrid2.Columns("Conta Corrente (Fornecedor)").Visible = False
'        TDBGrid1.Columns("Vencimento").Visible = False
'        LblVlDesc.Left = LblVencimento3.Left
'        TxtValorDesconto.Left = DtpVencDesconto.Left
'        TabAuxiliar.TabCaption(1) = "Recebimento"
'        LblVencimento3.Visible = False
'        DtpVencDesconto.Visible = False
'
'    ElseIf EntSaida = "S" Then
    
     LblNomeFantasia.Caption = "Fornecedor:"
     LblValorNota2.Caption = "Valor a Pagar"
'     subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
'     subConectarControleDadosNV DatFornecedor, "SELECT * FROM ConsGENFornCli WHERE focl_tx_classe = 'F' and empr_cd_empresa = " & PCodEmpresa & " ORDER BY focl_tx_razaosocial ", Estatico
'     subConectarControleDadosNV DatFornecedor, "SELECT * FROM ConsGENFornCli A JOIN Corretores B ON A.focl_cd_FornCli = B.focl_cd_FornCli WHERE focl_tx_Classe = 'F' AND empr_cd_Empresa = " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
     subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
     subConectarControleDadosNV DatContaContabilCredito, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
     subConectarControleDados DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    
'    If (EntSaida = "S") And (PCliente = "QUALIDADOS") Then
'         cboPlanoContasCredito.Visible = True
'         MskcontaCredito.Visible = True
'         lblContaCredito.Visible = True
'     End If
     
     CboTipoPagto1.RemoveItem (5)
     CboTipoPagto2.RemoveItem (5)
        
'    End If
     
    EntrouDesc = False
    EntrouDupl = False
    EntrouDist = False
    ParcelasMultiplas = False
    XFB_ESTOQUE = False
    XFB_QUALIFAT = False
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
        
        PreencheCampos (XFormulario)
               
'        If XFormulario = "TelaDuplicatas" Then 'Foi chamado da tela de Duplicata
'            If Status <> "0" And Tipo = "0" Then 'Não é previsto e o tipo Doc. não gera dupl. automaticamente
'                FrmNotasFiscais.TabNotasFiscais.Tab = 1
'            End If
'        End If
                
        If XFB_ESTOQUE = True Then 'A nota é do Estoque
        
            If (PEmpresa <> "QUALIDADOS C. E SISTEMAS LTDA SALVADOR" And PEmpresa <> "QUALINFO S. DE INFORMATICA LTDA - EPP SALVADOR") Then
        
                SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
                DesabilitaRealizado "Estoque", False
                'Não permite alterar notas do estoque, se não tiver selecionado o campo que
                'permite alterar em padrões de empresa - Lucas Santiago 23.01.09
                If (ResEstoque!empr_tx_AltNotasEstoque) = "S" Then
                    ChkDistribuicao.Enabled = True
                    'Habilita botão gravar - Cássio Medeiros 04/08/2014
                    CmdGravar.Enabled = True
                    
                    If ChkDistribuicao.Value = 0 Then
                        MskConta.Enabled = True
                        MskPcr.Enabled = True
                        CboPlanoContas.Enabled = True
                        CboPcr.Enabled = True
                        CboIdentificador2.Enabled = True
                    Else
                        MskPcr2.Enabled = True
                        CboPcr2.Enabled = True
                        MskConta2.Enabled = True
                        CboPlanoContas2.Enabled = True
                        CboIdentificador.Enabled = True
                        TxtPercDistribuicao.Enabled = True
                        TxtVldistribuicao.Enabled = True
                        
                        CmdInsereDistribuicao.Enabled = True
                        CmdRemoverDistribuicao.Enabled = True
                    End If
                End If
            End If
        'Se a nota for do QualiFat ou QualiCapi, desabilitar todas as opções, só permitir consultar
        ElseIf XFB_QUALIFAT = True Or XFB_QUALICAPI = True Then
            DesabilitaRealizado "Estoque", False
            TxtPrazo.Enabled = False
            CmdAlteraDistribuicao.Enabled = False
            CmdAlteraDesconto.Enabled = False
            CmdInsereDesconto.Enabled = False
            CmdGravar.Enabled = False
        Else
            If Status <> "2" Then
                HabilitaControles
            Else
                If Not FunVerifica_Permissao("TelaNotasFiscais", "3") Then
                    DesabilitaRealizado "Nota", False
                Else
                    HabilitaControles
                End If
            End If
        End If
        
    End If
    
    If DatTipoDoc.Recordset.Fields("tido_tx_sigla") = "PC" Then
           ChkDistribuicao.Enabled = False
    End If
    
    
    '******************REGISTRA LOG DA OPERAÇÃO - PARTE 1**************************
    
    'Daniel Coelho - 18/06/2012 antes XFV_VETORLOG(50)
    'Inicializa variáveis de log - Larissa 25/06/2008
    ReDim XFV_VETORLOG(100) As Variant 'linha de código alterada em 29/01/09 - Patrícia. Anterior:ReDim XFV_VETORLOG(31) As Variant
    XFT_CONTADORLOG = 0
    
    'NOTA FISCAL
    If XInserir = "I" Then ReDim XFM_MATRIZNOTAFISCAL(10, 2) As Variant
    If XInserir = "A" Then ReDim XFM_MATRIZNOTAFISCAL(10, 3) As Variant
        
    'Nome dos campos da nota fiscal
    XFM_MATRIZNOTAFISCAL(0, 0) = "Nº Doc"
    XFM_MATRIZNOTAFISCAL(1, 0) = "Num Fatura"
    XFM_MATRIZNOTAFISCAL(2, 0) = "Valor"
    XFM_MATRIZNOTAFISCAL(3, 0) = "Serie"
    XFM_MATRIZNOTAFISCAL(4, 0) = "Status"
    If EntSaida = "E" Then XFM_MATRIZNOTAFISCAL(5, 0) = "Cliente"
    If EntSaida = "S" Then XFM_MATRIZNOTAFISCAL(5, 0) = "Fornecedor"
    XFM_MATRIZNOTAFISCAL(6, 0) = "Tipo Doc."
    XFM_MATRIZNOTAFISCAL(7, 0) = "Data Entrada" 'Data de entrada da nota no sistema
    XFM_MATRIZNOTAFISCAL(8, 0) = "Data Emissão"  'Data de emissão da nota
    XFM_MATRIZNOTAFISCAL(9, 0) = "Historico"
    
    'Dados originais da nota
    XFM_MATRIZNOTAFISCAL(0, 1) = MskNumdocumento.Text
    XFM_MATRIZNOTAFISCAL(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
    XFM_MATRIZNOTAFISCAL(3, 1) = TxtSerie.Text
    XFM_MATRIZNOTAFISCAL(4, 1) = CboStatus.Text
    XFM_MATRIZNOTAFISCAL(5, 1) = CboFornecedor.Text
    XFM_MATRIZNOTAFISCAL(6, 1) = CboTipoDoc.Text
    XFM_MATRIZNOTAFISCAL(7, 1) = Format$(DtpEntrada.Value, "DD/MM/YY")
    XFM_MATRIZNOTAFISCAL(8, 1) = Format$(DtpEmissao.Value, "DD/MM/YY")
    XFM_MATRIZNOTAFISCAL(9, 1) = FunNulo(CboHistorico.Text)
    
    'PAGAMENTO
    If XInserir = "I" Then ReDim XFM_MATRIZPAGAMENTO(10, 2) As Variant
    If XInserir = "A" Then ReDim XFM_MATRIZPAGAMENTO(10, 3) As Variant
    
    XFM_MATRIZPAGAMENTO(0, 0) = "Num Fatura" 'Número da fatura
    XFM_MATRIZPAGAMENTO(1, 0) = "Cod"
    XFM_MATRIZPAGAMENTO(2, 0) = "Valor"
    XFM_MATRIZPAGAMENTO(3, 0) = "Status"
    XFM_MATRIZPAGAMENTO(4, 0) = "Data Venc"
    XFM_MATRIZPAGAMENTO(5, 0) = "Tipo Pagto"
    XFM_MATRIZPAGAMENTO(6, 0) = "Conta Corrente"
    If (CboBancoForn1.Enabled = True) Then XFM_MATRIZPAGAMENTO(7, 0) = "Banco Fornecedor"
    XFM_MATRIZPAGAMENTO(8, 0) = "Doc Pagto"
    XFM_MATRIZPAGAMENTO(9, 0) = "Data Pagto"
    
    'Dados originais do pagamento
'    XFM_MATRIZPAGAMENTO(0, 1) = MskNumdocumento.Text
    XFM_MATRIZPAGAMENTO(1, 1) = "U"
    XFM_MATRIZPAGAMENTO(2, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
    XFM_MATRIZPAGAMENTO(3, 1) = CboStatus.Text
    XFM_MATRIZPAGAMENTO(4, 1) = Format$(DtpVencParcelaUnica.Value, "DD/MM/YYYY")
    XFM_MATRIZPAGAMENTO(5, 1) = CboTipoPagto1.Text
    XFM_MATRIZPAGAMENTO(6, 1) = CboCCorrente.Text
    XFM_MATRIZPAGAMENTO(7, 1) = FunNulo(CboBancoForn1.Text)
    XFM_MATRIZPAGAMENTO(8, 1) = TxtDocPagto1.Text
    XFM_MATRIZPAGAMENTO(9, 1) = Format$(DtpPagtoParcelaUnica.Value, "DD/MM/YYYY")
    
    'APROPRIAÇÃO
    If XInserir = "I" Then ReDim XFM_MATRIZAPROPRIACAO(6, 2) As Variant
    If XInserir = "A" Then ReDim XFM_MATRIZAPROPRIACAO(6, 3) As Variant
     
    XFM_MATRIZAPROPRIACAO(0, 0) = "Nota Fiscal"
    XFM_MATRIZAPROPRIACAO(1, 0) = "Centro Custo"
    XFM_MATRIZAPROPRIACAO(2, 0) = "Conta Contábil"
    XFM_MATRIZAPROPRIACAO(3, 0) = "Percentual"
    XFM_MATRIZAPROPRIACAO(4, 0) = "Valor"
    XFM_MATRIZAPROPRIACAO(5, 0) = "Identificador"
    
    'Dados originais da apropriação
    XFM_MATRIZAPROPRIACAO(0, 1) = MskNumdocumento.Text
    XFM_MATRIZAPROPRIACAO(1, 1) = CboPcr.Text
    XFM_MATRIZAPROPRIACAO(2, 1) = CboPlanoContas.Text
    XFM_MATRIZAPROPRIACAO(3, 1) = "100%"
    XFM_MATRIZAPROPRIACAO(4, 1) = Format$(FunNuloVal(TxtValor.Text), "##,##0.00")
    XFM_MATRIZAPROPRIACAO(5, 1) = CboIdentificador2.Text
    
    '*****************************FIM LOG - PARTE 1********************************
    
    On Error Resume Next
    If XGB_CONTABILIZA Then
        Grava
        GoTo Fim
        DoEvents
    End If

    If TDBGrid1.ApproxCount - 1 > 0 Then
        CmdExcluirTodos.Enabled = True
    Else: CmdExcluirTodos.Enabled = False
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
    
Fim:
'Unload Me
On Error Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmNotasFiscais"
End Sub

'Private Sub Label4_Click()
'
'End Sub

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
            MsgBox "Não existe item do Plano de Contas com este código reduzido, ou não está vinculado ao Centro de Custo.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta.SetFocus
        End If
    
    End If
End Sub

Private Sub MskConta2_GotFocus()
    Call subSelecionaMSK(MskConta2)
End Sub


Private Sub MskConta2_LostFocus()
    If MskConta2.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskConta2.Text = Format(MskConta2.Text, XGT_MSKCODREDUZIDO)
        End If
    
        CboPlanoContas2.BoundText = MskConta2.Text
        If CboPlanoContas2.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido, ou não está vinculado ao Centro de Custo.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta2.SetFocus
        End If
    End If
End Sub

Private Sub dtpEmissao_LostFocus()
    If Not ConsistenciaDatas Then Exit Sub
    Verifica_DataEmissao
    
    If pVinculaCrCc = True Then
        If ChkDistribuicao.Value = 1 And (DtpEmissao) >= CDate(Data_VinculoCRCC) Then
            subLimpaVetor Vetor3, TDBGrid3
            LblValorTotalDistrib.Caption = ""
            ValorTotalDistribuicao = 0
            XDistribuido = "N"
        End If
    End If
    
End Sub

Private Sub dtpEntrada_LostFocus()
    VerificaImpostosAReter
End Sub

Private Sub MskcontaCredito_GotFocus()
    Call subSelecionaMSK(MskcontaCredito)
End Sub

Private Sub MskcontaCredito_LostFocus()
    If MskcontaCredito.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskcontaCredito.Text = Format(MskcontaCredito.Text, XGT_MSKCODREDUZIDO)
        End If
        cboPlanoContasCredito.BoundText = MskcontaCredito.Text
        If cboPlanoContasCredito.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskcontaCredito.SetFocus
        End If
    
    End If
End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_LostFocus()
    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
            Exit Sub
        End If
        'subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapImoveis WHERE empd_cd_empreendimento='" & MskEmpreendimento.Text & "' and imov_tx_status='A' ORDER BY imov_cd_imovel", Estatico
    Else
      CboEmpreendimento.BoundText = ""
    End If
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
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

Private Sub MskNumdocumento_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8 And KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 65 And KeyAscii <> 97) Then
        KeyAscii = 0
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

Private Sub MskPcr2_GotFocus()
    Call subSelecionaMSK(MskPcr2)
End Sub

Private Sub MskPcr2_LostFocus()
    MskPcr2.PromptInclude = False
    If MskPcr2.Text <> "" Then
        If Len(MskPcr2.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            MskPcr2.Text = funFormataCCusto(MskPcr2.Text)
        End If
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

Private Sub CmdLimparIdentif2_Click()
    CboIdentificador.BoundText = 0
    CboIdentificador.Text = ""
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

Private Sub TabAuxiliar_GotFocus()
    If TabAuxiliar.Tab = 0 Then
        If MskPcr.Enabled Then MskPcr.SetFocus
    ElseIf TabAuxiliar.Tab = 1 Then
        If DtpVencParcelaUnica.Enabled Then DtpVencParcelaUnica.SetFocus
    End If
End Sub

Private Sub TabNotasFiscais_Click(PreviousTab As Integer)
    Dim Sql As String
    Dim ResForn As Object
    Dim XPrevRecolhimento As String
    Dim XVlDesconto As Double
    Dim XLF_INSS As Currency    'Valor de INSS para reduzir a base de cálculo para o INSS
    Dim XLF_BASECALCULO As Currency    'Base de Cálculo p/ IRPF = Valor NF - Valor INSS
    Dim XData As Date
    Dim Ano As Integer
    Dim Mes As Integer
    Dim XMesPosterior As String
    Dim ResEstoque As Object
    
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
      
    'ORELHA DE DUPLICATAS
    If TabNotasFiscais.Tab = 1 And Not EntrouDupl Then
               
        If EntSaida = "E" Then 'Prepara o painel para Contas a Receber
            LblTipoPagamento.Caption = "Tipo Recebimento:"
            LblDocPagRec.Caption = "Doc. Recebimento:"
            LblBcFornecedor.Visible = False
            CboBancoForn2.Visible = False
            LblPagamento.Caption = "Recebimento:"
            PanPagto2.Top = 375
        End If
        
        Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
        subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "ValorDuplicata", "dupl_vl_GlosaDevolucao", "Valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BAEmpresa", "bafo_cd_codigo", "BAFornecedor", "dupl_cd_duplicata", "dupl_tx_estorno", "dupl_tx_Consolidado", "dupl_dt_concilia", "dupl_dt_ProcPagamento", "fopa_cd_codigo", "fopa_tx_descricao"), Vetor2, TDBGrid2
    
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
                
                'O qualifin não está autorizado a alterar uma nota originada do qualimat
                If XFB_ESTOQUE = False Then 'A nota não é do Estoque
                    SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
                    'Não permite alterar notas do estoque, se não tiver selecionado o campo que
                    'permite alterar em padrões de empresa - Lucas Santiago 23.01.09
                    If (ResEstoque!empr_tx_AltNotasEstoque) = "S" Then
                        CmdAlteraduplicata_Click
                    End If
                End If
                TDBGrid2.Bookmark = PosCursor
            End If
            
        End If
        LblValorTotal = funSomaColuna(Vetor2, 2)
        ValorTotalDuplicata = CDbl(LblValorTotal.Caption)
        If EntrouDesc Then
            If Vetor1(0, 0) <> Empty Then
                LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
            End If
        End If
    
    'ORELHA DE DESCONTOS
    ElseIf TabNotasFiscais.Tab = 2 Then
        If Not EntrouDesc Then
            If EntSaida = "S" Then  'É um fornecedor
                If XGT_ORDEMDESCONTO = "C" Then
                    'Ordernar por código reduzido
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE plco_cd_CodReduzido + ' - ' + desc_tx_descricao END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'D' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) AND NOT plco_cd_CodReduzido IS NULL ORDER BY plco_cd_CodReduzido"
                Else
                    'Ordenar por Descrição
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE desc_tx_descricao + ' (' + plco_cd_CodReduzido + ')' END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'D' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao"
                End If
                subConectarControleDadosNV DatDescontosNF, Sql, "Estatico"
            Else  'É um Cliente
                If XGT_ORDEMDESCONTO = "C" Then
                    'Ordernar por código reduzido
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE plco_cd_CodReduzido + ' - ' + desc_tx_descricao END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'I' and desc_tx_fonte = 'S' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) AND NOT plco_cd_CodReduzido IS NULL ORDER BY plco_cd_CodReduzido"
                Else
                    'Ordenar por Descrição
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE desc_tx_descricao + ' (' + plco_cd_CodReduzido + ')' END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'I' and desc_tx_fonte = 'S' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao"
                End If
                subConectarControleDadosNV DatDescontosNF, Sql, "Estatico"
            End If
            
            Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
            subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura", "deen_dt_recolhimento", "deen_nr_DarfDam"), Vetor1, TDBGrid1
            If Vetor1(0, 0) = Empty Then
                DesabilitaBotoes
                If XFB_QUALIFAT = False Then 'A Nota não é do QualiFat
                    CmdInsereDesconto.Enabled = True
                End If
            End If
        
            'Verifica se o fornecedor possui um imposto cadastrado
            If EntSaida = "S" And (XInserir = "I" Or XFT_STATUSANT = "P") Then  'É um fornecedor
            'O IF acima testa se é uma saída e se está inserindo ou se está alterando, mas o status anterior era previsto
            'para esse último caso é interessante que o sistema permita o cálculo automático dos impostos
                'O select abaixo foi ordenado por tipobase para que seja possível pegar o valor do INSS para dedução da base de cálculo do IRRF, qdo for pessoa física
                SubQOpenRecordset ResForn, "SELECT DesImpForn.*,DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase FROM DesImpForn LEFT JOIN DescontosImpostos ON DesImpForn.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DesImpForn.focl_cd_forncli=" & CboFornecedor.BoundText & " ORDER BY desc_tx_tipobase DESC", "Estatico"
                If Not (ResForn.BOF And ResForn.EOF) Then
                    ResForn.MoveFirst
                    Do While Not ResForn.EOF
                        If ResForn!desc_tx_tipobase = "1" Or ResForn!desc_tx_tipobase = "3" Then
                            'INSS ou ISS
                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            If EntrouDupl Then
                                'Se o vetor de duplicatas tiver preenchido, calcular com base no vencimento da 1ª parcela
                                If Vetor2(0, 0) <> Empty Then
                                    Ano = Year(CDate(Vetor2(0, 1)))
                                    Mes = Month(CDate(Vetor2(0, 1)))
                                    
                                    If Mes = 12 Then
                                        Mes = 1
                                        Ano = Ano + 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                'senão, calcular com base na data de emissão
                                Else
                                    If ResForn!desc_tx_tipobase = "1" Then 'ISS
                                        Ano = Year(DtpEmissao)
                                        Mes = Month(DtpEmissao)
                                        
                                        If Mes = 12 Then
                                            Mes = 1
                                            Ano = Ano + 1
                                        Else
                                            Mes = Mes + 1
                                        End If
                                        XData = DateSerial(Ano, Mes, 5)
                                        If XVlDesconto >= 6 Then
                                            subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                        End If
                                    Else 'INSS
                                        Ano = Year(DtpEmissao)
                                        Mes = Month(DtpEmissao)
                                        
                                        If Mes = 12 Then
                                            Mes = 1
                                            Ano = Ano + 1
                                        Else
                                            Mes = Mes + 1
                                        End If
                                        XData = DateSerial(Ano, Mes, 20) '30/04/09 - Patrícia. Anterior:XData = DateSerial(Ano, Mes, 2)
                                        
                                        'Solicitado pela APAE - teto máximo de contribuição é R$ 275,95 para P. Física
                                        If DatFornecedor.Recordset.Fields("focl_tx_Tipo") = "F" And XVlDesconto > 275.95 Then
                                            XVlDesconto = 275.95
                                        End If
                                        XLF_INSS = XVlDesconto
                                        
                                        subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                    End If
                                End If
                            'se não entrou na orelha de duplicatas, calcular com base na data de emissão
                            Else
                                If ResForn!desc_tx_tipobase = "1" Then 'ISS
                                    Ano = Year(DtpEmissao)
                                    Mes = Month(DtpEmissao)
                                    
                                    If Mes = 12 Then
                                        Mes = 1
                                        Ano = Ano + 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    XData = DateSerial(Ano, Mes, 2)
                                    If XVlDesconto >= 6 Then
                                        subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                    End If
                                Else 'INSS
                                    Ano = Year(DtpEmissao)
                                    Mes = Month(DtpEmissao)
                                    
                                    If Mes = 12 Then
                                        Mes = 1
                                        Ano = Ano + 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    XData = DateSerial(Ano, Mes, 20) '30/04/09 - Patrícia. Anterior:XData = DateSerial(Ano, Mes, 2)
                                    
                                    'Solicitado pela APAE - teto máximo de contribuição é R$ 275,95 para P. Física
                                    If DatFornecedor.Recordset.Fields("focl_tx_Tipo") = "F" And XVlDesconto > 275.95 Then
                                        XVlDesconto = 275.95
                                    End If
                                    XLF_INSS = XVlDesconto
                                    
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                End If
                            End If
'Solicitado pela APAE para que o cálculo do IRRF fosse da mesma forma que PIS/COFINS e CSLL
'                        ElseIf ResForn!desc_tx_tipobase = "2" Then
'                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
'                            If EntrouDupl Then
'                                If Vetor2(0, 0) <> Empty Then
'                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, Format(DtpVencimento.Value, "dd/mm/yyyy"), XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
'                                Else
'                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
'                                End If
'                            Else
'                                subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "", XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
'                            End If
                        ElseIf ResForn!desc_tx_tipobase = "2" Or ResForn!desc_tx_tipobase = "4" Or ResForn!desc_tx_tipobase = "5" Or ResForn!desc_tx_tipobase = "6" Then
                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            
                            'se entrou na orelha de duplicatas
                            If EntrouDupl Then
                                'se o vetor estiver preenchido, calcular com a data de
                                'vencimento da 1ª parcela
                                If Vetor2(0, 0) <> Empty Then
                                    XData = CDate(Vetor2(0, 1))
                                Else
                                    XData = DtpEmissao.Value
                                End If
                            'se não entrou na orelha de duplicatas, calcular c/ dt. emissão
                            Else
                                XData = DtpEmissao.Value
                            End If
                            
                            'Nova regra do IRRF 2006
                            Mes = Month(XData)
                            Ano = Year(XData)
                            If Mes = 12 Then
                                Mes = 1
                                Ano = Ano + 1
                            Else
                                Mes = Mes + 1
                            End If

                           XMesPosterior = "20/" & Mes & "/" & Ano 'Alteração em 30/04/09 - Patrícia. Anterior:XMesPosterior = "10/" & Mes & "/" & Ano
                        
                           If Weekday(XMesPosterior) = "7" Then
                                XData = "19/" & Mes & "/" & Ano
                            ElseIf Weekday(XMesPosterior) = "1" Then
                                XData = "18/" & Mes & "/" & Ano
                            Else
                                XData = XMesPosterior
                            End If
                            
                            'Tudo o que for retido na semana será recolhido na próxima quarta-feira
'                            Select Case Weekday(XData)
'                                Case 1 'Domingo
'                                    XData = DateAdd("d", 10, XData)
'                                Case 2 'Segunda
'                                    XData = DateAdd("d", 9, XData)
'                                Case 3 'Terça
'                                    XData = DateAdd("d", 8, XData)
'                                Case 4 'Quarta
'                                    XData = DateAdd("d", 7, XData)
'                                Case 5 'Quinta
'                                    XData = DateAdd("d", 6, XData)
'                                Case 6 'Sexta
'                                    XData = DateAdd("d", 5, XData)
'                                Case 7 'Sábado
'                                    XData = DateAdd("d", 11, XData)
'                            End Select
                            
                            'Solicitado pela APAE para IRPF P. Física
                            'Base de Cálculo = Vl Bruto - Vl INSS
                            'Para proventos de R$1.164,00 até 2.326,00 alicota de 15% - parcela a deduzir de R$ 174,60
                            'Para proventos acima de R$ 2.326,00 alicota de 27,5% - parcela a deduzir de 465,35.
                            Dim XLV_VALORDESCTEMP As Currency
                            XLV_VALORDESCTEMP = 0
                            If ResForn!desc_tx_tipobase = "2" And DatFornecedor.Recordset.Fields("focl_tx_Tipo") = "F" Then
                                XLF_BASECALCULO = CDbl(TxtValor.Text) - XLF_INSS
                                XVlDesconto = Format$(((XLF_BASECALCULO * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                                
                                'onde estava sendo usado CDbl(txtValor.Text) foi mudado para XLF_BASECALCULO
                                If XLF_BASECALCULO = 1257.12 Then
                                    XVlDesconto = 0
                                ElseIf XLF_BASECALCULO > 1257.12 And XLF_BASECALCULO <= 2512.08 Then
                                    XLV_VALORDESCTEMP = (XLF_BASECALCULO * 0.15) - 188.57
                                ElseIf XLF_BASECALCULO > 2512.08 Then
                                    XLV_VALORDESCTEMP = (XLF_BASECALCULO * 0.275) - 465.35
                                End If

                                If XVlDesconto > XLV_VALORDESCTEMP And XLV_VALORDESCTEMP <> 0 Then
                                    XVlDesconto = XLV_VALORDESCTEMP
                                End If
                            End If
                            
                            If XVlDesconto <> 0 Then
                                subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                            End If
                        Else
                            'OUTROS IMPOSTOS
                            XVlDesconto = Format$(((CDbl(TxtValor.Text) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            'Ano = Year(CDate(MskVencimento.Text))
                            'Mes = Month(CDate(MskVencimento.Text))
                            If EntrouDupl Then
                                'Se o vetor de duplicatas tiver preenchido, calcular com base no
                                'vencimento da 1ª parcela
                                If Vetor2(0, 0) <> Empty Then
                                    Ano = Year(CDate(Vetor2(0, 1)))
                                    Mes = Month(CDate(Vetor2(0, 1)))

                                    If Mes = 12 Then
                                        Mes = 1
                                        Ano = Ano + 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    
                                    If IsNull(ResForn!desc_tx_DiaRecolhimento) Then
                                        XData = DateSerial(Ano, Mes, Day(CDate(Vetor2(0, 1))))
                                    Else
                                        XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                                    End If
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                'senão, calcular com base na data de emissão
                                Else
                                    Ano = Year(DtpEmissao)
                                    Mes = Month(DtpEmissao)

                                    If Mes = 12 Then
                                        Mes = 1
                                        Ano = Ano + 1
                                    Else
                                        Mes = Mes + 1
                                    End If
                                    
                                    If IsNull(ResForn!desc_tx_DiaRecolhimento) Then
                                        XData = DateSerial(Ano, Mes, Day(DtpEmissao))
                                    Else
                                        XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                                    End If
                                    subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                                End If
                            'se não entrou na orelha de duplicatas, calcular com base na data de emissão
                            Else
                                Ano = Year(DtpEmissao)
                                Mes = Month(DtpEmissao)

                                If Mes = 12 Then
                                    Mes = 1
                                    Ano = Ano + 1
                                Else
                                    Mes = Mes + 1
                                End If
                                
                                If IsNull(ResForn!desc_tx_DiaRecolhimento) Then
                                    XData = DateSerial(Ano, Mes, Day(DtpEmissao))
                                Else
                                    XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                                End If
                                subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                            End If
                        End If
                        ResForn.MoveNext
                        HabilitaBotoes
                    Loop
                    ResForn.Close
                End If
                LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
            ElseIf EntSaida = "E" And (XInserir = "I" Or XFT_STATUSANT = "P") Then 'É um Cliente
            'O IF acima testa se é uma entrada e se está inserindo ou se está alterando, mas o status anterior era previsto
            'para esse último caso é interessante que o sistema permita o cálculo automático dos impostos
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
                       subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "  /  /  ", XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor1, TDBGrid1
                       ResForn.MoveNext
                       HabilitaBotoes
                    Loop
                    ResForn.Close
                End If
                LblValorLiquido1.Caption = Format$((CDbl(TxtValor.Text) - funSomaColuna(Vetor1, 2) - CDbl(TxtDevolucao.Text)), "##,##0.00")
            End If
            EntrouDesc = True
        End If
        
    'ORELHA DE DISTRIBUICAO
    ElseIf TabNotasFiscais.Tab = 3 And Not EntrouDist Then
        'Sql = "SELECT Distribuicao.*,Notasfiscais.nofi_vl_valor,(Distribuicao.dist_vl_valor / NotasFiscais.nofi_vl_valor) AS Percentual,CentroResultados.cere_cd_estruturado, PlanoContas.plco_cd_codreduzido " & _
        "FROM NotasFiscais,Distribuicao,CentroResultados, PlanoContas WHERE NotasFiscais.nofi_cd_notafiscal = " & Chave & " and Distribuicao.nofi_cd_notafiscal = " & Chave & " and Distribuicao.cere_cd_pcr =CentroResultados.cere_cd_pcr and Distribuicao.plco_cd_conta = PlanoContas.plco_cd_conta"
        'subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "dist_vl_valor", "percentual"), Vetor3, TDBGrid3
        'Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " and nofi_cd_notafiscal = " & Chave & ""
        'subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "percentual"), Vetor3, TDBGrid3
        Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & " ORDER BY cere_cd_Estruturado"
        subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "dist_vl_GlosaDevolucao", "dist_vl_valor", "percentual", "iden_cd_Projeto", "iden_tx_Descricao"), Vetor3, TDBGrid3
        
        '       Incluído por Felipe Andrade - 28/09/2021 - Carregando novamente o DatPcr e DatConta, para servir de fonte de dados do CboPcr2 eCboPlanoContas2
        subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_tx_status = 'A' ORDER BY cere_cd_estruturado", Estatico
            If EntSaida = "E" Then
                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
            ElseIf EntSaida = "S" Then
                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
            End If
        
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

'Private Sub Text1_Change()
'
'End Sub

Private Sub TxtCCorrente_GotFocus()
    Call subSelecionaTXT(TxtCCorrente)
End Sub

Private Sub TxtCCorrente_LostFocus()
    If TxtCCorrente.Text <> "" Then
        CboCCorrente.BoundText = TxtCCorrente.Text
        If CboCCorrente.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente.SetFocus
        End If
    End If
End Sub

Private Sub TxtCCorrente2_GotFocus()
    Call subSelecionaTXT(TxtCCorrente2)
End Sub

Private Sub TxtCCorrente2_LostFocus()
    If TxtCCorrente2.Text <> "" Then
        CboCCorrente2.BoundText = TxtCCorrente2.Text
        If CboCCorrente2.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente2.SetFocus
        End If
    End If
End Sub

Private Sub TxtCodTipoDoc_GotFocus()
    Call subSelecionaTXT(TxtCodTipoDoc)
End Sub

Private Sub TxtCodTipoDoc_LostFocus()
    If TxtCodTipoDoc.Text <> "" Then
        CboTipoDoc.BoundText = TxtCodTipoDoc.Text
        If CboTipoDoc.Text = "" Then
            MsgBox "Não existe Tipo de Documento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCodTipoDoc.SetFocus
        End If
    End If
End Sub


Private Sub TxtDocPagto1_LostFocus()
    Dim XLO_CHECAR As ADODB.Recordset
    If CboTipoPagto1.ListIndex = 0 Then
        TxtDocPagto1.Text = Format(TxtDocPagto1, "0000000")
    End If
    
        
    'Checar se o nº do doc. pagto. já não está associado a um pagto nesta data e para esta conta
    If TxtDocPagto1.Text <> "" And CboStatus.Text = "Realizado" And CboCCorrente.Text <> "" Then
        If EntSaida = "S" Then ' Despesa
            SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex) & "' and dupl_tx_docpagrec ='" & TxtDocPagto1.Text & "' and dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "", "Estatico"
            If Not XLO_CHECAR.EOF Then
                MsgBox "Este nº de Doc. de Pagto. já foi emitido!", vbCritical, "ATENÇÃO!"
                TxtDocPagto1.SetFocus
                Exit Sub
            End If
            XLO_CHECAR.Close
        Else
            SubQOpenRecordset XLO_CHECAR, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='C' and coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '" & CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex) & "' and dupl_tx_docpagrec ='" & TxtDocPagto1.Text & "'  and dupl_dt_pagrec = " & FunNuloData(DtpPagtoParcelaUnica.Value, NomeSgbd) & "", "Estatico"
            If Not XLO_CHECAR.EOF Then
                MsgBox "Este nº doc. de crédito já foi registrado!", vbCritical, "ATENÇÃO!"
                TxtDocPagto1.SetFocus
                Exit Sub
            End If
            XLO_CHECAR.Close
        End If
    End If
End Sub

Private Sub TxtDocPagto2_LostFocus()
    If CboTipoPagto2.ListIndex = 0 Then
        TxtDocPagto2.Text = Format(TxtDocPagto2, "0000000")
    End If
End Sub

Private Sub TxtHistorico_GotFocus()
    Call subSelecionaTXT(TxtHistorico)
End Sub

Private Sub TxtHistorico_LostFocus()
    If TxtHistorico.Text <> "" Then
        CboHistorico.BoundText = TxtHistorico.Text
        If CboHistorico.Text = "" Then
            MsgBox "Não existe Histórico com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtHistorico.SetFocus
        End If
    End If
End Sub

Private Sub TxtPercDistribuicao_GotFocus()
    Call subSelecionaTXT(TxtPercDistribuicao)
End Sub

Private Sub TxtPercDistribuicao_KeyPress(KeyAscii As Integer)
    'If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        'KeyAscii = 0
    'End If
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If Len(TxtPercDistribuicao.Text) >= 7 Then
        KeyAscii = 0
        TxtPercDistribuicao.Text = Left$(TxtPercDistribuicao.Text, 7)
    End If

End Sub

Private Sub TxtPercDistribuicao_LostFocus()
    If TxtPercDistribuicao.Text <> "" Then
        TxtVldistribuicao.Text = Format$(((CDbl(TxtValor.Text) * CDbl(TxtPercDistribuicao.Text)) / 100), "##,##0.00")
    End If
End Sub

Private Sub TxtPrazo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValor_LostFocus()
    If IsNumeric(TxtValor.Text) Then
        TxtValor.Text = Format$(TxtValor.Text, "##,##0.00")
        LblValorLiquido1.Caption = Format$(TxtValor.Text - CDbl(TxtDevolucao.Text), "##,##0.00")
        LblValorNota.Caption = Format$(TxtValor.Text, "##,##0.00")
    Else
        TxtValor.Text = ""
    End If
End Sub

Private Sub TxtValorDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtValorDuplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
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
'    If KeyAscii = 46 Then KeyAscii = 44
'    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
'        KeyAscii = 0
'    End If

End Sub

Private Sub TxtVldistribuicao_LostFocus()
    If TxtVldistribuicao.Text <> "" Then
        'TxtPercDistribuicao.Text = Format(CStr(((CDbl(TxtVldistribuicao.Text) / CDbl(TxtValor.Text)) * 100)), "standard")
        TxtPercDistribuicao.Text = ((CDbl(TxtVldistribuicao.Text) / CDbl(TxtValor.Text)) * 100)
    End If

End Sub

Function FunVerificaClassEmprObr_Combo(CentroCusto As String, CContabil As String, Empresa As Integer) As String
'Criação da Função em 09/06/08 - Patrícia
'Função que realiza a verificação da classificação "Empreendimento\Obra" dos Centros de Custos e Contas Contábeis
'Só serão aceitos CRs e Contas Contábeis com a mesma Classificação "Empreendimento\Obra" ou Classificação Nula.

    Dim ResEmpObr As Object
    Dim Sql As String
    Dim Desc_CREmpObr As String
    Dim Desc_CCEmpObr As String
    Dim Mensagen As String

    '******** Início Rotina de Classificação "Empreendimento\Obra"
    Mensagen = ""
    If (DtpEmissao) >= CDate(Data_EmpObr) Then
                    Sql = "SELECT * FROM CentroResultados WHERE cere_cd_Estruturado= '" & CentroCusto & "' and empr_cd_Empresa= " & Str(Empresa)
                    SubQOpenRecordset ResEmpObr, Sql, "Estatico"
                    If Not (ResEmpObr.BOF And ResEmpObr.EOF) Then
                        XCLASSE_CENTROCUSTO = RTrim(LTrim(FunNulo(ResEmpObr!cere_tx_EmpObr)))
                        If XCLASSE_CENTROCUSTO = "E" Then
                                    Desc_CREmpObr = "Empreendimento"
                        ElseIf XCLASSE_CENTROCUSTO = "O" Then
                                    Desc_CREmpObr = "Obra"
                       End If
                    End If
                    ResEmpObr.Close
                    Sql = "SELECT * FROM PlanoContas WHERE plco_cd_CodReduzido= '" & CContabil & "'" & " and empr_cd_Empresa= " & Str(Empresa)
                    SubQOpenRecordset ResEmpObr, Sql, "Estatico"
                    If Not (ResEmpObr.BOF And ResEmpObr.EOF) Then
                            If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "E" Then
                                Desc_CCEmpObr = "Empreendimento"
                            ElseIf Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "O" Then
                                Desc_CCEmpObr = "Obra"
                                Else
                                    Desc_CCEmpObr = "Sem Classificação"
                            End If
            
                            If XCLASSE_CENTROCUSTO = "E" Or XCLASSE_CENTROCUSTO = "O" Then 'Se o CR estiver classificado, só poderá informar conta com a mesma classificação ou classificação = Nullo
                                If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "E" Or Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) = "O" Then
                                    If Trim(FunNulo(ResEmpObr!plco_tx_EmpObr)) <> XCLASSE_CENTROCUSTO Then
                                        Mensagen = "O Centro de Custo é do tipo " & Desc_CREmpObr & " e a Conta Corrente é do tipo " & Desc_CCEmpObr & "!"
                                    End If
                                End If
                            End If
                    End If
                    ResEmpObr.Close
    End If
'******** Fim Rotina de Classificação "Empreendimento\Obra"
  FunVerificaClassEmprObr_Combo = Mensagen
End Function

Function funVerificaProvisao(DataPag As Variant, DatEntr As Variant, DataVenc As Variant, QtdDiasProv As Integer, Status As String) As Boolean
    If (DataPag) <> "" Then
        If QtdDiasProv >= 0 And CboStatus.Text <> "Previsto" Then '13/04/10
            If ((DataPag - DatEntr) >= XGI_QTDDIASPROVISAO) Then
                funVerificaProvisao = True
            Else
                funVerificaProvisao = False
            End If
        Else
            funVerificaProvisao = False
        End If
        
    Else
        If QtdDiasProv >= 0 And CboStatus.Text <> "Previsto" Then '13/04/10
            If ((DataVenc - DatEntr) >= XGI_QTDDIASPROVISAO) Then
                funVerificaProvisao = True
            Else
                funVerificaProvisao = False
            End If
        Else
            funVerificaProvisao = False
        End If
        
    End If

End Function
