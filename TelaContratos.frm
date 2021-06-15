VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TelaManContratos 
   Caption         =   "Tabela de Contratos"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9765
   Begin Threed.SSPanel PanTransfEmp 
      Height          =   4890
      Left            =   1560
      TabIndex        =   48
      Top             =   -240
      Visible         =   0   'False
      Width           =   8010
      _Version        =   65536
      _ExtentX        =   14129
      _ExtentY        =   8625
      _StockProps     =   15
      Caption         =   "Transferência de Imóveis - Entre Empresas"
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
      Begin VB.CheckBox chkIncluirPagos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incluir Pagos"
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
         Left            =   3400
         TabIndex        =   86
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkAtualizaValorContratoOrigem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atualizar valor do contrato de origem"
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
         Left            =   3360
         TabIndex        =   80
         Top             =   4400
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame FraTitulos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Títulos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   61
         Top             =   2200
         Width           =   7815
         Begin VB.CheckBox chkNaoCopiarTitulos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não Copiar Títulos"
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
            Left            =   5700
            TabIndex        =   81
            Top             =   195
            Width           =   1995
         End
         Begin MSMask.MaskEdBox mskTituloInicial 
            Height          =   315
            Left            =   1900
            TabIndex        =   63
            ToolTipText     =   "Digite o título e aperte ENTER"
            Top             =   195
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "##.###.##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskTituloFinal 
            Height          =   315
            Left            =   4000
            TabIndex        =   64
            ToolTipText     =   "Digite o título e aperte ENTER"
            Top             =   195
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "##.###.##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "até:"
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
            Left            =   3500
            TabIndex        =   82
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "A partir de:"
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
            Left            =   855
            TabIndex        =   62
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.CheckBox chkExcluirTitulosOrigem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir títulos na origem"
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
         TabIndex        =   77
         Top             =   4440
         Width           =   3075
      End
      Begin VB.CheckBox ChkExportaFinanceiro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar Títulos para o QualiFin"
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
         TabIndex        =   76
         Top             =   4080
         Width           =   3075
      End
      Begin MSAdodcLib.Adodc DatContaCorrente 
         Height          =   330
         Left            =   6240
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSAdodcLib.Adodc DatConta 
         Height          =   330
         Left            =   6600
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSAdodcLib.Adodc DatPcr 
         Height          =   330
         Left            =   6600
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSAdodcLib.Adodc DatEmpresa 
         Height          =   330
         Left            =   6240
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "DatEmpresa"
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
      Begin MSDataListLib.DataCombo CboEmpresa 
         Bindings        =   "TelaContratos.frx":0000
         Height          =   315
         Left            =   3000
         TabIndex        =   50
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empr_tx_Fantasia"
         BoundColumn     =   "empr_tx_Fantasia"
         Text            =   ""
      End
      Begin VB.Frame FraImovel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imóvel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   1560
         Width           =   7815
         Begin MSMask.MaskEdBox MaskImovel 
            Height          =   315
            Left            =   5760
            TabIndex        =   60
            Top             =   200
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "AAAA.AAAA"
            PromptChar      =   " "
         End
         Begin VB.Label LblImovelOrigem1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imóvel Origem:"
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
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label LblImovelOrigemEmpr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "0001.1001"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   58
            Top             =   240
            Width           =   915
         End
         Begin VB.Label LblImovelDestino1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imóvel de Destino:"
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
            Index           =   0
            Left            =   3960
            TabIndex        =   59
            Top             =   240
            Width           =   1635
         End
      End
      Begin VB.Frame FraEmpreendimento 
         BackColor       =   &H00E0E0E0&
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
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   7815
         Begin MSMask.MaskEdBox MaskEmpree 
            Height          =   300
            Left            =   5760
            TabIndex        =   55
            Top             =   180
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin VB.Label LblEpreendimentoDestino 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Empreend. Destino:"
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
            Index           =   1
            Left            =   3960
            TabIndex        =   54
            Top             =   180
            Width           =   1755
         End
         Begin VB.Label LblCodigoEmpreendimentoOrigem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "0001"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   53
            Top             =   180
            Width           =   435
         End
         Begin VB.Label LblEmpreendimentoOrigem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Empreend. Origem:"
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
            Left            =   1200
            TabIndex        =   52
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.CommandButton CmdConfirmarTransEmp 
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
         Height          =   270
         Left            =   5880
         TabIndex        =   78
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton CmdRetornarTransfEmp 
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
         Height          =   270
         Left            =   6960
         TabIndex        =   79
         Top             =   4080
         Width           =   975
      End
      Begin MSMask.MaskEdBox MskConta 
         Height          =   330
         Left            =   1530
         TabIndex        =   66
         Top             =   2880
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskPcr 
         Height          =   330
         Left            =   1530
         TabIndex        =   70
         Top             =   3270
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboPcr 
         Bindings        =   "TelaContratos.frx":0019
         Height          =   330
         Left            =   2370
         TabIndex        =   71
         Top             =   3270
         Width           =   5115
         _ExtentX        =   9022
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
      Begin MSDataListLib.DataCombo CboPlanoContas 
         Bindings        =   "TelaContratos.frx":002E
         Height          =   330
         Left            =   2370
         TabIndex        =   67
         Top             =   2880
         Width           =   5115
         _ExtentX        =   9022
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
      Begin Threed.SSCommand CmdLimparPcr 
         Height          =   330
         Left            =   7515
         TabIndex        =   72
         Top             =   3270
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         MouseIcon       =   "TelaContratos.frx":0045
         Picture         =   "TelaContratos.frx":0497
      End
      Begin Threed.SSCommand CmdLimparplanoContas 
         Height          =   330
         Left            =   7515
         TabIndex        =   68
         Top             =   2880
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483633
         MouseIcon       =   "TelaContratos.frx":05A9
         Picture         =   "TelaContratos.frx":09FB
      End
      Begin MSDataListLib.DataCombo CboCCorrente 
         Bindings        =   "TelaContratos.frx":0B0D
         Height          =   330
         Left            =   1515
         TabIndex        =   74
         Top             =   3660
         Width           =   5985
         _ExtentX        =   10557
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
      Begin Threed.SSCommand CmdLimparCCorrente 
         Height          =   330
         Left            =   7515
         TabIndex        =   75
         Top             =   3660
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         MouseIcon       =   "TelaContratos.frx":0B2C
         Picture         =   "TelaContratos.frx":0F7E
      End
      Begin VB.Label LblCCorrente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "C/C Destino:"
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
         Left            =   120
         TabIndex        =   73
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label LblConta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Conta Destino:"
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
         Left            =   180
         TabIndex        =   65
         Top             =   2940
         Width           =   1320
      End
      Begin VB.Label LblPcr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "CR Destino:"
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
         Left            =   300
         TabIndex        =   69
         Top             =   3330
         Width           =   1200
      End
      Begin VB.Label LblEmpresa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empresa Destino:"
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
         Left            =   1440
         TabIndex        =   49
         Top             =   480
         Width           =   1575
      End
   End
   Begin Threed.SSPanel PanAlterarCodigo 
      Height          =   2010
      Left            =   3120
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   2970
      _Version        =   65536
      _ExtentX        =   5239
      _ExtentY        =   3545
      _StockProps     =   15
      Caption         =   "Transferência de Imóvel"
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
      Begin VB.CommandButton CmdRetornarCodigo 
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
         Height          =   270
         Left            =   1680
         TabIndex        =   42
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton CmdConfirmarCodigo 
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
         Height          =   270
         Left            =   420
         TabIndex        =   41
         Top             =   1500
         Width           =   975
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   315
         Left            =   1740
         TabIndex        =   43
         Top             =   960
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "AAAA.AAAA"
         PromptChar      =   " "
      End
      Begin VB.Label LlbImove 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imóvel de Destino:"
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
         TabIndex        =   46
         Top             =   1020
         Width           =   1635
      End
      Begin VB.Label LblCodigoContrato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0001.0101"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   540
         Width           =   915
      End
      Begin VB.Label LblImovelOrigem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imóvel de Origem:"
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
         TabIndex        =   44
         Top             =   540
         Width           =   1575
      End
   End
   Begin Threed.SSPanel PanDistratar 
      Height          =   1830
      Left            =   3720
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   4290
      _Version        =   65536
      _ExtentX        =   7567
      _ExtentY        =   3228
      _StockProps     =   15
      Caption         =   " Distrato"
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
      Begin VB.CheckBox ChkLimparImovel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpar o valor do imóvel e as garagens do imóvel?"
         Height          =   195
         Left            =   300
         TabIndex        =   36
         Top             =   840
         Width           =   3915
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
         Height          =   330
         Left            =   1980
         TabIndex        =   18
         Top             =   1320
         Width           =   975
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
         Height          =   330
         Left            =   3120
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpDtDistrato 
         Height          =   315
         Left            =   1860
         TabIndex        =   17
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   103809027
         CurrentDate     =   37180
      End
      Begin VB.Label LblDtDistrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Distrato:"
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
         Left            =   300
         TabIndex        =   20
         Top             =   480
         Width           =   1515
      End
   End
   Begin Threed.SSPanel PanCessao 
      Height          =   2685
      Left            =   840
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   7365
      _Version        =   65536
      _ExtentX        =   12991
      _ExtentY        =   4736
      _StockProps     =   15
      Caption         =   "Cessão"
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
      Begin VB.CheckBox ChkLimparImovel2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpar o valor do imóvel e as garagens do imóvel?"
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   1800
         Width           =   3915
      End
      Begin VB.CheckBox ChkTitulosPagos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copiar os títulos pagos?"
         Height          =   195
         Left            =   660
         TabIndex        =   24
         Top             =   1800
         Width           =   2115
      End
      Begin VB.CommandButton CmdRetornarCessao 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   2220
         Width           =   975
      End
      Begin VB.CommandButton CmdConfirmarCessao 
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
         Left            =   4980
         TabIndex        =   25
         Top             =   2220
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpDtDistratoCessao 
         Height          =   315
         Left            =   6000
         TabIndex        =   23
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   103809027
         CurrentDate     =   37180
      End
      Begin MSDataListLib.DataCombo CboContratoInicial 
         Bindings        =   "TelaContratos.frx":1090
         Height          =   315
         Left            =   1740
         TabIndex        =   30
         Top             =   960
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Parcela"
         BoundColumn     =   "titulo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CboContratoFinal 
         Bindings        =   "TelaContratos.frx":10B1
         Height          =   315
         Left            =   1740
         TabIndex        =   31
         Top             =   1380
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Parcela"
         BoundColumn     =   "titulo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc DatContratoInicial 
         Height          =   330
         Left            =   2700
         Top             =   2220
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
         Caption         =   "DatTitulos"
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
      Begin MSAdodcLib.Adodc DatContratoFinal 
         Height          =   330
         Left            =   2700
         Top             =   2220
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
         Caption         =   "DatTitulos"
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
      Begin VB.Label LblDtContrato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1740
         TabIndex        =   35
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label LblDataContrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Contrato:"
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
         Left            =   60
         TabIndex        =   34
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label LblNomeCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1740
         TabIndex        =   33
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente:"
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
         Left            =   1020
         TabIndex        =   32
         Top             =   360
         Width           =   675
      End
      Begin VB.Label LblCodigoInicial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Título Inicial:"
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
         Left            =   480
         TabIndex        =   29
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label LlbCodigoFinal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Título Final:"
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
         Left            =   600
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LblDtDistratoCessao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Distrato:"
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
         Left            =   4500
         TabIndex        =   27
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1980
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaContratos.frx":10D0
      Height          =   3435
      Left            =   60
      TabIndex        =   15
      Top             =   780
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   6059
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Contrato"
      Columns(0).DataField=   "Contrato"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Status"
      Columns(1).DataField=   "Status"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cliente"
      Columns(2).DataField=   "Cliente"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Data da Venda"
      Columns(3).DataField=   "cont_dt_Venda"
      Columns(3).NumberFormat=   "dd/mm/yy"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Data do Contrato"
      Columns(4).DataField=   "cont_dt_RegistroContrato"
      Columns(4).NumberFormat=   "dd/mm/yy"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Data Base"
      Columns(5).DataField=   "cont_dt_Base"
      Columns(5).NumberFormat=   "dd/mm/yy"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Data do Distrato"
      Columns(6).DataField=   "cont_dt_Distrato"
      Columns(6).NumberFormat=   "dd/mm/yy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Data da Quitação"
      Columns(7).DataField=   "cont_dt_Quitacao"
      Columns(7).NumberFormat=   "dd/mm/yy"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2249"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2170"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6509"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6429"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=529"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1799"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=529"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1852"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1905"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1826"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=529"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(27)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.valignment=2"
      _StyleDefs(39)  =   ":id=28,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(71)  =   "Named:id=33:Normal"
      _StyleDefs(72)  =   ":id=33,.parent=0"
      _StyleDefs(73)  =   "Named:id=34:Heading"
      _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   ":id=34,.wraptext=-1"
      _StyleDefs(76)  =   "Named:id=35:Footing"
      _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=36:Selected"
      _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=37:Caption"
      _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(82)  =   "Named:id=38:HighlightRow"
      _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(84)  =   "Named:id=39:EvenRow"
      _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(86)  =   "Named:id=40:OddRow"
      _StyleDefs(87)  =   ":id=40,.parent=33"
      _StyleDefs(88)  =   "Named:id=41:RecordSelector"
      _StyleDefs(89)  =   ":id=41,.parent=34"
      _StyleDefs(90)  =   "Named:id=42:FilterBar"
      _StyleDefs(91)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   780
      Left            =   60
      TabIndex        =   21
      Top             =   5160
      Width           =   9300
      _Version        =   65536
      _ExtentX        =   16404
      _ExtentY        =   1376
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
      BevelOuter      =   0
      Begin VB.CommandButton CmdAssistencia 
         Caption         =   "Assis.Técnica"
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
         Left            =   6840
         TabIndex        =   83
         Top             =   420
         Width           =   1375
      End
      Begin VB.CommandButton CmdTransfEmp 
         Caption         =   "Transferência Entre Empresas"
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
         Left            =   1680
         TabIndex        =   39
         Top             =   420
         Width           =   2715
      End
      Begin VB.CommandButton CmdAlterarTitulosEntrega 
         Caption         =   "Entrega Chaves"
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
         Left            =   7560
         TabIndex        =   47
         Top             =   0
         Width           =   1515
      End
      Begin VB.CommandButton CmdTransferencia 
         BackColor       =   &H00000000&
         Caption         =   "&Transferência"
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
         Left            =   0
         TabIndex        =   38
         Top             =   420
         Width           =   1515
      End
      Begin VB.CommandButton CmdTitulo 
         Caption         =   "&Títulos"
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
         Left            =   6360
         TabIndex        =   5
         Top             =   0
         Width           =   1095
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
         Left            =   1560
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdInserir 
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
         Left            =   360
         TabIndex        =   0
         Top             =   0
         Width           =   1095
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
         Left            =   8280
         TabIndex        =   8
         Top             =   420
         Width           =   975
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
         Left            =   2760
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdMapaFechamento 
         BackColor       =   &H00000000&
         Caption         =   "&Mapa"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   420
         Width           =   915
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
         Left            =   5600
         TabIndex        =   7
         Top             =   420
         Width           =   1155
      End
      Begin VB.CommandButton CmdCessao 
         Caption         =   "&Cessão"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdDistratar 
         Caption         =   "&Distratar"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin Threed.SSPanel PanCodigo 
      Height          =   630
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   1111
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
      Begin VB.Label LblCodImovel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8220
         TabIndex        =   14
         Top             =   180
         Width           =   735
      End
      Begin VB.Label LblNomeEmpreendimento 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   180
         Width           =   3735
      End
      Begin VB.Label LblCodEmpreendimento 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   180
         Width           =   675
      End
      Begin VB.Label LblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Código do Imóvel:"
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
         Left            =   6540
         TabIndex        =   11
         Top             =   180
         Width           =   1605
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
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   180
         Width           =   1485
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaContratos.frx":10E5
      Height          =   315
      Left            =   2700
      TabIndex        =   84
      Top             =   0
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "titulo"
      BoundColumn     =   "titulo"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Código:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   85
      Top             =   60
      Width           =   2670
   End
End
Attribute VB_Name = "TelaManContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFT_CODIMOVEL As String 'Código do imóvel
Dim XFT_CODEMPREENDIMENTO As String 'Código do empreendimento
Dim XFT_CODCONTRATO As String 'Código do Contrato
Dim XFT_CODNOVOCONTRATO As String
Dim XLO_QUALIFIN As ADODB.Recordset
Dim XLL_FATURA As Long

Sub SubExcluiFinanceiro()
    
    'Filtrar a empresa
    'Excluir NotaFiscal nofi_nr_titulocapi
    'Excluir Duplicata nofi_nr_fatura/dupl_nr_fatura -
    'Excluir DuplDescAcresc dupl_cd_duplicata
    'Excluir Distribuicao nofi_cd_notafiscal
    
    
    Dim XLT_TITULOQUALIFIN As String
    
    XLT_TITULOQUALIFIN = Adodc1.Recordset.Fields("empd_cd_Empreendimento") + "." + Adodc1.Recordset.Fields("imov_cd_Imovel") + "." + Adodc1.Recordset.Fields("cont_cd_Contrato")
        
    'Já exclui na tabela de descacresc de duplicatas pois o relacionamento com duplicatas é cascade
    Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura " & _
                    " IN ( SELECT nofi_nr_fatura FROM NotasFiscais where nofi_nr_titulocapi LIKE '" & XLT_TITULOQUALIFIN & _
                        "%' AND nofi_tx_status <> 'R' AND empr_cd_empresa = " & PCodEmpresa & ")"

    Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal IN ( SELECT nofi_cd_notafiscal FROM NotasFiscais where nofi_nr_titulocapi LIKE '" & XLT_TITULOQUALIFIN & _
                        "%' AND nofi_tx_status <> 'R' AND empr_cd_empresa = " & PCodEmpresa & ")"
    
    Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_titulocapi LIKE '" & XLT_TITULOQUALIFIN & _
                    "%' AND  nofi_tx_status <> 'R'" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
               
End Sub
Private Sub subHabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdDistratar.Enabled = True
    CmdCessao.Enabled = True
    CmdAssistencia.Enabled = True
    CmdTitulo.Enabled = True
    CmdMapaFechamento.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    
End Sub

Private Sub subDesabilitaBotoes()
    
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdDistratar.Enabled = False
    CmdCessao.Enabled = False
    CmdAssistencia.Enabled = False
    CmdTitulo.Enabled = False
    CmdMapaFechamento.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    
End Sub

Sub subAjustaTelaContrato()
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 1380
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.PanCodigo.Left = 75
    Formulario.PanCodigo.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Top = PanCodigo.Top + PanCodigo.Height + 40
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 100
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
End Sub

Function funChecarCampos() As Boolean

  funChecarCampos = False

  If Len(Trim(MaskEmpree.Text)) < 4 Then
    MsgBox "O novo código do empreendimento deve possuir 4 dígitos!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  
  If Len(Trim(MaskImovel.Text)) < 9 Then
    MsgBox "O novo código do imóvel deve possuir 9 dígitos!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  
  If CboEmpresa.Text = "" Then
    MsgBox "Selecione a Empresa Destino!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  
  If CboPlanoContas.Text = "" Then
    MsgBox "Selecione a Conta Contábil Destino!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  
  If CboPcr.Text = "" Then
    MsgBox "Selecione o Centro de Resultado Destino!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  
  If CboCCorrente.Text = "" Then
    MsgBox "Selecione a Conta Corrente Destino!", vbCritical, "ATENÇÃO"
    Exit Function
  End If
  funChecarCampos = True
End Function

Private Sub CboEmpresa_LostFocus()
  
  Dim XLT_EMPRESADESTINO As String
  
  If Not IsNull(DatEmpresa.Recordset.Fields!empr_cd_Empresa) Then
    DatEmpresa.Recordset.bookmark = CboEmpresa.SelectedItem
    XLT_EMPRESADESTINO = DatEmpresa.Recordset.Fields!empr_cd_Empresa
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & XLT_EMPRESADESTINO & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & XLT_EMPRESADESTINO & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & XLT_EMPRESADESTINO & " ORDER BY coco_tx_Descricao", Estatico
  End If
    
  If CboPlanoContas <> "" Or MskConta <> "" Or MskPcr <> "" Or CboPcr <> "" Or CboCCorrente <> "" Then
    CboPlanoContas = ""
    CboCCorrente = ""
    MskConta = ""
    MskPcr = ""
    CboPcr = ""
  End If
    
End Sub

Private Sub CboPcr_Change()
If CboPcr.Text <> "" Then
    DatPcr.Recordset.bookmark = CboPcr.SelectedItem
    If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
      MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
      CboPcr.BoundText = MskPcr.Text
      CboPcr.SetFocus
      Exit Sub
    End If
    MskPcr.Text = CboPcr.BoundText
  End If
End Sub

Private Sub CboPlanoContas_Change()
    If CboPlanoContas.Text <> "" Then
        DatConta.Recordset.bookmark = CboPlanoContas.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = MskConta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        MskConta.Text = CboPlanoContas.BoundText
    End If
End Sub

Private Sub ChkExportaFinanceiro_Click()
    If ChkExportaFinanceiro.Value = 1 Then
        chkIncluirPagos.Visible = True
    Else
        chkIncluirPagos.Visible = False
        chkIncluirPagos.Value = 0
    End If
    
End Sub

'Desabilita campos quando o usuário selecionar a opção de não copiar títulos do contrato de origem
Private Sub chkNaoCopiarTitulos_Click()
    If chkNaoCopiarTitulos.Value = 1 Then
       'Desabilita os campos para inserção do intervalo dos títulos a serem transferidos
        Label1.Enabled = False
        Label2.Enabled = False
        mskTituloInicial.Enabled = False
        mskTituloFinal.Enabled = False
       
       'Desabilita as opções de operações sobre os títulos
        ChkExportaFinanceiro.Enabled = False
        chkExcluirTitulosOrigem.Enabled = False
        chkAtualizaValorContratoOrigem.Enabled = False
        
        'Limpa Campos:
        mskTituloInicial.Text = "  .   .  "
        mskTituloFinal.Text = "  .   .  "
        ChkExportaFinanceiro.Value = 0
        chkExcluirTitulosOrigem.Value = 0
        chkAtualizaValorContratoOrigem.Value = 0
    Else
      'Habilita os campos para inserção do intervalo dos títulos a serem transferidos
        Label1.Enabled = True
        Label2.Enabled = True
        mskTituloInicial.Enabled = True
        mskTituloFinal.Enabled = True
        
       'Habilita as opções de operações sobre os títulos
        ChkExportaFinanceiro.Enabled = True
        chkExcluirTitulosOrigem.Enabled = True
        chkAtualizaValorContratoOrigem.Enabled = True
    End If
End Sub

Private Sub chkExcluirTitulosOrigem_Click()
    If chkExcluirTitulosOrigem.Value = 1 Then
        chkAtualizaValorContratoOrigem.Visible = True
    Else
        chkAtualizaValorContratoOrigem.Value = 0
        chkAtualizaValorContratoOrigem.Visible = False
    End If
End Sub

Private Sub CmdAlterar_Click()
    Dim XLO_Contrato As New ADODB.Recordset 'Pega o contrato selecionado 17/02/09 - Patrícia
    Dim XLO_PCR As New ADODB.Recordset 'Pega o centro de custo do contrato selecionado 17/02/09 - Patrícia
    
     SubQOpenRecordset XLO_Contrato, "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'", "Estatico" 'inserido em 17/02/09 - Patrícia
     If Not (XLO_Contrato.EOF) And Not IsNull(XLO_Contrato!cere_cd_Pcr) Then
          SubQOpenRecordset XLO_PCR, "SELECT *, LEFT(cere_tx_Descricao + SPACE(40), 40) + ' - ' + cere_cd_Estruturado AS DESCCOMBO FROM CentroResultados WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_cd_Pcr = " & XLO_Contrato!cere_cd_Pcr & " ORDER BY cere_cd_estruturado", Estatico
          If Not (XLO_PCR.EOF) Then
            If XLO_PCR!cere_tx_status = "I" Then
                MsgBox "O Centro de Custo " & XLO_PCR!cere_cd_Estruturado & " que está sendo utilizado pelo contrato, está inativo. Será necessário ativar o centro de custo para realizar as alterações do contrato. Para ativar o centro de custo, acesse a tela de cadastro do centro de custo e altere o status para ATIVO.", vbOKOnly, "ATENÇÃO"
                Exit Sub
            End If
          End If
     End If
    
    If Not FunVerifica_Permissao(Me.Name, "3") Then
      Exit Sub
    End If
    
    Dim XLI_CONT As Integer
    Dim XLB_ENTRA As Boolean
    
    XLB_ENTRA = True
    
    For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
      If TDBGrid1.Columns(1).Text = "RESERVADO" Then
        If MsgBox("Deseja alterar essa reserva para um contrato ATIVO?", vbYesNo, "ATENÇÃO!") = vbNo Then
          XLB_ENTRA = False
        End If
      End If
    Next
    
    If XLB_ENTRA = True Then
      Chave = Adodc1.Recordset.Fields("Contrato")
      XInserir = "A"
      FrmManContratos.Show 1
      subCarregaDadosNV Adodc1, "", "Contrato", _
            "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    End If
End Sub

Private Sub CmdAlterarTitulosEntrega_Click()
    If Not FunVerifica_Permissao(Me.Name, "9") Then
      Exit Sub
    End If
    
    PanBotoes.Enabled = False
    TDBGrid1.Enabled = False
    
    Chave = Adodc1.Recordset!contrato
    FrmEntregaChaves.Show 1
      
    PanBotoes.Enabled = True
    TDBGrid1.Enabled = True
End Sub

Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
    
    If XLT_OPERACAO = ATIVAR_CONTRATO Or XLT_OPERACAO = EXCLUIR Then
    
        ReDim XGM_MATRIZLOG(3, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Empreendimento"
        XGM_MATRIZLOG(1, 0) = "Imóvel"
        XGM_MATRIZLOG(2, 0) = "Contrato"
                        
        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("empd_cd_Empreendimento")
        XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("imov_cd_imovel")
        XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("cont_cd_Contrato")
        
    ElseIf XLT_OPERACAO = DISTRATAR_CONTRATO Then
        
        ReDim XGM_MATRIZLOG(5, 2) As Variant

        XGM_MATRIZLOG(0, 0) = "Empreendimento"
        XGM_MATRIZLOG(1, 0) = "Imóvel"
        XGM_MATRIZLOG(2, 0) = "Contrato"
        XGM_MATRIZLOG(3, 0) = "Data Distrato"
        XGM_MATRIZLOG(4, 0) = "Limpar Valor do Imóvel e das Garagens?"
        
        XGM_MATRIZLOG(0, 1) = XFT_CODEMPREENDIMENTO
        XGM_MATRIZLOG(1, 1) = XFT_CODIMOVEL
        XGM_MATRIZLOG(2, 1) = XFT_CODCONTRATO
        XGM_MATRIZLOG(3, 1) = FunNulo(DtpDtDistrato.Value)
        If ChkLimparImovel.Value = 1 Then XGM_MATRIZLOG(4, 1) = "SIM" Else XGM_MATRIZLOG(4, 1) = "NÃO"
    
    ElseIf XLT_OPERACAO = CESSAO_CONTRATO Then
            
        ReDim XGM_MATRIZLOG(11, 2) As Variant

        XGM_MATRIZLOG(0, 0) = "Data de Cessão"
        XGM_MATRIZLOG(1, 0) = "Empreendimento Origem"
        XGM_MATRIZLOG(2, 0) = "Imóvel Origem"
        XGM_MATRIZLOG(3, 0) = "Contrato Origem"
        XGM_MATRIZLOG(4, 0) = "Empreendimento Destino"
        XGM_MATRIZLOG(5, 0) = "Imóvel Destino"
        XGM_MATRIZLOG(6, 0) = "Contrato Destino"
        XGM_MATRIZLOG(7, 0) = "Título Inicial"
        XGM_MATRIZLOG(8, 0) = "Título Final"
        XGM_MATRIZLOG(9, 0) = "Títulos Pagos Copiados?"
        XGM_MATRIZLOG(10, 0) = "Limpar Imóvel?"

        XGM_MATRIZLOG(0, 1) = FunNulo(DtpDtDistratoCessao.Value)
        XGM_MATRIZLOG(1, 1) = XFT_CODEMPREENDIMENTO
        XGM_MATRIZLOG(2, 1) = XFT_CODIMOVEL
        XGM_MATRIZLOG(3, 1) = XFT_CODCONTRATO
        XGM_MATRIZLOG(4, 1) = XFT_CODEMPREENDIMENTO
        XGM_MATRIZLOG(5, 1) = XFT_CODIMOVEL
        XGM_MATRIZLOG(6, 1) = Chave
        XGM_MATRIZLOG(7, 1) = CboContratoInicial.BoundText
        XGM_MATRIZLOG(8, 1) = CboContratoFinal.BoundText
        XGM_MATRIZLOG(9, 1) = IIf(ChkTitulosPagos.Value = 1, "SIM", "NÃO")
        XGM_MATRIZLOG(10, 1) = IIf(ChkLimparImovel.Value = 1, "SIM", "NÃO")
                
    ElseIf XLT_OPERACAO = TRANSFERENCIA Then
       
        ReDim XGM_MATRIZLOG(6, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Empreendimento Origem"
        XGM_MATRIZLOG(1, 0) = "Imóvel Origem"
        XGM_MATRIZLOG(2, 0) = "Contrato Origem"
        XGM_MATRIZLOG(3, 0) = "Empreendimento Destino"
        XGM_MATRIZLOG(4, 0) = "Imóvel Destino"
        XGM_MATRIZLOG(5, 0) = "Contrato Destino"
                        
        XGM_MATRIZLOG(0, 1) = XFT_CODEMPREENDIMENTO
        XGM_MATRIZLOG(1, 1) = XFT_CODIMOVEL
        XGM_MATRIZLOG(2, 1) = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
        XGM_MATRIZLOG(3, 1) = Left(MskCodigo.Text, 4)
        XGM_MATRIZLOG(4, 1) = Right(MskCodigo.Text, 4)
        XGM_MATRIZLOG(5, 1) = XFT_CODNOVOCONTRATO
        
    ElseIf XLT_OPERACAO = TRANSFERENCIA_ENTRE_EMPRESAS Then
        
        ReDim XGM_MATRIZLOG(13, 2) As Variant

        XGM_MATRIZLOG(0, 0) = "Empresa Origem"
        XGM_MATRIZLOG(1, 0) = "Empreendimento Origem"
        XGM_MATRIZLOG(2, 0) = "Imóvel Origem"
        XGM_MATRIZLOG(3, 0) = "Empresa Destino"
        XGM_MATRIZLOG(4, 0) = "Empreendimento Destino"
        XGM_MATRIZLOG(5, 0) = "Imóvel Destino"
        XGM_MATRIZLOG(6, 0) = "Conta Contábil Destino"
        XGM_MATRIZLOG(7, 0) = "Centro de Custo Destino"
        XGM_MATRIZLOG(8, 0) = "Conta Corrente Destino"
        XGM_MATRIZLOG(9, 0) = "Títulos Copiados para o Contrato de Destino?"
        XGM_MATRIZLOG(10, 0) = "Títulos Exportados para o QualiFin?"
        XGM_MATRIZLOG(11, 0) = "Títulos Originais Excluídos?"
        XGM_MATRIZLOG(12, 0) = "Valor do Contrato Original Atualizado?"

        XGM_MATRIZLOG(0, 1) = PCodEmpresa
        XGM_MATRIZLOG(1, 1) = Left(LblImovelOrigemEmpr.Caption, 4)
        XGM_MATRIZLOG(2, 1) = Right(LblImovelOrigemEmpr.Caption, 4)
        XGM_MATRIZLOG(3, 1) = DatEmpresa.Recordset.Fields!empr_cd_Empresa
        XGM_MATRIZLOG(4, 1) = Left(MaskImovel.Text, 4)
        XGM_MATRIZLOG(5, 1) = Right(MaskImovel.Text, 4)
        XGM_MATRIZLOG(6, 1) = FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil"))
        XGM_MATRIZLOG(7, 1) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
        XGM_MATRIZLOG(8, 1) = DatContaCorrente.Recordset.Fields!coco_cd_codigo
        If chkNaoCopiarTitulos.Value = 1 Then XGM_MATRIZLOG(9, 1) = "SIM" Else XGM_MATRIZLOG(9, 1) = "NÃO"
        If ChkExportaFinanceiro.Value = 1 Then XGM_MATRIZLOG(10, 1) = "SIM" Else XGM_MATRIZLOG(10, 1) = "NÃO"
        If chkExcluirTitulosOrigem.Value = 1 Then XGM_MATRIZLOG(11, 1) = "SIM" Else XGM_MATRIZLOG(11, 1) = "NÃO"
        If chkAtualizaValorContratoOrigem.Value = 1 Then XGM_MATRIZLOG(12, 1) = "SIM" Else XGM_MATRIZLOG(12, 1) = "NÃO"
                
    End If
    
    Call subRegistraLog("TelaManContratos", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, CONTRATO_CAPI))
End Sub

Private Sub CmdAssistencia_Click()
If Not FunVerifica_Permissao("TelaFinTitulos", "1") Then
      Exit Sub
  End If
  
  XOrigem = "TelaAssistencia"
  Chave = Adodc1.Recordset.Fields("Contrato")
  TelaAssistencia.Show
  
  If FunLocalizaJanelasAtivas("TelaAssistencia") <> -1 Then
      TelaAssistencia.SetFocus
  End If
End Sub

Private Sub CmdConfirmarTransEmp_Click()
  
    If funChecarCampos = True Then
        
        Dim XLO_CODIGO As New ADODB.Recordset
        Dim XLT_SQL As String
        Dim XLT_CONTRATO As String
        Dim XLT_EMPRESADESTINO As String
        Dim XLT_CONTACONTABIL As String
        Dim XLT_PCR As String
        Dim XLT_CONTACORRENTE As String
        Dim XLT_FORNCLIENTES As String
        Dim XLO_TITULO As New ADODB.Recordset
        Dim XLT_NATUREZA As String
        Dim XLO_TIPOPLANO As New ADODB.Recordset
        Dim XLT_TIPOPLANO As String
        Dim XLT_TITULOQUALIFIN As String 'Armazena o código do título para o Qualifin
        Dim XLO_QUALIFIN As Object 'Recordset para interligar ao qualifin
        Dim XLD_VALORCONTRATO As Double 'Armazena a soma de todos os títulos selecionados para a transferência
        Dim XLD_VALORPLANO As Double 'Armazena a soma dos valores dos títulos selecionados pertencentes a um mesmo plano
        Dim XLB_SOMAVALORCONTRATO As Boolean 'Determina se é para ser calculado o valor do contrato
        Dim XLB_VALORPLANOSOMADO As Boolean 'Informa se já foram somados todos os títulos de um plano
        Dim XLT_CAMPOPARCELA As String 'Define o campo valor da parcela na tabela Contratos que será atualizado
        Dim XLI_COUNT As Integer 'Contador de títulos a serem transferidos
        Dim XLO_Contrato As New ADODB.Recordset
                
        Conexao.BeginTrans
        
        'Soma valor do contrato apenas se for transferida uma parte dos títulos:
        If (mskTituloInicial.Text = "  .   .  ") And (mskTituloFinal.Text = "  .   .  ") Then
          XLB_SOMAVALORCONTRATO = False
        Else
          XLB_SOMAVALORCONTRATO = True
        End If
        
        XLD_VALORCONTRATO = 0
        XLD_VALORPLANO = 0
        XLI_COUNT = 0
        XLT_NATUREZA = ""
        XLT_CAMPOPARCELA = ""
        XLB_VALORPLANOSOMADO = False
        
        'Selecionando os códigos
        DatEmpresa.Recordset.bookmark = CboEmpresa.SelectedItem
        If Not IsNull(DatEmpresa.Recordset.Fields!empr_cd_Empresa) Then
            XLT_EMPRESADESTINO = DatEmpresa.Recordset.Fields!empr_cd_Empresa
        Else
            XLT_EMPRESADESTINO = ""
        End If
        
        DatConta.Recordset.bookmark = CboPlanoContas.SelectedItem
        XLT_CONTACONTABIL = DatConta.Recordset.Fields!plco_cd_Conta
        
        DatPcr.Recordset.bookmark = CboPcr.SelectedItem
        XLT_PCR = DatPcr.Recordset.Fields!cere_cd_Pcr
        
        DatContaCorrente.Recordset.bookmark = CboCCorrente.SelectedItem
        XLT_CONTACORRENTE = DatContaCorrente.Recordset.Fields!coco_cd_codigo
        
        
        'Verifica se o novo empreendimento já existe na empresa destino
        XLT_SQL = "SELECT empd_cd_Empreendimento FROM Imoveis WHERE " & _
                " empd_cd_Empreendimento ='" & Left(MaskEmpree.Text, 4) & "'" & _
                " AND empr_cd_empresa =" & XLT_EMPRESADESTINO
        SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
        
        If XLO_CODIGO.EOF Then
            XLO_CODIGO.Close
            Set XLO_CODIGO = Nothing
            Conexao.RollbackTrans
            
            MsgBox "O empreendimento do novo código não existe!", vbCritical, "ATENÇÃO"
            Exit Sub
        Else
            XLO_CODIGO.Close
        End If
        
        
        'Verifica se o novo imóvel já existe na empresa destino
        XLT_SQL = "SELECT empd_cd_Empreendimento FROM Imoveis WHERE " & _
                " empd_cd_Empreendimento ='" & Left(MaskImovel.Text, 4) & "'" & _
                " AND imov_cd_Imovel = '" & Right(MaskImovel.Text, 4) & "'" & _
                " AND empr_cd_empresa =" & XLT_EMPRESADESTINO
        SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
        
        If XLO_CODIGO.EOF Then
            XLO_CODIGO.Close
            Set XLO_CODIGO = Nothing
            Conexao.RollbackTrans
            MsgBox "O imóvel do novo código não existe!", vbCritical, "ATENÇÃO"
            Exit Sub
        Else
            XLO_CODIGO.Close
        End If
        
        'Se o usuário clicou em "excluir títulos da origem", mas não clicou em "exportar para o qualifin", avisar que as notas referentes aos títulos serão excluídas.
        If chkExcluirTitulosOrigem.Value = 1 And ChkExportaFinanceiro.Value = 0 Then
            If MsgBox("Ao excluir os títulos da origem também serão excluídas as notas fiscais correspondentes no QualiFin. Deseja continuar?", vbExclamation, vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
       
        'Abre o recordset do contrato da empresa origem
        XLT_SQL = "SELECT * FROM Contratos WHERE " & _
                " empd_cd_Empreendimento ='" & Left(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND imov_cd_Imovel = '" & Right(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND cont_tx_status = 'A' " & _
                " AND empr_cd_empresa =" & PCodEmpresa
        SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
        
        
        Do While Not XLO_CODIGO.EOF
            XLT_CONTRATO = XLO_CODIGO!cont_cd_Contrato
            XLT_FORNCLIENTES = XLO_CODIGO!focl_cd_FornCli
            
            Sql = "INSERT INTO Contratos (cont_cd_Contrato,imov_cd_imovel, empd_cd_empreendimento, " & _
                    "empr_cd_Empresa, cere_cd_Pcr, plco_cd_conta, moed_cd_Codmoeda, " & _
                    "cont_vl_Avista, cont_vl_sinal, cont_vl_Poupanca, cont_vl_Total, " & _
                    "cont_vl_Fgts, cont_vl_Financiamento, cont_vl_Outros, cont_dt_base, " & _
                    "cont_dt_RegistroContrato, cont_dt_EntregaChaves, cont_vl_TaxaSegVida, " & _
                    "cont_dt_Distrato,focl_cd_FornCli, cont_tx_MultaProRata, cont_tx_CorrecProRata, " & _
                    "cont_nr_TipoSeguro, cont_tx_TipoJurosMora, cont_tx_Desagio, cont_vl_PercJurosFin, " & _
                    "cont_vl_PercJurosMulta, cont_vl_PercJurosMora, cont_tx_Status, cont_cd_ContratoOrigem, " & _
                    "cont_dt_Venda, cont_dt_Quitacao, cont_dt_LimiteReserva, iden_cd_Projeto, coco_cd_Codigo, cont_nr_LimiteDias) " & _
                    " VALUES ('" & XLT_CONTRATO & "','" & Right(MaskImovel.Text, 4) & "','" & Left(MaskImovel.Text, 4) & "','" & XLT_EMPRESADESTINO & "','" & _
                    "" & XLT_PCR & "','" & XLT_CONTACONTABIL & "'," & FunNuloBanco(XLO_CODIGO!moed_cd_CodMoeda) & ","
                  
            If XLB_SOMAVALORCONTRATO = True Then
               'Se for para somar o valor da cópia do contrato, inserir o contrato com os campos referente a valores do contrato = 0, p/ que sejam atualizados a seguir.
                Sql = Sql & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ","
            Else
               'Senão, insere os valores do contrato original
                Sql = Sql + "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Avista)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Sinal)) & "," & _
                    "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Poupanca)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Total)) & "," & _
                    "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Fgts)) & "," & FunTrataFloat(XLO_CODIGO!cont_vl_Financiamento) & "," & _
                    "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_Outros)) & ","
            End If
            
            Sql = Sql + FunNuloData(XLO_CODIGO!cont_dt_Base, NomeSgbd) & "," & _
                  "" & FunNuloData(XLO_CODIGO!cont_dt_RegistroContrato, NomeSgbd) & "," & FunNuloData(XLO_CODIGO!cont_dt_EntregaChaves, NomeSgbd) & ",'" & _
                  "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_TaxaSegVida)) & "'," & FunNuloData(XLO_CODIGO!cont_dt_Distrato, NomeSgbd) & "," & FunNuloBanco(XLO_CODIGO!focl_cd_FornCli) & ",'" & _
                  "" & XLO_CODIGO!cont_tx_MultaProRata & "','" & XLO_CODIGO!cont_tx_CorrecProRata & "','" & _
                  "" & XLO_CODIGO!cont_nr_TipoSeguro & "','" & XLO_CODIGO!cont_tx_TipoJurosMora & "','" & _
                  "" & XLO_CODIGO!cont_tx_Desagio & "'," & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_PercJurosFin)) & "," & _
                  "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_PercJurosMulta)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!cont_vl_PercJurosMora)) & ",'" & _
                  "" & XLO_CODIGO!cont_tx_Status & "'," & FunNuloBanco(XLO_CODIGO!cont_cd_ContratoOrigem) & "," & _
                  "" & FunNuloData(XLO_CODIGO!cont_dt_Venda, NomeSgbd) & "," & FunNuloData(XLO_CODIGO!cont_dt_Quitacao, NomeSgbd) & "," & FunNuloData(XLO_CODIGO!cont_dt_LimiteReserva, NomeSgbd) & "," & FunNuloBanco(XLO_CODIGO!iden_cd_Projeto) & ",'" & _
                  "" & XLT_CONTACORRENTE & "'," & FunNuloBanco(XLO_CODIGO!cont_nr_LimiteDias) & ") "
            
            Conexao.Execute Sql
            XLO_CODIGO.MoveNext
        Loop
        
        XLO_CODIGO.Close
        Set XLO_CODIGO = Nothing
        
        'Inserindo na tabela de compradores
        Sql = "INSERT INTO Compradores (focl_cd_FornCli, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, empr_cd_Empresa)" & _
                " VALUES ('" & XLT_FORNCLIENTES & "','" & Left(MaskImovel.Text, 4) & "','" & _
                "" & Right(MaskImovel.Text, 4) & "','" & XLT_CONTRATO & "','" & XLT_EMPRESADESTINO & "')"
        Conexao.Execute Sql
              
        'Início - Se o usuário não marcou a opção "Não copiar títulos", faz a transferência dos títulos.
        If chkNaoCopiarTitulos.Value <> 1 Then
        
            'Abre o recordset do titulos da empresa origem e insere o titulos na empresa destino
            
            If chkIncluirPagos.Value = 0 Then
                XLT_SQL = _
                "SELECT * FROM " & _
                "   titulos " & _
                "WHERE " & _
                " empd_cd_Empreendimento ='" & Left(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND imov_cd_Imovel = '" & Right(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND (titu_dt_Pagamento is null or titu_dt_Pagamento = '')  " & _
                " AND (cont_cd_contrato = '" & XLT_CONTRATO & "')" & _
                " AND empr_cd_empresa =" & PCodEmpresa
            Else
                XLT_SQL = _
                "SELECT * FROM " & _
                "   titulos " & _
                "WHERE " & _
                " empd_cd_Empreendimento ='" & Left(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND imov_cd_Imovel = '" & Right(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                " AND (cont_cd_contrato = '" & XLT_CONTRATO & "')" & _
                " AND empr_cd_empresa =" & PCodEmpresa
             End If
                      
              'CONSULTA OS TITULOS PELA FAIXA INFORMADA
              If (mskTituloInicial.Text <> "  .   .  ") Then
                  XLT_SQL = XLT_SQL & _
                      " AND titu_cd_Plano >= " & Mid(mskTituloInicial.Text, 1, 2) & _
                      " AND titu_cd_Parcela >= " & Mid(mskTituloInicial.Text, 4, 3)
              End If
              
              'Se os títulos pertencem a mais de um plano, inclui a condição:
              If (mskTituloFinal.Text <> "  .   .  ") And (Val(Mid(mskTituloFinal.Text, 1, 2)) - Val(Mid(mskTituloInicial.Text, 1, 2)) > 0) Then
                  XLT_SQL = XLT_SQL & " AND titu_cd_Plano <= " & Mid(mskTituloFinal.Text, 1, 2)
              'Se os títulos pertencem a um único plano, inclui a condição:
              ElseIf (mskTituloFinal.Text <> "  .   .  ") And (Val(Mid(mskTituloFinal.Text, 1, 2)) - Val(Mid(mskTituloInicial.Text, 1, 2)) = 0) Then
                  XLT_SQL = XLT_SQL & _
                      " AND titu_cd_Plano <= " & Mid(mskTituloFinal.Text, 1, 2) & _
                      " AND titu_cd_Parcela <= " & Mid(mskTituloFinal.Text, 4, 3)
              End If
                      
              XLT_SQL = XLT_SQL + " ORDER BY titu_cd_plano"
                      
              SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
            
            'Insere os títulos do contrato da empresa de origem no novo contrato da empresa de destino.
            Do While Not XLO_CODIGO.EOF
                If ((mskTituloFinal.Text) = "  .   .  ") Or (((mskTituloFinal.Text) <> "  .   .  ") And Not ((XLO_CODIGO!titu_cd_Plano = Mid(mskTituloFinal.Text, 1, 2)) And (XLO_CODIGO!titu_cd_Parcela > Mid(mskTituloFinal.Text, 4, 3)))) Then
                        
                    XLI_COUNT = XLI_COUNT + 1
                    
                    Sql = "INSERT INTO Titulos (titu_cd_Plano,titu_cd_Parcela, titu_cd_Residuo, " & _
                        "empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, " & _
                        "tipl_cd_TipoPlano, coco_cd_Codigo, obse_cd_Observacao, moed_cd_Moeda1, " & _
                        "moed_cd_Moeda2, titu_vl_Parcela, titu_vl_ParcelaIndex1, " & _
                        "titu_vl_ParcelaIndex2, titu_dt_Vencimento,titu_vl_ValorPago,titu_dt_Pagamento, titu_vl_Multa, titu_vl_Juros, " & _
                        "titu_vl_Seguro, titu_tx_IndicePagamento, titu_vl_Outros, titu_vl_Desconto, titu_vl_Desagio, titu_tx_Financia, titu_vl_PercJurosFin, titu_vl_PercMultaMora, " & _
                        "titu_vl_PercJurosMora, titu_dt_Deposito, titu_dt_Repasse, titu_tx_LocalPagto, titu_tx_DocPagto, titu_dt_Prorrogacao, " & _
                        "titu_tx_TipoPag, titu_dt_Base, cere_cd_Pcr, plco_cd_Conta, titu_tx_NaturezaPlano, moed_cd_Moeda3, titu_tx_CorrigeParcela, titu_tx_PreDatado, titu_dt_BasePagto, " & _
                        "titu_dt_Criacao, nofi_cd_NotaFiscal, iden_cd_Projeto, titu_tx_ExpBanco, titu_tx_EntregaChave) " & _
                        "VALUES ('" & XLO_CODIGO!titu_cd_Plano & "','" & XLO_CODIGO!titu_cd_Parcela & "','" & XLO_CODIGO!titu_cd_residuo & "','" & XLT_EMPRESADESTINO & "','" & Left(MaskImovel.Text, 4) & "','" & Right(MaskImovel.Text, 4) & "','" & _
                        "" & XLO_CODIGO!cont_cd_Contrato & "','" & XLO_CODIGO!tipl_cd_TipoPlano & "','" & XLT_CONTACORRENTE & "'," & FunNuloBanco(XLO_CODIGO!obse_cd_Observacao) & "," & _
                        "" & FunNuloBanco(XLO_CODIGO!moed_cd_Moeda1) & "," & FunNuloBanco(XLO_CODIGO!moed_cd_Moeda2) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Parcela)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_ParcelaIndex1)) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_ParcelaIndex2)) & "," & _
                        "" & FunNuloData(XLO_CODIGO!titu_dt_Vencimento, NomeSgbd) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_ValorPago)) & "," & FunNuloData(XLO_CODIGO!titu_dt_Pagamento, NomeSgbd) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Multa)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Juros)) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Seguro)) & ",'" & (XLO_CODIGO!titu_tx_IndicePagamento) & "'," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Outros)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Desconto)) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Desagio)) & ",'" & (XLO_CODIGO!titu_tx_Financia) & "'," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_PercJurosFin)) & "," & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_PercMultaMora)) & "," & _
                        "" & FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_PercJurosMora)) & "," & FunNuloData(XLO_CODIGO!titu_dt_Deposito, NomeSgbd) & "," & FunNuloData(XLO_CODIGO!titu_dt_Repasse, NomeSgbd) & ",'" & (XLO_CODIGO!titu_tx_LocalPagto) & "','" & _
                        "" & (XLO_CODIGO!titu_tx_DocPagto) & "'," & FunNuloData(XLO_CODIGO!titu_dt_Prorrogacao, NomeSgbd) & ",'" & (XLO_CODIGO!titu_tx_TipoPag) & "'," & FunNuloData(XLO_CODIGO!titu_dt_Base, NomeSgbd) & ",'" & _
                        "" & XLT_PCR & "','" & XLT_CONTACONTABIL & "','" & XLO_CODIGO!titu_tx_NaturezaPlano & "'," & FunNuloBanco(XLO_CODIGO!moed_cd_Moeda3) & ",'" & _
                        "" & XLO_CODIGO!titu_tx_CorrigeParcela & "','" & XLO_CODIGO!titu_tx_PreDatado & "'," & FunNuloData(XLO_CODIGO!titu_dt_BasePagto, NomeSgbd) & "," & FunNuloData(XLO_CODIGO!titu_dt_Criacao, NomeSgbd) & "," & FunNuloBanco(XLO_CODIGO!nofi_cd_NotaFiscal) & "," & FunNuloBanco(XLO_CODIGO!iden_cd_Projeto) & ",'" & XLO_CODIGO!titu_tx_ExpBanco & "','" & XLO_CODIGO!titu_tx_EntregaChave & "') "
                    Conexao.Execute Sql
                
                    'Define a natureza do plano do título atual:
                    Select Case XLO_CODIGO!titu_tx_NaturezaPlano
                        Case "S"
                            If XLT_NATUREZA <> "SINAL" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_sinal"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "SINAL"
                            
                        Case "P"
                            If XLT_NATUREZA <> "POUPANCA" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_poupanca"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "POUPANCA"
                            
                        Case "F"
                            If XLT_NATUREZA <> "FINANCIAMENTO" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_financiamento"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "FINANCIAMENTO"
                            
                        Case "G"
                            If XLT_NATUREZA <> "FGTS" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_fgts"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "FGTS"
                            
                        Case "O"
                            If XLT_NATUREZA <> "OUTROS" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_outros"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "OUTROS"
                            
                        Case "V"
                            If XLT_NATUREZA <> "A VISTA" Then
                                XLB_VALORPLANOSOMADO = True
                            Else
                                XLT_CAMPOPARCELA = "cont_vl_avista"
                                XLB_VALORPLANOSOMADO = False
                            End If
                            
                            XLT_NATUREZA = "A VISTA"
                            
                    End Select
                    
                    
                    If (XLB_SOMAVALORCONTRATO = True) Then
                        
                        'Soma valor total do contrato:
                        XLD_VALORCONTRATO = XLD_VALORCONTRATO + XLO_CODIGO!titu_vl_Parcela
                            
                        If ((XLB_VALORPLANOSOMADO = True) Or (XLO_CODIGO.RecordCount = XLI_COUNT)) And (XLI_COUNT <> 1) Then
                            
                            'Soma a última parcela
                            If XLO_CODIGO.RecordCount = XLI_COUNT Then
                                XLD_VALORPLANO = XLD_VALORPLANO + XLO_CODIGO!titu_vl_Parcela
                            End If
                            
                            ' Função trata Float - Modificado 15/03/2012  - Daniel Coelho
                            'Atualiza os valores dos contratos dos empreendimentos de origem e destino
                            If XLT_CAMPOPARCELA <> "" Then
                                XLT_SQL = "UPDATE Contratos SET " & _
                                XLT_CAMPOPARCELA & " =" & FunNuloVal(FunTrataFloat(XLD_VALORPLANO)) & _
                                " , cont_vl_Total =" & FunNuloVal(FunTrataFloat(XLD_VALORCONTRATO)) & _
                                " WHERE empd_cd_Empreendimento ='" & Left(MaskImovel.Text, 4) & "'" & _
                                " AND imov_cd_Imovel = '" & Right(MaskImovel.Text, 4) & "'" & _
                                " AND cont_tx_status = 'A' " & _
                                " AND empr_cd_empresa =" & PCodEmpresa
                                
                                Conexao.Execute XLT_SQL
                                
                                'Atualizar o valor do contrato da origem se o usuário selecionou essa opção
                                If chkAtualizaValorContratoOrigem.Value = 1 Then
                                    XLT_SQL = "SELECT " & XLT_CAMPOPARCELA & ", cont_vl_Total FROM Contratos WHERE " & _
                                      " empd_cd_Empreendimento ='" & Left(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                                      " AND imov_cd_Imovel = '" & Right(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                                      " AND cont_tx_status = 'A' " & _
                                      " AND empr_cd_empresa =" & PCodEmpresa
                                    SubQOpenRecordset XLO_Contrato, XLT_SQL, Dinamico
                                    
                                    XLO_Contrato.Fields(XLT_CAMPOPARCELA) = XLO_Contrato.Fields(XLT_CAMPOPARCELA) - XLD_VALORPLANO
                                    XLO_Contrato.Fields("cont_vl_Total") = XLO_Contrato.Fields("cont_vl_Total") - XLD_VALORPLANO
                                    XLO_Contrato.Update
                                End If
                                
                            End If
                                                
                            'Inicializa a soma dos títulos pertencentes a um plano diferente
                            XLD_VALORPLANO = XLO_CODIGO!titu_vl_Parcela
                        
                        Else
                            'Soma valor do plano
                            XLD_VALORPLANO = XLD_VALORPLANO + XLO_CODIGO!titu_vl_Parcela
                        End If
                        
                    End If
                  
                  
                    'Se o usuário selecionou a opção de exportar títulos para o QualiFin
                    If ChkExportaFinanceiro.Value = 1 Then
                        
                        'Se o QualiFin está interligado e a empresa está habilitada para exportar os títulos para o QualiFin:
                        If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then
                    
                            XLT_SQL = "SELECT * FROM Titulos " & _
                                        "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                                        "titu_cd_Parcela='-1' AND " & _
                                        "titu_cd_Residuo='00'"
                            SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
                            
                            XLO_TITULO.AddNew
                        
                            XLO_TITULO!empr_cd_Empresa = XLT_EMPRESADESTINO
                            XLO_TITULO!empd_cd_Empreendimento = Left(MaskImovel.Text, 4)
                            XLO_TITULO!imov_cd_Imovel = Right(MaskImovel.Text, 4)
                            XLO_TITULO!cont_cd_Contrato = XLO_CODIGO!cont_cd_Contrato
                            XLO_TITULO!titu_cd_Plano = XLO_CODIGO!titu_cd_Plano
                            XLO_TITULO!titu_cd_Parcela = XLO_CODIGO!titu_cd_Parcela
                            XLO_TITULO!titu_cd_residuo = XLO_CODIGO!titu_cd_residuo
                            XLO_TITULO!obse_cd_Observacao = funNull(XLO_CODIGO!obse_cd_Observacao)
                            XLO_TITULO!moed_cd_Moeda1 = funNull(XLO_CODIGO!moed_cd_Moeda1)
                            XLO_TITULO!moed_cd_Moeda2 = funNull(XLO_CODIGO!moed_cd_Moeda2)
                            XLO_TITULO!titu_vl_Parcela = FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_Parcela))
                            XLO_TITULO!titu_vl_ParcelaIndex1 = FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_ParcelaIndex1))
                            XLO_TITULO!titu_vl_ParcelaIndex2 = FunNuloVal(FunTrataFloat(XLO_CODIGO!titu_vl_ParcelaIndex2))
                            XLO_TITULO!titu_dt_Vencimento = Format(funNull(XLO_CODIGO!titu_dt_Vencimento), "dd/MM/yy")
                            XLO_TITULO!tipl_cd_TipoPlano = XLO_CODIGO!tipl_cd_TipoPlano
                            XLO_TITULO!titu_vl_PercJurosFin = FunNuloVal(XLO_CODIGO!titu_vl_PercJurosFin)
                            XLO_TITULO!titu_vl_PercMultaMora = FunNuloVal(XLO_CODIGO!titu_vl_PercMultaMora)
                            XLO_TITULO!titu_vl_PercJurosMora = FunNuloVal(XLO_CODIGO!titu_vl_PercJurosMora)
                            XLO_TITULO!cere_cd_Pcr = XLT_PCR
                            XLO_TITULO!plco_cd_Conta = XLT_CONTACONTABIL
                            XLO_TITULO!coco_cd_codigo = XLT_CONTACORRENTE
                            XLO_TITULO!titu_dt_Base = Format(funNull(XLO_CODIGO!titu_dt_Base), "dd/MM/yy")
                            XLO_TITULO!titu_dt_Prorrogacao = Format(funNull(XLO_CODIGO!titu_dt_Prorrogacao))
                            XLO_TITULO!titu_tx_NaturezaPlano = XLO_CODIGO!titu_tx_NaturezaPlano
                            XLO_TITULO!titu_dt_Criacao = Format(funNull(XLO_CODIGO!titu_dt_Criacao))
                            XLO_TITULO!iden_cd_Projeto = FunNuloVal(XLO_CODIGO!iden_cd_Projeto)
                                                                    
                            'Pega a descrição do tipo de plano ao qual o título atual pertence:
                            SubQOpenRecordset XLO_TIPOPLANO, "SELECT * FROM TiposPlanos WHERE tipl_cd_TipoPlano = " & XLO_CODIGO!tipl_cd_TipoPlano, Estatico
                            XLT_TIPOPLANO = XLO_TIPOPLANO!tipl_tx_Descricao
                            

                                                
                            'Grava os títulos no Financeiro:
                            'subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", Val(XLT_FORNCLIENTES), XLT_TIPOPLANO, XLT_NATUREZA
                            'Cássio Medeiros
                            '23/05/2014
                            'Foi necessário adaptar a função acima para passar o empreendimento e imóvel de origem
                            If IsNull(XLO_CODIGO!titu_dt_Pagamento) Or IsEmpty(XLO_CODIGO!titu_dt_Pagamento) Then
                                subGravaFinanceiroTransf XLO_TITULO, "COMPROMISSAR", Val(XLT_FORNCLIENTES), XLT_TIPOPLANO, XLT_NATUREZA, Left(LblImovelOrigemEmpr.Caption, 4), Right(LblImovelOrigemEmpr.Caption, 4)
                            ElseIf Not IsNull(XLO_CODIGO!titu_dt_Pagamento) And Not IsEmpty(XLO_CODIGO!titu_dt_Pagamento) And chkIncluirPagos.Value = 1 Then
                                XLO_TITULO!titu_dt_Pagamento = XLO_CODIGO!titu_dt_Pagamento
                                XLO_TITULO!titu_dt_Deposito = XLO_CODIGO!titu_dt_Deposito
                                XLO_TITULO!titu_vl_ValorPago = FunNuloVal(XLO_CODIGO!titu_vl_ValorPago)
                                XLO_TITULO!titu_vl_Seguro = FunNuloVal(XLO_CODIGO!titu_vl_Seguro)
                                XLO_TITULO!titu_vl_Juros = FunNuloVal(XLO_CODIGO!titu_vl_Juros)
                                XLO_TITULO!titu_vl_Multa = FunNuloVal(XLO_CODIGO!titu_vl_Multa)
                                XLO_TITULO!titu_vl_Outros = FunNuloVal(XLO_CODIGO!titu_vl_Outros)
                                XLO_TITULO!titu_vl_Desagio = FunNuloVal(XLO_CODIGO!titu_vl_Desagio)
                                XLO_TITULO!titu_vl_Desconto = FunNuloVal(XLO_CODIGO!titu_vl_Desconto)
                                subGravaFinanceiroTransf XLO_TITULO, "PAGAR", Val(XLT_FORNCLIENTES), XLT_TIPOPLANO, XLT_NATUREZA, Left(LblImovelOrigemEmpr.Caption, 4), Right(LblImovelOrigemEmpr.Caption, 4)
                            End If
                        Else
                            MsgBox "Não será possível gravar os títulos para o QualiFin pois esta empresa não está habilitada para exportar títulos. Para exportá-los acesse a rotina ' '.", vbInformation, "Exportação de Títulos"
                        End If
                    End If
                    
                    'Se o usuário selecionou a opção de excluir da empresa de origem os títulos transferidos
                    If chkExcluirTitulosOrigem.Value = 1 Then
                    
                        'Exclui títulos do empreendimento de origem do QUALIFIN
                        If XGB_PQUALIFIN Then
                            
                            'Determina o número do título no Qualifin: cod.empreendimento + cod.imovel + cod.contrato + cod.plano + num.parcela + residuo
                            XLT_TITULOQUALIFIN = Left(LblImovelOrigemEmpr, 4) + "." + Right(LblImovelOrigemEmpr, 4) + "." + XLO_CODIGO!cont_cd_Contrato + "." + XLO_CODIGO!titu_cd_Plano + "." + XLO_CODIGO!titu_cd_Parcela + "." + XLO_CODIGO!titu_cd_residuo
                            
                            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
                                "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
                            
                            If Not XLO_QUALIFIN.EOF Then
                                XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
                                    
                                'Exclui a distribuição da nota fiscal
                                Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal = " & XLO_QUALIFIN!nofi_cd_NotaFiscal & ""
                                XLO_QUALIFIN.Close
                                
                                'Exclui os descontos/acréscimos das duplicatas da nota fiscal
                                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATURA, Dinamico
                                Conexao.Execute "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata= " & XLO_QUALIFIN!dupl_cd_Duplicata
                                XLO_QUALIFIN.Close
                                
                                'Exclui as duplicatas da nota fiscal
                                Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura= " & XLL_FATURA
                                'Exclui a nota fiscal
                                Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_fatura= " & XLL_FATURA
                            
                            Else
                                XLO_QUALIFIN.Close
                            End If
                        End If
                        
                        'Exclui o título do empreendimento de origem do QUALICAPI (tabela de Titulos)
                        XLT_SQL = _
                            "DELETE FROM titulos WHERE " & _
                            " empd_cd_Empreendimento ='" & Left(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                            " AND imov_cd_Imovel = '" & Right(LblImovelOrigemEmpr.Caption, 4) & "'" & _
                            " AND (titu_dt_Pagamento is null or titu_dt_Pagamento = '')  " & _
                            " AND (cont_cd_contrato = '" & XLT_CONTRATO & "')" & _
                            " AND empr_cd_empresa =" & PCodEmpresa & _
                            " AND titu_cd_Plano = " & XLO_CODIGO!titu_cd_Plano & _
                            " AND titu_cd_Parcela >= " & XLO_CODIGO!titu_cd_Parcela
                        
                        Conexao.Execute XLT_SQL
                    End If
                End If
                XLO_CODIGO.MoveNext
            Loop
           
          XLO_CODIGO.Close
          Set XLO_CODIGO = Nothing
          
      End If
        
        'Atualiza o status do imóvel da empresa destino
        XLT_SQL = "UPDATE Imoveis SET " & _
            " imov_tx_Status ='V'" & _
            " WHERE " & _
            " empd_cd_Empreendimento ='" & Left(MaskImovel.Text, 4) & "'" & _
            " AND imov_cd_Imovel = '" & Right(MaskImovel.Text, 4) & "'" & _
            " AND empr_cd_empresa =" & XLT_EMPRESADESTINO
        Conexao.Execute XLT_SQL, , adCmdText
        
        'Aumenta a quantidade de imóveis vendidos no novo empreendimento
        XLT_SQL = "UPDATE Empreendimentos SET " & _
          " empd_nr_UnVendidas=empd_nr_UnVendidas +1 " & _
          " WHERE empr_cd_Empresa=" & XLT_EMPRESADESTINO & _
          " AND empd_cd_Empreendimento ='" & Left(MaskImovel.Text, 4) & "'"
        Conexao.Execute XLT_SQL, , adCmdText
        
        Conexao.CommitTrans
        
        GravaLog TRANSFERENCIA_ENTRE_EMPRESAS, "8"

        Adodc1.Recordset.Requery adCmdText
        
        Call CmdRetornarCodigo_Click
        
        PanTransfEmp.Visible = False
        MsgBox "Transferência realizada com sucesso!!!", vbInformation, "Transferência Entre Empresas"
        
    End If
    Exit Sub
  
Erro:
  Conexao.RollbackTrans
  MsgBox Err.Description, vbCritical, "ATENÇÃO"
  Exit Sub
End Sub

Private Sub CmdRetornarTransfEmp_Click()
  PanTransfEmp.Visible = False
  PanBotoes.Enabled = True
  TDBGrid1.Enabled = True
End Sub

Private Sub CmdTransfEmp_Click()
    
  If Not FunVerifica_Permissao(Me.Name, "8") Then
    Exit Sub
  End If
  
  PanBotoes.Enabled = False
  PanTransfEmp.Visible = True
  PanTransfEmp.Left = (TDBGrid1.Width - PanTransfEmp.Width) / 2
  PanTransfEmp.Top = (TDBGrid1.Height - PanTransfEmp.Height) / 2
  TDBGrid1.Enabled = False
  LblCodigoEmpreendimentoOrigem.Caption = Left(Adodc1.Recordset!contrato, 4)
  LblImovelOrigemEmpr.Caption = Left(Adodc1.Recordset!contrato, 9)
  MskCodigo.Text = "    .    "
  MaskEmpree.SetFocus
  
End Sub

Private Sub CmdTransferencia_Click()

  If Not FunVerifica_Permissao(Me.Name, "8") Then
    Exit Sub
  End If
  
  PanBotoes.Enabled = False
  PanAlterarCodigo.Left = (TDBGrid1.Width - PanAlterarCodigo.Width) / 2
  PanAlterarCodigo.Top = (TDBGrid1.Height - PanAlterarCodigo.Height) / 2
  TDBGrid1.Enabled = False
  PanAlterarCodigo.Visible = True
  LblCodigoContrato.Caption = Left(Adodc1.Recordset!contrato, 9)
  MskCodigo.Text = "    .    "
  MskCodigo.SetFocus
  
End Sub

Private Sub CmdCessao_Click()
    
    If Not FunVerifica_Permissao(Me.Name, "6") Then
      Exit Sub
    End If
    
    Dim XLT_SQL As String
    
    PanBotoes.Enabled = False
    PanCessao.Left = (TDBGrid1.Width - PanCessao.Width) / 2
    PanCessao.Top = (TDBGrid1.Height - PanCessao.Height) / 2
    TDBGrid1.Enabled = False
    
    DtpDtDistratoCessao.Value = Date
    LblNomeCliente = Adodc1.Recordset.Fields("Cliente")
    LblDtContrato = Format(Adodc1.Recordset.Fields("cont_dt_RegistroContrato"), "dd/mm/yy")
    
    XLT_SQL = "SELECT Titulo,CASE WHEN titu_dt_Pagamento IS NULL THEN " & _
    "(Titulo + ' - ' + CONVERT(varchar(8), titu_dt_Vencimento, 3) + ' - ' + tipl_tx_descricao + ' - ' + natureza) " & _
    "WHEN titu_dt_Pagamento IS NOT NULL THEN " & _
    "(Titulo + ' - ' + CONVERT(varchar(8), titu_dt_Vencimento, 3) + ' - ' + CONVERT(varchar(8), titu_dt_Pagamento, 3)" & _
    "+ ' - ' + tipl_tx_descricao + ' - ' + natureza) END AS Parcela" & _
    " FROM ConsCapTitulos" & _
    " WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "'" & _
    " AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & Right(Adodc1.Recordset.Fields("Contrato"), 2) & "'" & _
    " ORDER BY Titulo"
    
    subConectarControleDadosNV DatContratoInicial, XLT_SQL, Estatico
    
    If Not DatContratoInicial.Recordset.EOF = True Then
        DatContratoInicial.Recordset.MoveFirst
        CboContratoInicial.BoundText = DatContratoInicial.Recordset.Fields("Titulo")
        
        subConectarControleDadosNV DatContratoFinal, XLT_SQL, Estatico
        DatContratoFinal.Recordset.MoveLast
        CboContratoFinal.BoundText = DatContratoFinal.Recordset.Fields("Titulo")
        
        ChkTitulosPagos.Value = 0
        PanCessao.Visible = True
        PanCessao.Enabled = True
        PanCessao.Tag = "C"
        XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
        DtpDtDistratoCessao.SetFocus
    Else
        MsgBox "Esse contrato não possui nenhum título", vbCritical, "ATENÇÃO !"
        PanBotoes.Enabled = True
        TDBGrid1.Enabled = True
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub CmdConfirmar_Click()
On Error GoTo TrataErro

    Dim XLT_TITULOQUALIFIN As String
    Dim XLL_FATURA As Long
    
    XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
    
    If DtpDtDistrato.Value < Adodc1.Recordset.Fields("cont_dt_RegistroContrato") Then
        MsgBox "A Data do Distrato deve se maior ou igual que a Data do Contrato.", vbCritical, "ATENÇÃO !"
        Exit Sub
    End If
    
    Conexao.BeginTrans
    Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    Conexao.Execute "UPDATE Contratos SET cont_tx_Status='I',cont_dt_Distrato=" & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    
    If ChkLimparImovel.Value = 1 Then
        Conexao.Execute "UPDATE Imoveis SET imov_vl_AVista=0.00, imov_tx_Garagem='' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    End If
    
    '************EXCLUSÃO DOS TÍTULOS DO QUALIFIN****************
    If XGB_PQUALIFIN Then
        XLT_TITULOQUALIFIN = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL + "." + XFT_CODCONTRATO + "%"
        
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
            "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & " and nofi_tx_status <> 'R'", Estatico
        
        While Not XLO_QUALIFIN.EOF
            XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
            
            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal = " & XLO_QUALIFIN!nofi_cd_NotaFiscal & ""
            Conexao.Execute "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata in (SELECT dupl_cd_duplicata FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATURA & ")"
            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura= " & XLL_FATURA
            Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_fatura= " & XLL_FATURA
            
            XLO_QUALIFIN.MoveNext
        Wend
        XLO_QUALIFIN.Close
    
        Set XLO_QUALIFIN = Nothing
    End If
    '***************************************************
    
    Conexao.CommitTrans
    
    GravaLog DISTRATAR_CONTRATO, "5"
      
    subCarregaDadosNV Adodc1, "", "Contrato", _
         "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    Call CmdRetornar_Click
    
    Exit Sub
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
    Exit Sub
    
End Sub

Private Sub CmdConfirmarCessao_Click()
  On Error GoTo TrataErro
  
    Dim XLO_Contrato As New ADODB.Recordset
    Dim XLO_TITULO As New ADODB.Recordset
    Dim XLT_NOVOCODCONTRATO As String
    Dim XLF_TITULOINICIAL As String
    Dim XLF_TITULOFINAL As String
    Dim XLT_SQL As String
    Dim XLT_CODIMOVEL As String, XLT_CODEMPREENDIMENTO As String
    Dim XLT_CODCONTRATO As String, XLT_CODPLANO As String
    Dim XLT_CODPARCELA As String, XLT_CODRESIDUO As String
    
    If DtpDtDistratoCessao.Value < Adodc1.Recordset.Fields("cont_dt_RegistroContrato") Then
        MsgBox "A Data do Distrato deve se maior ou igual que a Data do Contrato.", vbCritical, "ATENÇÃO !"
        Exit Sub
    End If
  
    'Inativa o contrato que será cedido para poder ser criado um contrato ativo
  
    Conexao.BeginTrans
    
    Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    Conexao.Execute "UPDATE Contratos SET cont_tx_Status='I',cont_dt_Distrato=" & FunNuloData(DtpDtDistratoCessao.Value, NomeSgbd) & "  WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    
    Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL + "." + XFT_CODCONTRATO
    XInserir = "C"
    XOrigem = "PanCessao"
    FrmManContratos.Show 1
    XOrigem = ""
      
    'Codigo do novo contrato a ser criado
    XLT_NOVOCODCONTRATO = Chave
    
    'Verifica se o contrato foi realmente criado e com status ATIVO
    SubQOpenRecordset XLO_Contrato, "SELECT * FROM Contratos WHERE " & _
            "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
            "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato ='" & XLT_NOVOCODCONTRATO & "' " & _
            "AND cont_tx_Status='A' ", "Estatico"
            
    If XLO_Contrato.EOF Then
        Conexao.RollbackTrans
        MsgBox "Um contrato ATIVO não foi criado para herdar os títulos.", vbCritical, "ATENÇÃO !"
        XLO_Contrato.Close
        Set XLO_Contrato = Nothing
        Exit Sub
    Else
        Conexao.CommitTrans
        
        XLF_TITULOINICIAL = CboContratoInicial.BoundText
        XLF_TITULOFINAL = CboContratoFinal.BoundText
            
        If ChkTitulosPagos.Value = 0 Then
            SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo >= '" & XLF_TITULOINICIAL & "' AND Titulo <= '" & XLF_TITULOFINAL & "' AND titu_dt_Pagamento is Null AND empr_cd_Empresa = " & PCodEmpresa & " ORDER BY Titulo", Estatico
        Else
            SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo >= '" & XLF_TITULOINICIAL & "' AND Titulo <= '" & XLF_TITULOFINAL & "' AND empr_cd_Empresa=" & PCodEmpresa & " ORDER BY Titulo", Estatico
        End If
                
        Conexao.BeginTrans
        
        While Not XLO_TITULO.EOF
            XLT_CODPLANO = Mid$(XLO_TITULO!titulo, 14, 2)
            XLT_CODPARCELA = Mid$(XLO_TITULO!titulo, 17, 3)
            XLT_CODRESIDUO = Right$(XLO_TITULO!titulo, 2)
                 
            XLT_SQL = "INSERT INTO Titulos (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel,cont_cd_Contrato, " & _
              "titu_cd_Plano,titu_cd_Parcela,titu_cd_Residuo,obse_cd_Observacao, " & _
              "moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, " & _
              "titu_vl_ParcelaIndex1, titu_vl_ParcelaIndex2, titu_dt_Vencimento, tipl_cd_TipoPlano, " & _
              "titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
              "cere_cd_Pcr,plco_cd_Conta,coco_cd_Codigo,titu_dt_Base," & _
              "titu_dt_Prorrogacao,titu_tx_NaturezaPlano,titu_vl_ValorPago, " & _
              "titu_dt_Pagamento,titu_vl_Multa,titu_vl_Juros,titu_vl_Seguro, " & _
              "titu_vl_Outros,titu_vl_Desconto,titu_vl_Desagio,titu_tx_Financia," & _
              "titu_tx_IndicePagamento,titu_dt_Deposito,titu_dt_Repasse, " & _
              "titu_tx_LocalPagto,titu_tx_DocPagto,titu_tx_TipoPag, " & _
              "titu_tx_CorrigeParcela, titu_tx_PreDatado, titu_dt_Criacao, " & _
              "titu_dt_BasePagto,moed_cd_Moeda3, titu_vl_corrigido1, titu_vl_corrigido2) "
            
            XLT_SQL = XLT_SQL + "VALUES (" & PCodEmpresa & ",'" & XFT_CODEMPREENDIMENTO & "','" & XFT_CODIMOVEL & "','" & XLT_NOVOCODCONTRATO & "'," & _
              "'" & XLT_CODPLANO & "','" & XLT_CODPARCELA & "','" & XLT_CODRESIDUO & "'," & FunNuloBancoVal(XLO_TITULO!obse_cd_Observacao) & ", " & _
              "" & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda1) & "," & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda2) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Parcela)) & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ParcelaIndex1)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ParcelaIndex2)) & "," & FunNuloData(XLO_TITULO!titu_dt_Vencimento, NomeSgbd) & "," & XLO_TITULO!tipl_cd_TipoPlano & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosFin)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercMultaMora)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosMora)) & ", " & _
              "" & FunNuloBanco(XLO_TITULO!cere_cd_Pcr) & "," & FunNuloBanco(XLO_TITULO!plco_cd_Conta) & "," & FunNuloBancoVal(XLO_TITULO!coco_cd_codigo) & ", " & FunNuloData(XLO_TITULO!titu_dt_Base, NomeSgbd) & "," & _
              "" & FunNuloData(XLO_TITULO!titu_dt_Prorrogacao, NomeSgbd) & ",'" & XLO_TITULO!titu_tx_NaturezaPlano & "'," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ValorPago)) & "," & FunNuloData(XLO_TITULO!titu_dt_Pagamento, NomeSgbd) & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Multa)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Juros)) & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Seguro)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Outros)) & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Desconto)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Desagio)) & "," & _
              "'" & XLO_TITULO!titu_tx_Financia & "','" & XLO_TITULO!titu_tx_IndicePagamento & "'," & FunNuloData(XLO_TITULO!titu_dt_Deposito, NomeSgbd) & "," & _
              "" & FunNuloData(XLO_TITULO!titu_dt_Repasse, NomeSgbd) & ",'" & XLO_TITULO!titu_tx_LocalPagto & "','" & XLO_TITULO!titu_tx_DocPagto & "'," & _
              "'" & XLO_TITULO!titu_tx_TipoPag & "','" & XLO_TITULO!titu_tx_CorrigeParcela & "','" & XLO_TITULO!titu_tx_PreDatado & "'," & FunNuloData(XLO_TITULO!titu_dt_Criacao, NomeSgbd) & "," & _
              "" & FunNuloData(XLO_TITULO!titu_dt_BasePagto, NomeSgbd) & "," & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda3) & "," & FunTrataFloat(FunNuloVal(XLO_TITULO!titu_vl_Corrigido1)) & "," & FunTrataFloat(FunNuloVal(XLO_TITULO!titu_vl_Corrigido2)) & ") "
              
            Conexao.Execute (XLT_SQL)
            
            XLO_TITULO.MoveNext
        Wend
        
        Conexao.Execute ("UPDATE Contratos SET cont_tx_Status='C' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'")
        
        If ChkLimparImovel.Value = 1 Then
            Conexao.Execute "UPDATE Imoveis SET imov_vl_AVista=0.00, imov_tx_Garagem='' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        End If
        
        GravaLog CESSAO_CONTRATO, "6"
    End If
    
    Conexao.CommitTrans
    
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
    
    subCarregaDadosNV Adodc1, "", "Contrato", _
         "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    
    Call CmdRetornarCessao_Click
    XOrigem = ""
    
    Exit Sub
  
TrataErro:
  Conexao.RollbackTrans
  funTrataErros (ComMensagem)

End Sub

Private Sub CmdConfirmarCodigo_Click()
  
    If Len(Trim(MskCodigo.Text)) < 9 Then
        MsgBox "O novo código deve possuir 9 dígitos!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    Dim XLO_CODIGO As New ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLT_CONTRATO As String
    
    Conexao.BeginTrans
  
    XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
    
    'Verifica se o novo imóvel já existe
    XLT_SQL = "SELECT empd_cd_Empreendimento FROM Imoveis WHERE " & _
            " empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'" & _
            " AND imov_cd_Imovel = '" & Right(MskCodigo.Text, 4) & "'" & _
            " AND empr_cd_empresa =" & PCodEmpresa
    SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
    
    If XLO_CODIGO.EOF Then
        Conexao.RollbackTrans
        MsgBox "O imóvel do novo código não existe!", vbCritical, "ATENÇÃO"
        Exit Sub
    Else
        XLO_CODIGO.Close
        
        'Verifica se existe um contrato ativo no novo imóvel
        XLT_SQL = "SELECT cont_cd_contrato FROM Contratos WHERE " & _
              " empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'" & _
              " AND imov_cd_Imovel = '" & Right(MskCodigo.Text, 4) & "'" & _
              " AND empr_cd_empresa =" & PCodEmpresa & _
              " AND cont_tx_Status = 'A'"
        SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
        
        If Not XLO_CODIGO.EOF Then
            Conexao.RollbackTrans
            MsgBox "Já existe um contrato ATIVO para este novo imóvel!", vbCritical, "ATENÇÃO"
            Exit Sub
        Else
            XLO_CODIGO.Close
            
            'Pega o codigo do último contrato do novo imóvel
            XLT_SQL = "SELECT isnull(max(cont_cd_contrato),0) as Codigo  FROM Contratos WHERE " & _
              " empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'" & _
              " AND imov_cd_Imovel = '" & Right(MskCodigo.Text, 4) & "'" & _
              " AND empr_cd_empresa =" & PCodEmpresa
            SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
            
            XLT_CONTRATO = Format(CInt(XLO_CODIGO!Codigo) + 1, "00")
            XFT_CODNOVOCONTRATO = XLT_CONTRATO
                
            XLO_CODIGO.Close
        End If
    End If
  
    On Error GoTo Erro
    
    'Atualiza a tabela de contratos e a de título por update cascade
    XLT_SQL = "UPDATE Contratos SET " & _
        " empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'" & _
        " ,imov_cd_Imovel = '" & Right(MskCodigo.Text, 4) & "'" & _
        " ,cont_cd_Contrato = '" & XLT_CONTRATO & "'" & _
        " WHERE " & _
        " empd_cd_Empreendimento ='" & XFT_CODEMPREENDIMENTO & "'" & _
        " AND imov_cd_Imovel = '" & XFT_CODIMOVEL & "'" & _
        " AND cont_cd_Contrato = '" & XFT_CODCONTRATO & "'" & _
        " AND empr_cd_empresa =" & PCodEmpresa
    Conexao.Execute XLT_SQL, , adCmdText
        
    'Atualiza o status do imóvel original
    XLT_SQL = "UPDATE Imoveis SET " & _
        " imov_tx_Status ='D'" & _
        " WHERE " & _
        " empd_cd_Empreendimento ='" & XFT_CODEMPREENDIMENTO & "'" & _
        " AND imov_cd_Imovel = '" & XFT_CODIMOVEL & "'" & _
        " AND empr_cd_empresa =" & PCodEmpresa
    Conexao.Execute XLT_SQL, , adCmdText
    
    'Atualiza o status do novo imóvel
    XLT_SQL = "UPDATE Imoveis SET " & _
        " imov_tx_Status ='V'" & _
        " WHERE " & _
        " empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'" & _
        " AND imov_cd_Imovel = '" & Right(MskCodigo.Text, 4) & "'" & _
        " AND empr_cd_empresa =" & PCodEmpresa
    Conexao.Execute XLT_SQL, , adCmdText
    
    'Diminui a quantidade de imóveis vendidos no empreendimento original
    XLT_SQL = "UPDATE Empreendimentos SET " & _
      " empd_nr_UnVendidas=empd_nr_UnVendidas -1 " & _
      " WHERE empr_cd_Empresa=" & PCodEmpresa & _
      " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
    Conexao.Execute XLT_SQL, , adCmdText
    
    'Aumenta a quantidade de imóveis vendidos no novo empreendimento
    XLT_SQL = "UPDATE Empreendimentos SET " & _
      " empd_nr_UnVendidas=empd_nr_UnVendidas +1 " & _
      " WHERE empr_cd_Empresa=" & PCodEmpresa & _
      " AND empd_cd_Empreendimento ='" & Left(MskCodigo.Text, 4) & "'"
    Conexao.Execute XLT_SQL, , adCmdText
    
    'Atualiza os código dos títulos no financeiro
    XLT_SQL = "UPDATE NotasFiscais SET " & _
      " nofi_nr_TituloCapi='" & MskCodigo.Text & "." & XLT_CONTRATO & "' + right(nofi_nr_TituloCapi,10) " & _
      " ,nofi_tx_Historico= left(nofi_tx_Historico,6)+ '" & MskCodigo.Text & "." & XLT_CONTRATO & "' + right(nofi_tx_Historico,len(nofi_tx_Historico)-18)" & _
      " WHERE " & _
      " left(nofi_nr_TituloCapi,12)= '" & Adodc1.Recordset!contrato & "'"
    Conexao.Execute XLT_SQL, , adCmdText
    
    Conexao.CommitTrans
    
    GravaLog TRANSFERENCIA, "8"
  
    Adodc1.Recordset.Requery adCmdText
    Call CmdRetornarCodigo_Click
    
    Exit Sub
  
Erro:
    Conexao.RollbackTrans
    MsgBox Err.Description, vbCritical, "ATENÇÃO"
    Exit Sub
  
End Sub

Private Sub CmdRetornar_Click()
  
  PanDistratar.Visible = False
  PanDistratar.Enabled = False
  PanBotoes.Enabled = True
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  
End Sub

Private Sub CmdDistratar_Click()
    On Error GoTo TrataErro
        
    If Not FunVerifica_Permissao(Me.Name, "5") Then
        Exit Sub
    End If
    
    Dim XLO_Contrato As New ADODB.Recordset
    
    'Verifica se vai distratar ou ativar
    If CmdDistratar.Caption = "&Distratar" Then
        PanBotoes.Enabled = False
        PanDistratar.Left = (TDBGrid1.Width - PanDistratar.Width) / 2
        PanDistratar.Top = (TDBGrid1.Height - PanDistratar.Height) / 2
        TDBGrid1.Enabled = False
        DtpDtDistrato.Value = Date
        PanDistratar.Visible = True
        PanDistratar.Enabled = True
        PanDistratar.Tag = "D"
        DtpDtDistrato.SetFocus
    
    'Ativa um contrato
    Else
        If MsgBox("Confirma Ativação do Contrato ?", 273, "ATENÇÂO") = vbOK Then
            XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
             
            SubQOpenRecordset XLO_Contrato, "SELECT cont_tx_status,cont_cd_Contrato FROM Contratos WHERE " & _
                 "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
                 "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND (cont_tx_Status ='A' OR cont_tx_Status ='R')", "Estatico"
             
            If Not (XLO_Contrato.BOF And XLO_Contrato.EOF) Then
                'Verifica se já não existe um contrato ATIVO
                If XLO_Contrato!cont_tx_Status = "A" Then
                    MsgBox "Esse imóvel já está vendido.", vbCritical, "ATENÇÃO !"
                    Exit Sub
                End If
                  
                'Verifica se já não existe um contrato RESERVADO quando o  contrato que estiver sendo Ativado for INATIVO
                If Adodc1.Recordset.Fields("Status") = "INATIVO" Then
                    If XLO_Contrato!cont_tx_Status = "R" Then
                        MsgBox "Esse imóvel já está reservado.", vbCritical, "ATENÇÃO !"
                        Exit Sub
                    End If
                'Altera de RESERVADO para ATIVO, chama o formulário de contratos
                Else
                    Chave = Adodc1.Recordset.Fields("Contrato")
                    XInserir = "A"
                    FrmManContratos.Show 1
                    subCarregaDadosNV Adodc1, "", "Contrato", _
                       "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
                End If
            'Altera de INATIVO para ATIVO
            Else
                Conexao.BeginTrans
                Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas +1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
                Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='V' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
                Conexao.Execute "UPDATE Contratos SET cont_tx_Status='A',cont_dt_Distrato=Null WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
                Conexao.CommitTrans
                
                GravaLog ATIVAR_CONTRATO, "5"
                
                subCarregaDadosNV Adodc1, "", "Contrato", _
                     "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
            End If
             
            XLO_Contrato.Close
            Set XLO_Contrato = Nothing
        End If
    End If
    
    Exit Sub
  
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)

End Sub

Private Sub CmdExcluir_Click()
    
  If Not FunVerifica_Permissao(Me.Name, "4") Then
    Exit Sub
  End If
  
  Dim XLI_HOUVEERRO As Integer
  Dim XLT_STATUS As String 'Pega o Status do contrato
  Dim XLO_Contrato As New ADODB.Recordset
  
  
  XLI_HOUVEERRO = False
  
  subDesabilitaBotoes
     
  SubQOpenRecordset XLO_Contrato, "SELECT * FROM Titulos WHERE" & _
    " empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
    " AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'", Estatico
  
  If Not XLO_Contrato.EOF Then
        While Not XLO_Contrato.EOF
          If Not IsNull(XLO_Contrato!titu_dt_Pagamento) Then
            MsgBox "Esse contrato não poderá ser excluído pois contém Títulos pagos .", vbCritical, "ATENÇÃO"
            XLO_Contrato.Close
            Set XLO_Contrato = Nothing
            subHabilitaBotoes
            Exit Sub
          End If
          XLO_Contrato.MoveNext
        Wend
        
        If MsgBox("Confirma Remoção do Contrato e seus Títulos?", 273, "CUIDADO") = vbOK Then
             Conexao.BeginTrans
             
             Conexao.Execute ("DELETE FROM Titulos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'")
             Conexao.Execute "DELETE FROM Contratos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'"
               
            'Excluindo os dados no financeiro
            'Verifica se tem o financeiro e se exporta compromissos
            If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then
                Call SubExcluiFinanceiro
            End If
             
               
            XLT_STATUS = Adodc1.Recordset.Fields("cont_tx_Status")
            'Se o contrato deletado tiver o status ativo,
            'o status do imovel se torna disponível
            If XLT_STATUS = "A" Or XLT_STATUS = "R" Then
              Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
              Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            End If
            
            Conexao.CommitTrans
            
            GravaLog EXCLUIR, "4"
          
        End If
  Else
        XLT_STATUS = Adodc1.Recordset.Fields("cont_tx_Status")
         
        If MsgBox("Confirma Remoção do Contrato e seus Títulos?", 273, "CUIDADO") = vbOK Then
         
            Conexao.BeginTrans
            Conexao.Execute "DELETE FROM Contratos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'"
            
            'Se o contrato deletado tiver o status ativo,
            'o status do imovel se torna disponível
            If XLT_STATUS = "A" Or XLT_STATUS = "R" Then
              Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
              Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            End If
            
            Conexao.CommitTrans
            
            GravaLog EXCLUIR, "4"

        End If
  End If
      
  If XLI_HOUVEERRO = False Then
     subCarregaDadosNV Adodc1, "", "Contrato", _
     "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
     If ContFil <> 0 Then
        subContaReg Adodc1
     End If
  End If
  
  subHabilitaBotoes
  
  If FunTabelaVazia(Adodc1) = True Then
     CmdAlterar.Enabled = False
     CmdExcluir.Enabled = False
     CmdDistratar.Enabled = False
     CmdImprimir.Enabled = False
     CmdTitulo.Enabled = False
     CmdMapaFechamento.Enabled = False
     CmdCessao.Enabled = False
     CmdAssistencia.Enabled = False
  End If
  
  XLO_Contrato.Close
  Set XLO_Contrato = Nothing
  
  TDBGrid1.SetFocus
  
  Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Then
      MsgBox "Esse contrato não pode ser excluído pois já possui títulos relacionados!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
End Sub

Private Sub CmdImprimir_Click()
  Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Contratos")
End Sub

Private Sub CmdInserir_Click()
    If Not FunVerifica_Permissao(Me.Name, "2") Then
        Exit Sub
    End If
    
    Dim XLI_CONT As Integer
    
    For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
      TDBGrid1.row = XLI_CONT
      If TDBGrid1.Columns(1).Text = "ATIVO" Then
        MsgBox "Esse imóvel já possui um contrato ATIVO.", vbCritical, "ATENÇÃO!"
        Exit Sub
      ElseIf TDBGrid1.Columns(1).Text = "RESERVADO" Then
        MsgBox "Esse imóvel já possui um contrato RESERVADO, ative ou exclua a reserva.", vbCritical, "ATENÇÃO!"
        Exit Sub
      End If
    Next

    Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL
    XInserir = "I"
    FrmManContratos.Show 1
    subCarregaDadosNV Adodc1, "", "Contrato", _
        "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    If Not Adodc1.Recordset.EOF Then
      CmdAlterar.Enabled = True
      CmdExcluir.Enabled = True
      CmdImprimir.Enabled = True
    End If
End Sub

Private Sub CmdMapaFechamento_Click()
  If Not FunVerifica_Permissao(Me.Name, "7") Then
    Exit Sub
  End If
    
  Chave = Adodc1.Recordset.Fields("Contrato")
  XInserir = "I"
  FrmManMapaPlanos.Show 1
End Sub

Private Sub CmdRetornarCessao_Click()
  PanCessao.Visible = False
  PanCessao.Enabled = False
  PanBotoes.Enabled = True
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
End Sub

Private Sub CmdRetornarCodigo_Click()
  PanAlterarCodigo.Visible = False
  PanBotoes.Enabled = True
  TDBGrid1.Enabled = True
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub CmdTitulo_Click()
  If Not FunVerifica_Permissao("TelaFinTitulos", "1") Then
      Exit Sub
  End If
  
  XOrigem = "TelaManContratos"
  Chave = Adodc1.Recordset.Fields("Contrato")
  TelaFinTitulos.Show
  
  If FunLocalizaJanelasAtivas("TelaFinTitulos") <> -1 Then
      TelaFinTitulos.SetFocus
  End If
End Sub



Private Sub Form_Activate()
  Set Formulario = TelaManContratos
  XOrigem = "TelaManContratos"
End Sub

Private Sub Form_Load()
    Dim XLO_IMOVEL As New ADODB.Recordset
    Dim XLB_ACHOU As Boolean
    
    Set Formulario = TelaManContratos
    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Right$(Chave, 4)
 
    subCarregaDadosNV Adodc1, "", "Contrato", _
                            "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    
    subConectarControleDadosNV DatEmpresa, "SELECT * FROM Empresas ORDER BY Empr_tx_Fantasia", Estatico
    SubQOpenRecordset XLO_IMOVEL, "SELECT empd_tx_Nome FROM Empreendimentos WHERE empd_cd_Empreendimento ='" & XFT_CODEMPREENDIMENTO & "' AND empr_cd_empresa = " & PCodEmpresa, Estatico
    
    If Not XLO_IMOVEL.EOF Then
        LblNomeEmpreendimento.Caption = XLO_IMOVEL!empd_tx_Nome
        LblCodEmpreendimento.Caption = XFT_CODEMPREENDIMENTO
        LblCodImovel.Caption = XFT_CODIMOVEL
    End If
    
    XLO_IMOVEL.Close
    Set XLO_IMOVEL = Nothing
    
    If Adodc1.Recordset.EOF Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdCessao.Enabled = False
        CmdAssistencia.Enabled = False
        CmdDistratar.Enabled = False
        CmdTitulo.Enabled = False
        CmdMapaFechamento.Enabled = False
        CmdImprimir.Enabled = False
        CmdTransfEmp.Enabled = False
        CmdTransferencia.Enabled = False
    Else
        XLB_ACHOU = False
        While (Not Adodc1.Recordset.EOF) And (XLB_ACHOU = False)
            If Adodc1.Recordset("cont_tx_Status") = "A" Then
                XLB_ACHOU = True
            Else
                Adodc1.Recordset.MoveNext
            End If
        Wend
        If XLB_ACHOU = False Then
            Adodc1.Recordset.MoveFirst
            'Habilita o botão ativar, já que não foi localizado nenhum contrato ativo
            CmdDistratar.Enabled = True
            CmdDistratar.Caption = "&Ativar"
        End If
    End If
    
    Call subAjustaTelaContrato
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL
   Set TelaManContratos = Nothing
End Sub

Private Sub Mskconta_LostFocus()
    MskConta.PromptInclude = False
    If MskConta.Text <> "" Then
      If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
          MskConta.Text = Format(MskConta.Text, XGT_MSKCODREDUZIDO)
      End If
      
      MskConta.PromptInclude = True
      CboPlanoContas.BoundText = MskConta.Text
      
      If CboPlanoContas.Text = "" Then
        MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
        MskConta.SetFocus
        If MskConta.Enabled = True Then MskConta.SetFocus
      End If
    End If
    MskConta.PromptInclude = True
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

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.Fields("Status") = "ATIVO" Then
            CmdDistratar.Caption = "&Distratar"
            CmdDistratar.Enabled = True
            CmdCessao.Enabled = True
            CmdAssistencia.Enabled = True
            CmdAlterar.Enabled = True
            CmdMapaFechamento.Enabled = True
            CmdTitulo.Enabled = True
            CmdTransferencia.Enabled = True
        ElseIf Adodc1.Recordset.Fields("Status") = "INATIVO" Then
            CmdDistratar.Caption = "&Ativar"
            CmdDistratar.Enabled = True
            CmdCessao.Enabled = -False
            CmdAssistencia.Enabled = -False
            CmdAlterar.Enabled = True
            CmdMapaFechamento.Enabled = True
            CmdTitulo.Enabled = True
            CmdTransferencia.Enabled = False
        ElseIf Adodc1.Recordset.Fields("Status") = "RESERVADO" Then
            CmdDistratar.Caption = "&Ativar"
            CmdDistratar.Enabled = True
            CmdCessao.Enabled = False
            CmdAssistencia.Enabled = False
            CmdAlterar.Enabled = True
            CmdMapaFechamento.Enabled = False
            CmdTitulo.Enabled = False
            CmdTransferencia.Enabled = False
        ElseIf Adodc1.Recordset.Fields("Status") = "CEDIDO" Then
            If CmdDistratar.Caption <> "&Ativar" Then
                CmdDistratar.Enabled = False
            Else
                CmdDistratar.Enabled = True
            End If
            CmdCessao.Enabled = False
            CmdAssistencia.Enabled = False
            CmdAlterar.Enabled = False
            CmdMapaFechamento.Enabled = False
            CmdTitulo.Enabled = True
            CmdTransferencia.Enabled = False
        End If
    End If
End Sub

Sub subGravaFinanceiroTransf(XLO_TITULO As Object, XLT_STATUS As String, XLI_CODCLIENTE As Integer, XLT_TIPOPLANO As String, XLT_NATUREZAPLANO As String, XLT_EMPREEENDIMENTO As String, XLT_IMOVEL As String)
    'Esta função grava os lançamentos compromissados e/ou realizados no financeiro

    '********PARÂMETRO*********
    'XLO_TITULO 'Recordset com a tabela dos titulos
    'XLT_STATUS 'Indicar se está pagando, cancelando, ou cadastrando um título
                '(Pagar, Cancelar, Compromissar, Alterar)
    'XLI_CODCLIENTE 'Código do cliente
    'XLT_TIPOPLANO  'Tipo Plano
    'XLT_NATUREZAPLANO  'Natureza plano
    '**************************
    Dim XLT_SQL As String           'Prepara Select
    Dim XLT_NUMDOC As String        'Nº do doc. da NF (gerado pelo sistema)
    Dim XLT_STATUSGRAVAR As String  'Sttaus da Nota e da parcela para gravação no banco
    Dim XLL_FATURA As Long          'Nº da Fatura que será gravado em NF e Duplicatas
    Dim XLI_TIPODOC As Integer      'Código do tipo de documento = Nota Promissória
    Dim XLO_REGISTRO As ADODB.Recordset 'Recordset para os acessos às tabelas do Qualifin
    Dim XLT_DOCREC As String        'Nº do doc. recebimento. Foi gerado pelo sistema posi no capi vem tudo junto
    Dim XLT_TITULO As String        'Código completo do título
    Dim XLF_VALOR As Currency       'Valor A Receber(compromisso) Recebido (pago)
    Dim XLL_CODDUPL As Long         'Código da duplicat no financeiro
    Dim XLI_CODJUROS As Integer     'Código do Acrécimo Juros no Financeiro
    Dim XLI_CODMULTA As Integer     'Código do Acrécimo Multa no Financeiro
    Dim XLI_CODDESAGIO As Integer   'Código do Desconto Deságio no Financeiro
    Dim XLI_CODDESCONTO As Integer  'Código do Desconto no Financeiro
    Dim XLI_CODSEGURO As Integer    'Código do Acréscimo Seguro no Financeiro
    Dim XLI_CODOUTROS As Integer    'Código do Acréscimo Outros no Financeiro
    Dim XLD_REGCONTRATO As Date     'Data do Reg. Contrato Para ser exp. como dt. emissão p/ Qualifin qdo. pago
    Dim XLD_VENDACONTRATO As Date   'Dt. Venda do Imovel P/ ser exp. como dt. compet. e emissão p/ Qualifin qdo. não pago
    Dim XLT_TIPOPAGTO As String     'Tipo de Pagto/rec (se cheque, dinheiro...)
    Dim XLI_REGATUALIZADOS As Integer 'Indica quantos registros foram atualizados na tabela de Distribuição
    Dim XLO_TEMDESCONTO As ADODB.Recordset
    
    
    
    'Obter código do tipo de doc cuja sigla = NP (Nota Promissória)
    SubQOpenRecordset XLO_REGISTRO, "SELECT tido_cd_TipoDoc FROM TipoDocumentos WHERE tido_tx_Sigla= 'NP'", Estatico
    XLI_TIPODOC = XLO_REGISTRO!tido_cd_TipoDoc
    XLO_REGISTRO.Close
    
    'Obter Dt de registro do contrato e dt. venda do imovel
    SubQOpenRecordset XLO_REGISTRO, "SELECT cont_dt_RegistroContrato, cont_dt_Venda FROM Contratos WITH (ROWLOCK)  " & _
                " WHERE empd_cd_Empreendimento= '" & XLT_EMPREEENDIMENTO & "' AND " & _
                " imov_cd_Imovel='" & XLT_IMOVEL & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
                " empr_cd_Empresa = " & PCodEmpresa & "", Estatico
                
    XLD_REGCONTRATO = XLO_REGISTRO!cont_dt_RegistroContrato
    XLD_VENDACONTRATO = XLO_REGISTRO!cont_dt_Venda
    XLO_REGISTRO.Close
    
    XLT_TITULO = XLO_TITULO!empd_cd_Empreendimento + "." + XLO_TITULO!imov_cd_Imovel + "." + XLO_TITULO!cont_cd_Contrato + "." + XLO_TITULO!titu_cd_Plano + "." + XLO_TITULO!titu_cd_Parcela + "." + XLO_TITULO!titu_cd_residuo
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        XLO_REGISTRO.AddNew
        XLL_FATURA = FunPegaGuardaUltimo("F")
        XLT_NUMDOC = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_REGISTRO!nofi_nr_documento = XLT_NUMDOC
        XLO_REGISTRO!nofi_nr_fatura = XLL_FATURA
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM NotasFiscais WHERE nofi_nr_TituloCapi= '" & XLT_TITULO & "' AND empr_cd_Empresa = " & PCodEmpresa & "", Dinamico
        If Not XLO_REGISTRO.EOF Then
            XLT_NUMDOC = XLO_REGISTRO!nofi_nr_documento
            XLL_FATURA = XLO_REGISTRO!nofi_nr_fatura
        End If
    End If
    
    If XLO_REGISTRO.EOF Then 'Se não tiver sido compromissado no Qualifin
        XLO_REGISTRO.AddNew
        XLL_FATURA = FunPegaGuardaUltimo("F")
        XLT_NUMDOC = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_REGISTRO!nofi_nr_documento = XLT_NUMDOC
        XLO_REGISTRO!nofi_nr_fatura = XLL_FATURA
    End If
    
    If XLT_STATUS = "COMPROMISSAR" Or XLT_STATUS = "ALTERAR" Then
        XLT_STATUSGRAVAR = "C"
        XLF_VALOR = XLO_TITULO!titu_vl_Parcela
        XLO_REGISTRO!nofi_dt_entrada = Format$(XLD_VENDACONTRATO, "DD/MM/YY")
        XLO_REGISTRO!nofi_dt_emissao = Format$(XLD_VENDACONTRATO, "DD/MM/YY")
     ElseIf XLT_STATUS = "PAGAR" Then
        XLT_STATUSGRAVAR = "R"
        'XLF_VALOR = XLO_TITULO!titu_vl_ValorPago
        XLF_VALOR = Format(XLO_TITULO!titu_vl_ValorPago - FunNuloVal(XLO_TITULO!titu_vl_Seguro) - FunNuloVal(XLO_TITULO!titu_vl_Juros) - FunNuloVal(XLO_TITULO!titu_vl_Multa) - FunNuloVal(XLO_TITULO!titu_vl_Outros) + FunNuloVal(XLO_TITULO!titu_vl_Desagio) + FunNuloVal(XLO_TITULO!titu_vl_Desconto), "standard")
        XLO_REGISTRO!nofi_dt_entrada = Format$(XLO_TITULO!titu_dt_Pagamento, "DD/MM/YY")
        XLO_REGISTRO!nofi_dt_emissao = Format$(XLD_REGCONTRATO, "DD/MM/YY")
    End If
    
    XLO_REGISTRO!nofi_tx_tipomov = "E"   'Entrada
    XLO_REGISTRO!nofi_tx_status = XLT_STATUSGRAVAR
    XLO_REGISTRO!focl_cd_FornCli = XLI_CODCLIENTE
    XLO_REGISTRO!nofi_tx_serie = "U" 'Única
    XLO_REGISTRO!tido_cd_TipoDoc = XLI_TIPODOC
    XLO_REGISTRO!nofi_tx_historico = Mid("PARC. " + XLT_TITULO + " - " + XLT_TIPOPLANO + " - " + XLT_NATUREZAPLANO, 1, 60)
    XLO_REGISTRO!nofi_vl_valor = XLF_VALOR
    XLO_REGISTRO!nofi_tx_distribuir = "N" 'Sempre 1 conta para 1 CCusto
    XLO_REGISTRO!nofi_tx_tipofatura = "1" 'Uma nota para uma duplicata
    XLO_REGISTRO!nofi_nr_TituloCapi = XLT_TITULO 'Código do título
    'XLO_REGISTRO!empr_cd_Empresa = PCodEmpresa
    XLO_REGISTRO!empr_cd_Empresa = XLO_TITULO!empr_cd_Empresa

    XLO_REGISTRO.Update

    Chave = XLO_REGISTRO!nofi_cd_NotaFiscal
    XLL_FATURA = XLO_REGISTRO!nofi_nr_fatura
    XLO_REGISTRO.Close
        
    '*****************************TITULOS E NOTAS FISCAL*******************************
    'Gravar a nofi_cd_notafiscal na tabela de títulos
'    Conexao.Execute "UPDATE Titulos SET nofi_cd_notafiscal = " & Chave & " WHERE empd_cd_Empreendimento= '" & XLO_TITULO!empd_cd_Empreendimento & "' AND " & _
'                " imov_cd_Imovel='" & XLO_TITULO!imov_cd_Imovel & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
'                " titu_cd_plano = '" & XLO_TITULO!titu_cd_Plano & "' AND titu_cd_Parcela = '" & XLO_TITULO!titu_cd_Parcela & "' AND " & _
'                " titu_cd_residuo = '" & XLO_TITULO!titu_cd_residuo & "' AND empr_cd_Empresa = " & PCodEmpresa & ""
'
     Conexao.Execute "UPDATE Titulos SET nofi_cd_notafiscal = " & Chave & " WHERE empd_cd_Empreendimento= '" & XLO_TITULO!empd_cd_Empreendimento & "' AND " & _
                " imov_cd_Imovel='" & XLO_TITULO!imov_cd_Imovel & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
                " titu_cd_plano = '" & XLO_TITULO!titu_cd_Plano & "' AND titu_cd_Parcela = '" & XLO_TITULO!titu_cd_Parcela & "' AND " & _
                " titu_cd_residuo = '" & XLO_TITULO!titu_cd_residuo & "' AND empr_cd_Empresa = " & XLO_TITULO!empr_cd_Empresa & ""
     
    '*****************************DUPLICATAS*******************************
    'dupl_tx_debcred = (C - Crédito)
    'dupl_nr_duplicata=(sempre parcela U - Única)
    'DocPagrec 'Não está separado. Gera automat.
    If XLT_STATUS = "PAGAR" Then
        XLT_DOCREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    End If
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM Duplicatas WHERE dupl_cd_Duplicata = -1", Dinamico
        XLO_REGISTRO.AddNew
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & XLL_FATURA, Dinamico
    End If
        
    If XLO_REGISTRO.EOF Then 'Se não foi cadastrado no Qualifin
        XLO_REGISTRO.AddNew
    End If
    XLO_REGISTRO!dupl_nr_duplicata = "U"
    XLO_REGISTRO!dupl_dt_vencimento = XLO_TITULO!titu_dt_Prorrogacao
    XLO_REGISTRO!dupl_vl_valor = XLF_VALOR
    XLO_REGISTRO!dupl_tx_status = XLT_STATUSGRAVAR
    If XLO_TITULO!titu_tx_LocalPagto = "3" Then 'Dinheiro
        XLO_REGISTRO!dupl_tx_tipopagrec = "6" 'Dinheiro
        XLT_TIPOPAGTO = "6"
    Else
        XLO_REGISTRO!dupl_tx_tipopagrec = "5" 'Crédito em Conta
        XLT_TIPOPAGTO = "5"
    End If
        
    
    XLO_REGISTRO!coco_cd_codigo = XLO_TITULO!coco_cd_codigo
    XLO_REGISTRO!dupl_tx_debcred = "C"
    XLO_REGISTRO!dupl_nr_fatura = XLL_FATURA
    If (XLT_STATUS = "COMPROMISSAR" Or XLT_STATUS = "ALTERAR") Then
        XLO_REGISTRO!dupl_dt_pagrec = Null
    Else
        XLO_REGISTRO!dupl_dt_pagrec = XLO_TITULO!titu_dt_Deposito
    End If
    If (XLT_DOCREC <> "" And XLT_STATUS = "PAGAR") Then
        XLO_REGISTRO!dupl_tx_docpagrec = XLT_DOCREC
    End If
    
    
    XLO_REGISTRO.Update
    
    XLL_CODDUPL = XLO_REGISTRO!dupl_cd_Duplicata
    
    XLO_REGISTRO.Close
            
    '*********************************GRAVA OS DESCONTOS*************************************
    If XLO_TITULO!titu_vl_Multa <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODMULTA, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODMULTA & "," & FunTrataFloat(XLO_TITULO!titu_vl_Multa) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Multa) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODMULTA
        End If
        XLO_TEMDESCONTO.Close
    End If
    
    If XLO_TITULO!titu_vl_Juros <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODJUROS, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODJUROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Juros) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Juros) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODJUROS
        End If
        XLO_TEMDESCONTO.Close
    End If
    
    If XLO_TITULO!titu_vl_Desagio <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODDESAGIO, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODDESAGIO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desagio) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Desagio) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODDESAGIO
        End If
        XLO_TEMDESCONTO.Close
    End If

    If XLO_TITULO!titu_vl_Desconto <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODDESCONTO, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODDESCONTO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desconto) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Desconto) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODDESCONTO
        End If
        XLO_TEMDESCONTO.Close
    End If

    If XLO_TITULO!titu_vl_Seguro <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODSEGURO, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODSEGURO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Seguro) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Seguro) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODSEGURO
        End If
        XLO_TEMDESCONTO.Close
    End If
    
    If XLO_TITULO!titu_vl_Outros <> 0 And XLT_STATUS = "PAGAR" Then
        SubQOpenRecordset XLO_TEMDESCONTO, "SELECT dupl_cd_duplicata, deac_cd_descacresc FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODSEGURO, Estatico
        If XLO_TEMDESCONTO.EOF Then
            Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODOUTROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Outros) & ")"
        Else
            Conexao.Execute "UPDATE DuplDescAcresc SET dude_vl_descacresc = " & FunTrataFloat(XLO_TITULO!titu_vl_Outros) & " WHERE dupl_cd_duplicata = " & XLL_CODDUPL & " AND deac_cd_descacresc = " & XGI_CODOUTROS
        End If
        XLO_TEMDESCONTO.Close
    End If
    
    '*****************************DISTRIBUIÇÃO*******************************
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        XLT_SQL = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta, " & _
                        "dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal)" & _
                  " VALUES ('" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "'" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "" & FunTrataFloat(XLF_VALOR) & "," & _
                        "" & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "," & Chave & ")"
        Conexao.Execute XLT_SQL
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        
        XLI_REGATUALIZADOS = 0
        XLT_SQL = "UPDATE Distribuicao SET cere_cd_pcr='" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "plco_cd_conta='" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "dist_vl_valor=" & FunTrataFloat(XLF_VALOR) & "," & _
                        "iden_cd_Projeto= " & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "" & _
                    " WHERE plco_cd_conta='" & XLO_TITULO!plco_cd_Conta & _
                        "' And cere_cd_pcr='" & XLO_TITULO!cere_cd_Pcr & "' and " & "nofi_cd_NotaFiscal = " & Chave
        Conexao.Execute XLT_SQL, XLI_REGATUALIZADOS
        
        'Se o compromisso não foi exportado anteriormente, será necessário inserir a distribuição
        If XLI_REGATUALIZADOS = 0 Then
            XLT_SQL = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta, " & _
                        "dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal)" & _
                    " VALUES ('" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "'" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "" & FunTrataFloat(XLF_VALOR) & "," & _
                        "" & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "," & Chave & ")"
            Conexao.Execute XLT_SQL
        End If
    End If
    'XLO_TITULO.Close
    If XLT_STATUS = "PAGAR" Then
        'SubContabilizarPromissoria XLL_CODDUPL, XLL_FATURA, XLO_TITULO!coco_cd_codigo, XLT_DOCREC, XLO_TITULO!titu_dt_Deposito, XLT_TIPOPAGTO, True
        'Cássio Medeiros
        '04/06/2014
        'Cópia de "SubContabilizarPromissoria" com a inclusão do códico da empresa destino como parâmetro
        SubContabilizarPromissoriaTransf XLL_CODDUPL, XLL_FATURA, XLO_TITULO!coco_cd_codigo, XLT_DOCREC, XLO_TITULO!titu_dt_Deposito, XLT_TIPOPAGTO, True, XLO_TITULO!empr_cd_Empresa
    End If
End Sub

'Cássio Medeiros
'04/06/2014
'Cópia de "SubContabilizarPromissoria" com a inclusão do códico da empresa destino como parâmetro
Sub SubContabilizarPromissoriaTransf(XLL_CODDUPLICATA As Long, XLL_FATURA As Long, XLL_CODCCORRENTE As Long, XLT_DOCREC As String, XLD_PAGTOREC As Date, XLT_TIPOPAGTO As String, XLB_CONSOLIDADO As Boolean, XLI_EMPRESA As Integer)
 
    '********************DECRIÇÃO DA PROCEDURE***********************************
    'Contabilização das Notas Promissorias - Apenas Base Caixa
    '****************************************************************************
    
    '**************PARÂMETROS*************************
    'XLL_CODDUPLICATA (Código da Duplicata)
    'XLL_FATURA (Nº da fatura da Parcela)
    'XLL_CODCCORRENTE (Código da conta corrente)
    'XLT_DOCREC (Nº do doc pag/rec)
    'XLD_PAGTOREC (Data de PagtoRec da Parcela)
    'XLB_CONSOLIDADO (Indica se o pagto é consolidado)
    '*************************************************
    
    Dim XLI_CODLANC As Integer      'Código para a tabela de Lançamentos
    Dim XLI_LOTE As Integer         'Nº do Lote
    Dim XLI_IDENTIFICADOR As Integer 'Código do identificador
    Dim XLO_PARCELA As Object       'Recordset com os dados das parcelas
    Dim XLO_VLCONSOLIDADO As ADODB.Recordset  'Recordset para obter o valor do pagto consolidado
    Dim XLO_PARCESPECIAL As ADODB.Recordset  'Recordset para tipos de documentos como TB, DEV, REST etc.
    Dim XLO_CHECK As ADODB.Recordset         'Recordset para  definir o lote
    Dim XLO_DESCONTOS As ADODB.Recordset      'Recordset para a tabela de descontoas/acrescimos
    Dim XLB_PROVISIONADO As Boolean 'Indica se o lançamento foi provisionado
    Dim XLT_SQL As String           'Prepara Select
    Dim XLT_PCRDEBITO As String, XLT_PCRCREDITO As String         'C. Custo de débito e crédito
    Dim XLT_CONTADEBITO As String, XLT_CONTACREDITO As String     'C. Contábil de débito e crédito
    Dim XLT_CONTADEBITOCC As String, XLT_CONTACREDITOCC As String 'C. Contábil (Banco ou forn) de débito e crédito
    Dim XLT_HISTORICO As String     'Histórico Contábil
    Dim XLT_HISTCUSTO As String     'Histórico do Custos = financeiro
    Dim XLT_TIPO As String          'Tipo do lançamento(crédito ou débito)
    Dim XLF_VALOR As Double         'Valor do lançamento
    Dim XLF_VALORBANCO As Double    'Valor da parte do Banco
    Dim XLF_VALORDESC As Double     'Valor de Desconto
    Dim XLF_VALORDESCACRESC As Double  'Valor do desconto acrescimo
    Dim XLF_VLCONSOLIDADO As Double 'Total de valores consolidados
    Dim XLF_VALORDUPL As Double     'Valor da parcela
    Dim XLV_POSICAO As Variant      'Armazena a posição do grid
    Dim XLF_TOTALPORCODDESC As Double
    Dim XLL_CODDESC As Long
    Dim XLL_CODDESCPROX As Long
    
    'Pegar o número do Lote e Lancamento para cadastrar no mesmo lote do dia
    'Origem = C - Capi
    XLT_SQL = "SELECT lanc_nr_Lote FROM Lancamentos " & _
                "WHERE lanc_dt_data = " & FunNuloData(XLD_PAGTOREC, NomeSgbd) & " AND " & _
                    "lanc_tx_origem LIKE 'C%' AND " & _
                    "empr_cd_empresa = " & XLI_EMPRESA & _
                 " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If XLO_CHECK.EOF Then
        XLI_LOTE = (FunDefinirUltimoLoteTransf(XLD_PAGTOREC, XLI_EMPRESA))
    Else
        XLI_LOTE = XLO_CHECK!lanc_nr_Lote
    End If
    XLO_CHECK.Close
    Set XLO_CHECK = Nothing
    
    'pegar o valor total para posterior contabilização da parte do banco
    XLT_SQL = "SELECT sum(dupl_vl_valor) AS ValorChCons, " & _
            " COUNT (dupl_cd_duplicata) AS TotalCheques FROM Duplicatas " & _
            " WHERE dupl_tx_tipopagrec = '" & XLT_TIPOPAGTO & "' AND coco_cd_codigo = " & XLL_CODCCORRENTE & " " & _
            " AND dupl_tx_docpagrec= '" & XLT_DOCREC & "'"
    SubQOpenRecordset XLO_VLCONSOLIDADO, XLT_SQL, Estatico
    XLF_VLCONSOLIDADO = XLO_VLCONSOLIDADO!ValorChCons
    XLO_VLCONSOLIDADO.Close
    Set XLO_VLCONSOLIDADO = Nothing
    
    XLT_SQL = "SELECT PlcoContaCC, cere_cd_Pcr, ContaRecDesp, nofi_tx_historico, focl_tx_RazaoSocial, nofi_nr_documento, iden_cd_Projeto, dupl_dt_pagrec, desconto, Acrescimo, distvalor, dupl_cd_Duplicata FROM ConsFINExpContabilidade2 " & _
                    " WHERE coco_cd_codigo = " & XLL_CODCCORRENTE & "" & _
                    " AND dupl_tx_tipopagrec ='" & XLT_TIPOPAGTO & "'" & _
                    " AND dupl_tx_docpagrec= '" & XLT_DOCREC & "'" & _
                    " AND empr_cd_empresa = " & XLI_EMPRESA & ""
    SubQOpenRecordset XLO_PARCELA, XLT_SQL, Estatico
    
    If Not XLO_PARCELA.EOF Then
        
        XLT_CONTADEBITOCC = XLO_PARCELA!PlcoContaCC
        
        While Not XLO_PARCELA.EOF
            '**********************NOTA PROMISSÓRIA***********************************
            XLT_PCRCREDITO = XLO_PARCELA!cere_cd_Pcr ''XLO_PARCESPECIAL!cere_cd_Pcr
            XLT_CONTACREDITO = XLO_PARCELA!ContaRecDesp ''XLO_PARCESPECIAL!ContaRecDesp
            XLT_PCRDEBITO = ""
            XLT_CONTADEBITO = "" 'Debitar depois a soma total do dia
            
            If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                XLT_HISTORICO = Mid(FunNulo(XLO_PARCELA!nofi_tx_historico) + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
            Else
                XLT_HISTORICO = Mid(FunNulo(XLO_PARCELA!nofi_tx_historico) + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
            End If
            
            XLT_HISTCUSTO = FunNulo(XLO_PARCELA!nofi_tx_historico)
            XLI_IDENTIFICADOR = FunNuloVal(XLO_PARCELA!iden_cd_Projeto)
            XLD_PAGTOREC = XLO_PARCELA!dupl_dt_pagrec
            XLF_VALORDESC = XLO_PARCELA!desconto + XLO_PARCELA!Acrescimo
            XLF_VALORDUPL = XLO_PARCELA!distvalor
            XLF_VALORBANCO = 0
            
            XLI_CODLANC = funDefinirCodigoLancTransf(XLO_PARCELA!dupl_dt_pagrec, XLI_LOTE, XLI_EMPRESA)
            
            XLT_SQL = "INSERT INTO Lancamentos(lanc_nr_Lote, lanc_cd_Lanc, " & _
                            " lanc_dt_Data, " & _
                            " cere_cd_pcrCreditar, cere_cd_pcrDebitar, " & _
                            " plco_cd_Credito, plco_cd_Debito," & _
                            " lanc_tx_historico, lanc_vl_valor," & _
                            " empr_cd_empresa, lanc_tx_DescCustos, " & _
                            " iden_cd_Projeto, lanc_cd_OrigemFin, lanc_tx_Origem) " & _
                        " VALUES (" & XLI_LOTE & ", " & XLI_CODLANC & "," & _
                            "" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "," & _
                            "" & FunNuloBanco(XLT_PCRCREDITO) & "," & FunNuloBanco(XLT_PCRDEBITO) & "," & _
                            "" & FunNuloBanco(XLT_CONTACREDITO) & "," & FunNuloBanco(XLT_CONTADEBITO) & "," & _
                            "'" & FunNulo(FunTrataString(XLT_HISTORICO)) & "'," & FunTrataFloat(XLF_VALORDUPL) & "," & _
                            "" & XLI_EMPRESA & ",'" & FunNulo(FunTrataString(XLT_HISTCUSTO)) & "'," & _
                            "" & FunNuloBanco(XLI_IDENTIFICADOR) & "," & XLL_CODDUPLICATA & ",'C')"
            Conexao.Execute XLT_SQL
            
            '******************************DESCONTOS/ACRÉSCIMOS******************************
            If XLO_PARCELA!desconto <> 0 Or XLO_PARCELA!Acrescimo <> 0 Then
                XLT_SQL = "SELECT deac_cd_descacresc, deac_tx_Descricao, deac_tx_Descricao, natu_cd_Natureza, plco_cd_Conta, cere_cd_Pcr, dude_vl_descacresc, TOTALDESCACRESC, deac_tx_tipo FROM ConsFINExpDescacresc WHERE dupl_cd_duplicata = " & XLO_PARCELA!dupl_cd_Duplicata & " and empr_cd_empresa = " & PCodEmpresa & ""
                SubQOpenRecordset XLO_DESCONTOS, XLT_SQL, Estatico
                
                If XLO_DESCONTOS.EOF Then
                    MsgBox "O Desconto/Acréscimo da parcela não possui conta contábil. Peça ao seu administrador para cadastrar a conta contábil pelo sistema QualiFin.", vbCritical, "ATENÇÃO!"
'                    XGB_ContinaPagamento = False 'Cancela o pagamento do título
                    XLO_DESCONTOS.Close
                    Set XLO_DESCONTOS = Nothing
                    Exit Sub
                Else
                    If IsNull(XLO_DESCONTOS!plco_cd_Conta) Or XLO_DESCONTOS!plco_cd_Conta = "" Then
                        MsgBox "O Desconto/Acréscimo '" & XLO_DESCONTOS!deac_tx_Descricao & "' não possui conta contábil. Peça ao seu administrador para cadastrar a conta contábil pelo sistema QualiFin.", vbCritical, "ATENÇÃO!"
'                        XGB_ContinaPagamento = False 'Cancela o pagamento do título
                        
                        XLO_DESCONTOS.Close
                        Set XLO_DESCONTOS = Nothing
                        Exit Sub
                    End If
                End If
                
                XLF_TOTALPORCODDESC = 0
                While Not XLO_DESCONTOS.EOF
                    XLL_CODDESC = XLO_DESCONTOS!deac_cd_descacresc
                    XLL_CODDESCPROX = XLO_DESCONTOS!deac_cd_descacresc
                    If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                        If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                        Else
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
                        End If
                    Else
                        If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                        Else
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
                        End If
                    End If
                   
                    'Se for receita ou despesa, deve informar o Centro de custo
                    If XLO_DESCONTOS!natu_cd_Natureza = 3 Or XLO_DESCONTOS!natu_cd_Natureza = 4 Then
                        If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                            XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRCREDITO = XLO_DESCONTOS!cere_cd_Pcr
                            XLT_PCRDEBITO = ""
                            XLT_CONTADEBITO = ""
                        Else 'Desconto
                            XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRDEBITO = XLO_DESCONTOS!cere_cd_Pcr
                            XLT_PCRCREDITO = ""
                            XLT_CONTACREDITO = ""
                        End If
                        XLF_VALOR = ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                        XLF_TOTALPORCODDESC = XLF_TOTALPORCODDESC + XLF_VALOR
                        
                        XLO_DESCONTOS.MoveNext
                        If Not XLO_DESCONTOS.EOF Then
                            If XLL_CODDESC <> XLO_DESCONTOS!deac_cd_descacresc Then
                                XLO_DESCONTOS.MovePrevious
                                If XLF_TOTALPORCODDESC > XLO_DESCONTOS!TOTALDESCACRESC Then
                                    XLF_VALOR = XLF_VALOR - 0.01
                                ElseIf XLF_TOTALPORCODDESC < XLO_DESCONTOS!TOTALDESCACRESC Then
                                    XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_TOTALPORCODDESC)
                                End If
                                XLF_TOTALPORCODDESC = 0
                            Else
                                XLO_DESCONTOS.MovePrevious
                            End If
                        Else
                            XLO_DESCONTOS.MovePrevious
                            If XLF_TOTALPORCODDESC > XLO_DESCONTOS!TOTALDESCACRESC Then
                                XLF_VALOR = XLF_VALOR - 0.01
                            Else
                                XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_TOTALPORCODDESC)
                            End If
                            XLF_TOTALPORCODDESC = 0
                        End If
                    Else
                        '***********
                        If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                            XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRDEBITO = ""
                            XLT_PCRCREDITO = ""
                            XLT_CONTACREDITO = ""
                        Else 'Desconto
                            XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRCREDITO = ""
                            XLT_PCRDEBITO = ""
                            XLT_CONTADEBITO = ""
                        End If
                        '*************
                        XLF_VALOR = 0
                        While XLL_CODDESC = XLL_CODDESCPROX And Not XLO_DESCONTOS.EOF
                            XLF_VALOR = XLF_VALOR + ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                            XLO_DESCONTOS.MoveNext
                            If Not XLO_DESCONTOS.EOF Then
                                XLL_CODDESCPROX = XLO_DESCONTOS!deac_cd_descacresc
                                If XLL_CODDESC = XLL_CODDESCPROX Then
                                    XLF_VALORDESC = XLF_VALORDESC + ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                                End If
                            End If
                        Wend
                        XLO_DESCONTOS.MovePrevious
                        
                        If XLF_VALOR > XLO_DESCONTOS!TOTALDESCACRESC Then
                            XLF_VALOR = XLF_VALOR - 0.01
                        Else
                            XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_VALOR)
                        End If
                        XLF_TOTALPORCODDESC = 0
                    End If
                    
                    XLI_CODLANC = funDefinirCodigoLancTransf(XLD_PAGTOREC, XLI_LOTE, XLI_EMPRESA)
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_nr_Lote, lanc_cd_Lanc, " & _
                            " lanc_dt_Data, " & _
                            " cere_cd_pcrCreditar, cere_cd_pcrDebitar, " & _
                            " plco_cd_Credito, plco_cd_Debito," & _
                            " lanc_tx_historico, lanc_vl_valor," & _
                            " empr_cd_empresa, lanc_tx_DescCustos, " & _
                            " iden_cd_Projeto, lanc_cd_OrigemFin, lanc_tx_Origem) " & _
                        " VALUES (" & XLI_LOTE & ", " & XLI_CODLANC & "," & _
                            "" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "," & _
                            "" & FunNuloBanco(XLT_PCRCREDITO) & "," & FunNuloBanco(XLT_PCRDEBITO) & "," & _
                            "" & FunNuloBanco(XLT_CONTACREDITO) & "," & FunNuloBanco(XLT_CONTADEBITO) & "," & _
                            "'" & FunNulo(FunTrataString(XLT_HISTORICO)) & "'," & FunTrataFloat(XLF_VALOR) & "," & _
                            "" & XLI_EMPRESA & ",'" & FunTrataString(FunNulo(XLO_PARCELA!nofi_tx_historico)) & "'," & _
                            "" & FunNuloBanco(XLO_PARCELA!iden_cd_Projeto) & "," & XLO_PARCELA!dupl_cd_Duplicata & ",'C')"
                    Conexao.Execute XLT_SQL
                    XLO_DESCONTOS.MoveNext
                Wend
                
                XLO_DESCONTOS.Close
                Set XLO_DESCONTOS = Nothing
            End If
'////////////////////
            XLO_PARCELA.MoveNext
        Wend
        XLO_PARCELA.Close
        Set XLO_PARCELA = Nothing
        

        '*********************PARTE DO BANCO DA NOTA PROMISSÓRIA********************
        XLT_HISTORICO = ""
        XLT_CONTACREDITO = ""
        XLT_PCRCREDITO = ""
        XLT_PCRDEBITO = ""
        XLT_HISTORICO = "Total diário de Notas Promissória " + Format(CStr(XLD_PAGTOREC), "dd/MM/yy")
        
        'Obter o total de NotasPromissórias
        XLT_SQL = "SELECT sum(titu_vl_ValorPago) AS TotalNP " & _
                    " FROM titulos INNER JOIN" & _
                    " Contratos ON Titulos.cont_cd_Contrato = Contratos.cont_cd_Contrato AND " & _
                    " Titulos.imov_cd_Imovel = Contratos.imov_cd_Imovel AND " & _
                    " Titulos.empd_cd_Empreendimento = Contratos.empd_cd_Empreendimento AND " & _
                    " Titulos.empr_cd_Empresa = Contratos.empr_cd_Empresa " & _
                    " WHERE " & _
                    " titu_dt_deposito = " & FunNuloData(XLD_PAGTOREC, NomeSgbd) & " " & _
                    " AND Titulos.empr_cd_empresa = " & XLI_EMPRESA & " AND Titulos.coco_cd_codigo = " & XLL_CODCCORRENTE & " "
        SubQOpenRecordset XLO_PARCELA, XLT_SQL, Dinamico
        
        If Not XLO_PARCELA.EOF Then
            XLF_VLCONSOLIDADO = Format(XLO_PARCELA!TotalNP, "standard")
        End If
        
        XLO_PARCELA.Close
        Set XLO_PARCELA = Nothing
        
        XLT_SQL = "SELECT * FROM Lancamentos " & _
                    " WHERE plco_cd_Debito = '" & XLT_CONTADEBITOCC & "'" & _
                    " AND lanc_dt_Data =" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "" & _
                    " AND empr_cd_empresa = " & XLI_EMPRESA & "  and lanc_tx_origem = 'C'"
        SubQOpenRecordset XLO_PARCELA, XLT_SQL, Dinamico
        
        If XLO_PARCELA.EOF Then
            XLO_PARCELA.AddNew
            XLI_CODLANC = funDefinirCodigoLancTransf(XLD_PAGTOREC, XLI_LOTE, XLI_EMPRESA)
            XLO_PARCELA!lanc_cd_Lanc = XLI_CODLANC
        End If
                
        XLO_PARCELA!lanc_nr_Lote = XLI_LOTE
        XLO_PARCELA!lanc_dt_Data = XLD_PAGTOREC
        XLO_PARCELA!cere_cd_pcrCreditar = Null 'XLT_PCRCREDITO
        XLO_PARCELA!cere_cd_pcrDebitar = Null 'XLT_PCRDEBITO
        XLO_PARCELA!plco_cd_credito = Null 'XLT_CONTACREDITO
        XLO_PARCELA!plco_cd_Debito = XLT_CONTADEBITOCC
        XLO_PARCELA!lanc_tx_historico = XLT_HISTORICO
        XLO_PARCELA!lanc_vl_valor = XLF_VLCONSOLIDADO
        XLO_PARCELA!empr_cd_Empresa = XLI_EMPRESA
        XLO_PARCELA!lanc_tx_DescCustos = XLT_HISTCUSTO
        XLO_PARCELA!iden_cd_Projeto = funNull(XLI_IDENTIFICADOR)
        XLO_PARCELA!lanc_cd_OrigemFin = XLL_CODDUPLICATA
        XLO_PARCELA!lanc_tx_Origem = "C"
                                        
        XLO_PARCELA.Update
        XLO_PARCELA.Close
        
        Set XLO_PARCELA = Nothing
    End If
End Sub

'Cássio Medeiros
'04/06/2014
'Cópia de "FunDefinirUltimoLote" com a inclusão do códico da empresa como parâmetro
Function FunDefinirUltimoLoteTransf(Data As Date, XLI_EMPRESA As Integer) As Integer
  '*****************************************************************************
  'Define o numero do maior Lote existente para a data do lancamento
  'Chamado qnd abre o Formulario de Lancamento e qnd Muda a Data
  '*****************************************************************************
  Dim XLO_RSLANC As New ADODB.Recordset
  Dim XLT_SQL As String
        
  XLT_SQL = "SELECT Max(lanc_nr_lote) as Lanc_nr_Lote " & _
              "FROM Lancamentos " & _
              "WHERE lanc_dt_data = " & FunNuloData(Data, NomeSgbd) & _
              "  AND empr_cd_empresa = " & XLI_EMPRESA
  SubQOpenRecordset XLO_RSLANC, XLT_SQL, Estatico
  
  'se o retorno do select for vazio eh pq não houve nenhum lancamento para
  'essa data, entao o lote vai para 1. se houver retorno entao afuncao
  'retorna o numero do maior lote para a data
  If IsNull(XLO_RSLANC("lanc_nr_Lote")) Then
    FunDefinirUltimoLoteTransf = 1
  Else
    FunDefinirUltimoLoteTransf = XLO_RSLANC("Lanc_nr_Lote").Value + 1
  End If
  
  XLO_RSLANC.Close
  Set XLO_RSLANC = Nothing
  
End Function

'Cássio Medeiros
'04/06/2014
'Cópia de "funDefinirCodigoLanc" com a inclusão do códico da empresa como parâmetro
Function funDefinirCodigoLancTransf(XLD_DATA As Date, XLI_NrLote As Integer, XLI_EMPRESA As Integer) As Integer
  '**********************************************************************
  'Encontra o próximo cod. livre de lancamento para a data e o lote atual
  '*****************************************************************************
  Dim XLO_RSLANC As New ADODB.Recordset

  XLO_RSLANC.Open _
    "SELECT max(lanc_cd_lanc) as lanc_cd_lanc " & _
    "FROM Lancamentos " & _
    "WHERE lanc_dt_data = " & FunNuloData(XLD_DATA, NomeSgbd) & _
    " AND lanc_nr_Lote = " & XLI_NrLote & _
    " AND empr_cd_empresa = " & XLI_EMPRESA, Conexao
      
  If Not IsNull(XLO_RSLANC("lanc_cd_Lanc").Value) Then
    funDefinirCodigoLancTransf = XLO_RSLANC("lanc_cd_Lanc").Value + 1
  Else
    funDefinirCodigoLancTransf = 1
  End If
  
  XLO_RSLANC.Close
  Set XLO_RSLANC = Nothing
End Function

