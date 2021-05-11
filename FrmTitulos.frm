VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Títulos"
   ClientHeight    =   5265
   ClientLeft      =   2130
   ClientTop       =   2760
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9375
   Begin VB.Frame FraPrincipal 
      Height          =   4950
      Left            =   60
      TabIndex        =   31
      Top             =   -60
      Width           =   9345
      Begin VB.Frame FraPlano1 
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   1035
         Width           =   9135
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmTitulos.frx":0000
            Left            =   7140
            List            =   "FrmTitulos.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   120
            Width           =   1875
         End
         Begin MSMask.MaskEdBox MskPlano 
            Height          =   315
            Left            =   780
            TabIndex        =   3
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmTitulos.frx":0051
            Height          =   315
            Left            =   3780
            TabIndex        =   4
            Top             =   120
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tipl_tx_Descricao"
            BoundColumn     =   "tipl_cd_TipoPlano"
            Text            =   "CboTiposPlanos"
         End
         Begin MSAdodcLib.Adodc DatTipoPlano 
            Height          =   330
            Left            =   3960
            Top             =   120
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
            Caption         =   "DatTipoPlano"
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
         Begin VB.Label LblNaturezaPlano 
            Caption         =   "Natureza:"
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
            Left            =   6300
            TabIndex        =   65
            Top             =   180
            Width           =   915
         End
         Begin VB.Label LblParcela2 
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
            Left            =   2520
            TabIndex        =   62
            Top             =   180
            Width           =   495
         End
         Begin VB.Label LblTipoPlano 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
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
            Height          =   210
            Left            =   3240
            TabIndex        =   42
            Top             =   180
            Width           =   495
         End
         Begin VB.Label LblPlano 
            Alignment       =   1  'Right Justify
            Caption         =   "Plano:"
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
            TabIndex        =   41
            Top             =   180
            Width           =   570
         End
         Begin VB.Label LblParcela 
            Caption         =   "Parcela:"
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
            Left            =   1800
            TabIndex        =   40
            Top             =   180
            Width           =   735
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   3375
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Dados Financeiros"
         TabPicture(0)   =   "FrmTitulos.frx":006C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblObs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CmdLimparObs"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DatObs"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CboObs"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "FraPlano2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "FraVlIndexador"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "ChkExportou"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Dados Contábeis"
         TabPicture(1)   =   "FrmTitulos.frx":0088
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DatIdentificador"
         Tab(1).Control(1)=   "DatContaCorrente"
         Tab(1).Control(2)=   "DatConta"
         Tab(1).Control(3)=   "DatPcr"
         Tab(1).Control(4)=   "FraConta"
         Tab(1).ControlCount=   5
         Begin VB.CheckBox ChkExportou 
            Caption         =   "Exportou para o Banco"
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
            Left            =   6660
            TabIndex        =   72
            Top             =   3000
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Frame FraConta 
            Height          =   1890
            Left            =   -74865
            TabIndex        =   66
            Top             =   420
            Width           =   8835
            Begin VB.TextBox TxtContratoBanco 
               Height          =   315
               Left            =   7800
               MaxLength       =   3
               TabIndex        =   73
               Top             =   1440
               Width           =   615
            End
            Begin MSMask.MaskEdBox Mskconta 
               Height          =   330
               Left            =   1440
               TabIndex        =   19
               Top             =   180
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas 
               Bindings        =   "FrmTitulos.frx":00A4
               Height          =   330
               Left            =   2340
               TabIndex        =   20
               Top             =   180
               Width           =   6030
               _ExtentX        =   10636
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
            Begin MSMask.MaskEdBox MskPcr 
               Height          =   330
               Left            =   1440
               TabIndex        =   22
               Top             =   585
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmTitulos.frx":00BB
               Height          =   330
               Left            =   2340
               TabIndex        =   23
               Top             =   600
               Width           =   6030
               _ExtentX        =   10636
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
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "FrmTitulos.frx":00D0
               Height          =   330
               Left            =   1440
               TabIndex        =   25
               Top             =   960
               Width           =   6915
               _ExtentX        =   12197
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
               Height          =   315
               Left            =   8385
               TabIndex        =   26
               Top             =   960
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":00EF
               Picture         =   "FrmTitulos.frx":0541
            End
            Begin Threed.SSCommand CmdLimparPcr 
               Height          =   315
               Left            =   8400
               TabIndex        =   24
               Top             =   585
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":0653
               Picture         =   "FrmTitulos.frx":0AA5
            End
            Begin Threed.SSCommand CmdLimparplanoContas 
               Height          =   315
               Left            =   8400
               TabIndex        =   21
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":0BB7
               Picture         =   "FrmTitulos.frx":1009
            End
            Begin MSDataListLib.DataCombo CboIdentificador 
               Bindings        =   "FrmTitulos.frx":111B
               Height          =   315
               Left            =   1440
               TabIndex        =   27
               Top             =   1350
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "iden_tx_descricao"
               BoundColumn     =   "iden_cd_Projeto"
               Text            =   ""
            End
            Begin Threed.SSCommand CmdLimparIdentif 
               Height          =   315
               Left            =   5040
               TabIndex        =   28
               Top             =   1350
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":113A
               Picture         =   "FrmTitulos.frx":158C
            End
            Begin VB.Label lblContratoBanco 
               Caption         =   "Nº do Contrato no Banco:"
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
               Left            =   5520
               TabIndex        =   74
               Top             =   1440
               Width           =   2415
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
               Left            =   270
               TabIndex        =   71
               Top             =   1380
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
               Left            =   105
               TabIndex        =   69
               Top             =   255
               Width           =   1320
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
               Height          =   195
               Left            =   165
               TabIndex        =   68
               Top             =   645
               Width           =   1260
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
               Height          =   255
               Left            =   60
               TabIndex        =   67
               Top             =   1020
               Width           =   1365
            End
         End
         Begin VB.Frame FraVlIndexador 
            Caption         =   "Valor Indexado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   120
            TabIndex        =   55
            Top             =   1920
            Width           =   8835
            Begin MSDataListLib.DataCombo CboMoeda2 
               Bindings        =   "FrmTitulos.frx":169E
               Height          =   315
               Left            =   6060
               TabIndex        =   15
               Top             =   180
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboImovel"
            End
            Begin MSDataListLib.DataCombo CboMoeda1 
               Bindings        =   "FrmTitulos.frx":16B6
               Height          =   315
               Left            =   1800
               TabIndex        =   13
               Top             =   180
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboMoeda1"
            End
            Begin MSAdodcLib.Adodc DatMoeda1 
               Height          =   330
               Left            =   6060
               Top             =   180
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
               Caption         =   "DatMoeda1"
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
            Begin MSAdodcLib.Adodc DatMoeda2 
               Height          =   330
               Left            =   1800
               Top             =   180
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
               Caption         =   "DatMoeda2"
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
            Begin Threed.SSCommand CmdLimparIndex1 
               Height          =   315
               Left            =   3600
               TabIndex        =   14
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":16CE
               Picture         =   "FrmTitulos.frx":1B20
            End
            Begin Threed.SSCommand CmdLimparIndex2 
               Height          =   315
               Left            =   7860
               TabIndex        =   16
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmTitulos.frx":1C32
               Picture         =   "FrmTitulos.frx":2084
            End
            Begin VB.Label LblValorMoeda12 
               Alignment       =   1  'Right Justify
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
               Left            =   1800
               TabIndex        =   64
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label LblValorMoeda22 
               Alignment       =   1  'Right Justify
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
               Left            =   6060
               TabIndex        =   63
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label LblValorMoeda2 
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
               Height          =   255
               Left            =   1080
               TabIndex        =   61
               Top             =   600
               Width           =   690
            End
            Begin VB.Label LblValorMoeda1 
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
               Height          =   255
               Left            =   5340
               TabIndex        =   60
               Top             =   600
               Width           =   690
            End
            Begin VB.Label LblMoeda2 
               Alignment       =   1  'Right Justify
               Caption         =   "Indexador 2:"
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
               Left            =   4920
               TabIndex        =   59
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label LblMoeda1 
               Alignment       =   1  'Right Justify
               Caption         =   "Indexador 1:"
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
               Left            =   600
               TabIndex        =   58
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label LblSimbolo1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   3600
               TabIndex        =   57
               Top             =   600
               Width           =   750
            End
            Begin VB.Label LblSimbolo2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   7860
               TabIndex        =   56
               Top             =   600
               Width           =   810
            End
         End
         Begin VB.Frame FraPlano2 
            Height          =   915
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   8835
            Begin VB.TextBox TxtJuros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   10
               Top             =   540
               Width           =   975
            End
            Begin VB.TextBox TxtMulta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6420
               TabIndex        =   11
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtJurosFin 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   9
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtValorReal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6420
               TabIndex        =   12
               Top             =   540
               Width           =   1515
            End
            Begin VB.Label LblValorReal 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Histórico da Parcela:"
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
               Left            =   3930
               TabIndex        =   70
               Top             =   600
               Width           =   2445
            End
            Begin VB.Label LblAD 
               Caption         =   "a.d."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   3240
               TabIndex        =   54
               Top             =   600
               Width           =   615
            End
            Begin VB.Label LblJuros 
               Alignment       =   1  'Right Justify
               Caption         =   "% Juros de Mora:"
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
               Left            =   180
               TabIndex        =   53
               Top             =   600
               Width           =   1920
            End
            Begin VB.Label LblMulta 
               Alignment       =   1  'Right Justify
               Caption         =   "% Multa:"
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
               Left            =   5580
               TabIndex        =   52
               Top             =   240
               Width           =   795
            End
            Begin VB.Label LblProRata 
               Alignment       =   1  'Right Justify
               Caption         =   "Pró-Rata"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   7560
               TabIndex        =   51
               Top             =   240
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label LblJurosFin 
               Alignment       =   1  'Right Justify
               Caption         =   "% Deságio:"
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
               Left            =   60
               TabIndex        =   50
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label LblAM 
               Alignment       =   1  'Right Justify
               Caption         =   "a.m."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   3240
               TabIndex        =   49
               Top             =   240
               Width           =   390
            End
            Begin VB.Label LblReal 
               Caption         =   "R$"
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
               Left            =   8040
               TabIndex        =   48
               Top             =   600
               Width           =   315
            End
         End
         Begin VB.Frame Frame1 
            Height          =   615
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   8835
            Begin MSComCtl2.DTPicker DtpVencimento 
               Height          =   315
               Left            =   4275
               TabIndex        =   7
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   64094211
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtBase 
               Height          =   315
               Left            =   1200
               TabIndex        =   6
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   64094211
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtProrrogacao 
               Height          =   315
               Left            =   7275
               TabIndex        =   8
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   64094211
               CurrentDate     =   37180
            End
            Begin VB.Label LblProrrogacao 
               Caption         =   "Prorrogação:"
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
               Left            =   6120
               TabIndex        =   46
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label LblVencimnto 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Left            =   2655
               TabIndex        =   45
               Top             =   240
               Width           =   1560
            End
            Begin VB.Label LblDtBase 
               Alignment       =   1  'Right Justify
               Caption         =   "Data Base:"
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
               TabIndex        =   44
               Top             =   240
               Width           =   1050
            End
         End
         Begin MSAdodcLib.Adodc DatPcr 
            Height          =   330
            Left            =   -71100
            Top             =   2760
            Visible         =   0   'False
            Width           =   2115
            _ExtentX        =   3731
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
            Left            =   -68280
            Top             =   2835
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
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
            Left            =   -73140
            Top             =   2880
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
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
         Begin MSDataListLib.DataCombo CboObs 
            Bindings        =   "FrmTitulos.frx":2196
            Height          =   315
            Left            =   1260
            TabIndex        =   17
            Top             =   2940
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "obse_tx_observacao"
            BoundColumn     =   "obse_cd_observacao"
            Text            =   "CboObs"
         End
         Begin MSAdodcLib.Adodc DatObs 
            Height          =   330
            Left            =   2280
            Top             =   2940
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
            Caption         =   "DatObs"
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
         Begin Threed.SSCommand CmdLimparObs 
            Height          =   315
            Left            =   6120
            TabIndex        =   18
            Top             =   2940
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmTitulos.frx":21AB
            Picture         =   "FrmTitulos.frx":25FD
         End
         Begin MSAdodcLib.Adodc DatIdentificador 
            Height          =   330
            Left            =   -74760
            Top             =   2760
            Visible         =   0   'False
            Width           =   2610
            _ExtentX        =   4604
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
         Begin VB.Label LblObs 
            Alignment       =   1  'Right Justify
            Caption         =   "Observação:"
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
            Left            =   120
            TabIndex        =   38
            Top             =   3000
            Width           =   1140
         End
      End
      Begin Threed.SSPanel PanCodigo 
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   9120
         _Version        =   65536
         _ExtentX        =   16087
         _ExtentY        =   1508
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
         Begin MSDataListLib.DataCombo CboImovel 
            Bindings        =   "FrmTitulos.frx":270F
            Height          =   315
            Left            =   1740
            TabIndex        =   2
            Top             =   480
            Width           =   5985
            _ExtentX        =   10557
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
            Left            =   1740
            TabIndex        =   0
            Top             =   60
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
            Bindings        =   "FrmTitulos.frx":2727
            Height          =   315
            Left            =   2460
            TabIndex        =   1
            Top             =   60
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "empd_tx_nome"
            BoundColumn     =   "empd_cd_empreendimento"
            Text            =   "CboEmpreendimento"
         End
         Begin MSAdodcLib.Adodc DatEmpreendimento 
            Height          =   330
            Left            =   3060
            Top             =   60
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
            Left            =   2100
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
         Begin VB.Label LblContrato2 
            Caption         =   "01"
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
            Left            =   8700
            TabIndex        =   36
            Top             =   540
            Width           =   255
         End
         Begin VB.Label LblContrato 
            Caption         =   "Contrato:"
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
            Left            =   7860
            TabIndex        =   35
            Top             =   540
            Width           =   795
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
            TabIndex        =   34
            Top             =   135
            Width           =   1485
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
            Left            =   840
            TabIndex        =   33
            Top             =   540
            Width           =   825
         End
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5835
      TabIndex        =   30
      Top             =   4920
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
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   2040
      TabIndex        =   29
      Top             =   4920
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
End
Attribute VB_Name = "FrmFinTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim errLoop As ADODB.Error
Dim StrError As String

Dim XFT_CODIMOVEL As String, XFT_CODEMPREENDIMENTO As String
Dim XFT_CODCONTRATO As String, XFT_CODPLANO As String
Dim XFT_CODPARCELA As String, XFT_CODRESIDUO As String
Dim XFT_PRORATA As String
Dim XFT_CONTROLE As String
Dim XFB_ALTEROU As Boolean

Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If FunObrigatorioMSK(MskEmpreendimento, "O Empreendimento é Obrigatório") Then Exit Function
    If FunObrigatorioCBO(CboImovel, "O Imóvel é Obrigatório") Then Exit Function
    If FunObrigatorioMSK(MskPlano, "O Plano é Obrigatório") Then Exit Function
    If FunObrigatorioCBO(CboTiposPlanos, "O Tipo de Plano é Obrigatório") Then Exit Function
    If FunObrigatorioCBOX(CboNaturezasPlanos, "A Natureza do Plano é Obrigatório") Then Exit Function
    If FunObrigatorioTXT(TxtValorReal, "O Valor Histórico da parcelas é Obrigatório") Then Exit Function
    If FunObrigatorioDTP(DtpDtBase, "A Data Base é Obrigatório") Then Exit Function
    If FunObrigatorioDTP(DtpVencimento, "A Data de Vencimento da Primeira Parcela é Obrigatório") Then Exit Function
    
    If XGB_PQUALIFIN Then 'Integrado com o Financeiro
        If FunObrigatorioCBO(CboCCorrente, "A Conta Corrente é Obrigatória") Then Exit Function
        If FunObrigatorioCBO(CboPcr, "O Centro de Custo é Obrigatório") Then Exit Function
        If FunObrigatorioCBO(CboPlanoContas, "A Conta Contábil é Obrigatória") Then Exit Function
    End If
        
    funChecarCampos = False
End Function
Sub SubDesabHabCampos(XHab As Boolean)
    
    'Desabilita/Habilita os campos em caso de inserção
    
    DtpDtBase.Enabled = XHab
    DtpVencimento.Enabled = XHab
    DtpDtProrrogacao.Enabled = XHab
    CboTiposPlanos.Enabled = XHab
    CboNaturezasPlanos.Enabled = XHab
    CboObs.Enabled = XHab
    TxtJurosFin.Enabled = XHab
    TxtJuros.Enabled = XHab
    TxtMulta.Enabled = XHab
    TxtValorReal.Enabled = XHab
    LblValorMoeda12.Enabled = XHab
    LblValorMoeda22.Enabled = XHab
    CboMoeda1.Enabled = XHab
    CboMoeda2.Enabled = XHab
    MskPcr.Enabled = XHab
    CboPcr.Enabled = XHab
    MskConta.Enabled = XHab
    CboPlanoContas.Enabled = XHab
    CboCCorrente.Enabled = XHab
    CboIdentificador.Enabled = XHab
    TabAuxiliar.Enabled = XHab
    CmdLimparCCorrente.Enabled = XHab
    CmdLimparIndex1.Enabled = XHab
    CmdLimparIndex2.Enabled = XHab
    CmdLimparObs.Enabled = XHab
    CmdLimparPcr.Enabled = XHab
    CmdLimparplanoContas.Enabled = XHab
    CmdLimparIdentif.Enabled = XHab
    Me.txtContratoBanco.Enabled = XHab
    
    If XInserir = "C" Then
      MskEmpreendimento.Enabled = XHab
      CboEmpreendimento.Enabled = XHab
      CboImovel.Enabled = XHab
      MskPlano.Enabled = XHab
      CmdGravar.Enabled = XHab
      TabAuxiliar.Enabled = True
    End If
    
End Sub

Sub subPreencheCampos()
    Dim XLT_SQL As String
    Dim XLO_TITULO As Object
    
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND titu_cd_Plano='" & XFT_CODPLANO & "' AND titu_cd_Parcela='" & XFT_CODPARCELA & "' AND titu_cd_Residuo='" & XFT_CODRESIDUO & "'", "Estatico"
    If Not (XLO_TITULO.BOF And XLO_TITULO.EOF) Then
      
      If XInserir <> "I" Then
        MskEmpreendimento.Text = XFT_CODEMPREENDIMENTO
        CboEmpreendimento.BoundText = XFT_CODEMPREENDIMENTO
        CboImovel.BoundText = XFT_CODIMOVEL
        LblContrato2.Caption = XFT_CODCONTRATO
        MskPlano.Text = XFT_CODPLANO
        LblParcela2.Caption = XFT_CODPARCELA
      End If
      
      Me.txtContratoBanco.Text = FunNulo(XLO_TITULO!titu_nr_ContratoBanco)
        
      CboTiposPlanos.BoundText = XLO_TITULO!tipl_cd_TipoPlano
      
       Select Case XLO_TITULO!titu_tx_NaturezaPlano
          Case "S"
               CboNaturezasPlanos.ListIndex = 0
          Case "P"
               CboNaturezasPlanos.ListIndex = 1
          Case "F"
               CboNaturezasPlanos.ListIndex = 2
          Case "G"
               CboNaturezasPlanos.ListIndex = 3
          Case "O"
               CboNaturezasPlanos.ListIndex = 4
          Case "V"
               CboNaturezasPlanos.ListIndex = 5
      End Select
      
      If Not IsNull(XLO_TITULO!titu_dt_Base) Then DtpDtBase.Value = XLO_TITULO!titu_dt_Base
      If Not IsNull(XLO_TITULO!titu_dt_Vencimento) Then DtpVencimento.Value = XLO_TITULO!titu_dt_Vencimento
      If Not IsNull(XLO_TITULO!titu_dt_Prorrogacao) Then DtpDtProrrogacao.Value = XLO_TITULO!titu_dt_Prorrogacao
                 
      TxtJurosFin.Text = Format(XLO_TITULO!titu_vl_PercJurosFin, "##0.0000")
      TxtJuros.Text = Format(XLO_TITULO!titu_vl_PercJurosMora, "##0.0000")
      TxtMulta.Text = Format(XLO_TITULO!titu_vl_PercMultaMora, "##0.0000")
      TxtValorReal.Text = Format(XLO_TITULO!titu_vl_Parcela, "standard")
      
      If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
        CboMoeda1.BoundText = XLO_TITULO!moed_cd_Moeda1
        LblValorMoeda12.Caption = Format(XLO_TITULO!titu_vl_ParcelaIndex1, "##0.0000")
        LblSimbolo1.Caption = XLO_TITULO!moeda1
      End If
      
      If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
        CboMoeda2.BoundText = XLO_TITULO!moed_cd_Moeda2
        LblValorMoeda22.Caption = Format(XLO_TITULO!titu_vl_ParcelaIndex2, "##0.0000")
        LblSimbolo2.Caption = XLO_TITULO!moeda2
      End If
      
      If Not IsNull(XLO_TITULO!obse_cd_Observacao) Then CboObs.BoundText = XLO_TITULO!obse_cd_Observacao
      If Not IsNull(XLO_TITULO!coco_cd_codigo) Then CboCCorrente.BoundText = XLO_TITULO!coco_cd_codigo
      
      If Not DatConta.Recordset.EOF Then
        DatConta.Recordset.MoveFirst
        If Not IsNull(XLO_TITULO!plco_cd_Conta) Then
            DatConta.Recordset.Find "plco_cd_Conta =" & XLO_TITULO!plco_cd_Conta
            MskConta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
            CboPlanoContas.BoundText = MskConta.Text
        Else
            MskConta.Text = ""
            CboPlanoContas.BoundText = ""
        End If
      End If

      If Not DatPcr.Recordset.EOF Then
        DatPcr.Recordset.MoveFirst
        If Not IsNull(XLO_TITULO!cere_cd_Pcr) Then
            DatPcr.Recordset.Find "cere_cd_pcr =" & XLO_TITULO!cere_cd_Pcr
            MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
            CboPcr.BoundText = MskPcr.Text
        Else
            MskPcr.Text = ""
            CboPcr.BoundText = ""
        End If
      End If

      If Not IsNull(XLO_TITULO!iden_cd_Projeto) Then
        CboIdentificador.BoundText = XLO_TITULO!iden_cd_Projeto
      End If
      
      If XLO_TITULO!titu_tx_ExpBanco = "S" Then
         ChkExportou.Visible = True
         ChkExportou.Value = 1
      Else
         ChkExportou.Value = 0
      End If
    End If
    
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
      
End Sub

Sub subPrepararInsercao()
   
    CboNaturezasPlanos.ListIndex = -1
    DtpDtBase.Value = Date
    DtpVencimento.Value = Date
    CboObs.Text = ""
    TxtJurosFin.Text = ""
    TxtJuros.Text = ""
    TxtMulta.Text = ""
    TxtValorReal.Text = ""
    LblValorMoeda12.Caption = ""
    LblValorMoeda22.Caption = ""
    CboMoeda1.Text = ""
    CboMoeda2.Text = ""
    MskPcr.Text = ""
    CboPcr.BoundText = ""
    MskConta.Text = ""
    CboPlanoContas.Text = ""
    CboCCorrente.Text = ""
    MskPlano.Text = ""
    CboIdentificador.Text = ""
    CboImovel.Enabled = False
    
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        CboImovel.Enabled = True
        CboImovel.Text = ""
        subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='A' ORDER BY imov_cd_Imovel", Estatico
    End If
End Sub

Private Sub CboEmpreendimento_DblClick(Area As Integer)
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub

Private Sub CboIdentificador_Change()
  
  If CboIdentificador.BoundText <> "" Then
    DatIdentificador.Recordset.bookmark = CboIdentificador.SelectedItem
  End If
  
End Sub

Private Sub CboImovel_Change()
    
    If CboImovel.Text <> "" Then
        DatImovel.Recordset.bookmark = CboImovel.SelectedItem
        'LblRazaoSocial.Caption = DatImovel.Recordset.Fields("Cliente")
        If Not IsNull(DatImovel.Recordset.Fields("cont_dt_base")) Then DtpDtBase.Value = DatImovel.Recordset.Fields("cont_dt_base")
        Select Case DatImovel.Recordset.Fields("cont_tx_TipoJurosMora")
            Case "D"
                LblAD.Caption = "a.d."
            Case "M"
                LblAD.Caption = "a.m."
            Case Else
                LblAD.Caption = "Cheio"
        End Select
        
        If DatImovel.Recordset.Fields("cont_tx_desagio") = "S" Then
            LblJurosFin.Caption = "% Juros C.E.F.:"
        Else
            LblJurosFin.Caption = "% Deságio:"
        End If
        
'        If DatImovel.Recordset.Fields("cont_tx_CorrecProRata") = 1 Then
'          XFB_PRORATA = True
'        Else
'          XFB_PRORATA = False
'        End If
        XFT_PRORATA = DatImovel.Recordset.Fields("cont_tx_CorrecProRata")
        
        If XInserir = "I" Then
        
          TxtJuros.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosMora, "##0.0000")
          If DatImovel.Recordset.Fields("cont_tx_multaprorata") Then LblProRata.Visible = True
          TxtMulta.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosMulta, "##0.0000")
          TxtJurosFin.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosFin, "##0.0000")
          
          LblContrato2.Caption = DatImovel.Recordset.Fields("cont_cd_Contrato")
          DtpDtBase.Value = DatImovel.Recordset.Fields!cont_dt_Base
          
          If Not DatConta.Recordset.EOF Then
            DatConta.Recordset.MoveFirst
            If Not IsNull(DatImovel.Recordset.Fields!plco_cd_Conta) Then
              DatConta.Recordset.Find "plco_cd_Conta =" & DatImovel.Recordset.Fields!plco_cd_Conta
              MskConta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
              CboPlanoContas.BoundText = MskConta.Text
            Else
              MskConta.Text = ""
              CboPlanoContas.BoundText = ""
            End If
          End If
            
          If Not DatPcr.Recordset.EOF Then
            DatPcr.Recordset.MoveFirst
            If Not IsNull(DatImovel.Recordset.Fields!cere_cd_Pcr) Then
                DatPcr.Recordset.Find "cere_cd_pcr =" & DatImovel.Recordset.Fields!cere_cd_Pcr
                MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
                CboPcr.BoundText = MskPcr.Text
            Else
              MskPcr.Text = ""
              CboPcr.BoundText = ""
            End If
          End If
          
          If Not DatContaCorrente.Recordset.EOF Then
          DatContaCorrente.Recordset.MoveFirst
            If Not IsNull(DatImovel.Recordset.Fields!coco_cd_codigo) Then
                DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & DatImovel.Recordset.Fields!coco_cd_codigo
                CboCCorrente.BoundText = FunNulo(DatImovel.Recordset.Fields!coco_cd_codigo)
            Else
              CboCCorrente.BoundText = ""
            End If
          End If
          
          If Not IsNull(DatImovel.Recordset.Fields!iden_cd_Projeto) Then
            CboIdentificador.BoundText = DatImovel.Recordset.Fields!iden_cd_Projeto
          End If
        End If
        
    End If
End Sub

Function FunGeraParcela() As String
    
  Dim XLO_TITULO As New ADODB.Recordset
     
  SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapTitulos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "' and imov_cd_imovel = '" & CboImovel.BoundText & "' AND cont_cd_Contrato= '" & LblContrato2.Caption & "' AND titu_cd_Plano='" & Format(MskPlano.Text, "00") & "'  ORDER BY Titulo desc", Estatico
  If Not XLO_TITULO.EOF Then
    FunGeraParcela = CStr(CInt(Mid$(XLO_TITULO!Titulo, 17, 3)) + 1)
  Else
    FunGeraParcela = ""
  End If
  XLO_TITULO.Close
  Set XLO_TITULO = Nothing
    
End Function

Private Sub CboMoeda1_Change()

  If CboMoeda1.Text <> "" Then
    DatMoeda1.Recordset.bookmark = CboMoeda1.SelectedItem
    If TxtValorReal.Text <> "" Then
      LblValorMoeda12.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
      If Not IsNull(DatMoeda1.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda12 <> "" Then
        LblSimbolo1.Caption = DatMoeda1.Recordset.Fields("moed_tx_simbolo")
      End If
      If LblValorMoeda12.Caption = "" Then
        CboMoeda1.BoundText = ""
        LblSimbolo1.Caption = ""
      End If
    End If
  End If

End Sub

Private Sub CboMoeda2_Change()

  If CboMoeda2.Text <> "" Then
    DatMoeda2.Recordset.bookmark = CboMoeda2.SelectedItem
    If TxtValorReal.Text <> "" Then
      LblValorMoeda22.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
      If Not IsNull(DatMoeda2.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda22.Caption <> "" Then
        LblSimbolo2.Caption = DatMoeda2.Recordset.Fields("moed_tx_simbolo")
      End If
      If LblValorMoeda22.Caption = "" Then
        CboMoeda2.BoundText = ""
        LblSimbolo2.Caption = ""
      End If
    End If
  End If

End Sub



Private Sub CmdLimparCCorrente_Click()
  CboCCorrente.BoundText = ""
End Sub

Private Sub CmdLimparIndex1_Click()
  CboMoeda1.BoundText = ""
  LblValorMoeda12.Caption = ""
  LblSimbolo1.Caption = ""
End Sub

Private Sub CmdLimparIndex2_Click()
  CboMoeda2.BoundText = ""
  LblValorMoeda22.Caption = ""
  LblSimbolo2.Caption = ""
End Sub

Private Sub CmdLimparObs_Click()
  CboObs.BoundText = ""
End Sub

Private Sub CmdLimparPcr_Click()
  CboPcr.BoundText = ""
  MskPcr.Text = ""
End Sub

Private Sub CmdLimparplanoContas_Click()
  CboPlanoContas.BoundText = ""
  MskConta.Text = ""
End Sub

Private Sub DtpDtBase_Change()
  
  XFT_CONTROLE = ""
  XFB_ALTEROU = True
  
End Sub

Private Sub DtpDtBase_LostFocus()

  
  If XFB_ALTEROU = True Then
    If XFT_CONTROLE <> "TxtValorReal" Then
      'Verifica se algum indexador foi escolhido e se tem algum valor histórico
      'se tiver calcula o valor indexado 1
      If CboMoeda1.Text <> "" Then
        If TxtValorReal.Text <> "" Then
          XFT_CONTROLE = "DtpDtBase"
          LblValorMoeda12.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
          If Not IsNull(DatMoeda1.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda12 <> "" Then
            LblSimbolo1.Caption = DatMoeda1.Recordset.Fields("moed_tx_simbolo")
          End If
          If LblValorMoeda12.Caption = "" Then
            CboMoeda1.BoundText = ""
            LblSimbolo1.Caption = ""
          End If
        End If
      End If
      
      'Verifica se algum indexador foi escolhido e se tem algum valor histórico
      'se tiver calcula o valor indexado 2
      If CboMoeda2.Text <> "" Then
        If TxtValorReal.Text <> "" Then
          XFT_CONTROLE = "DtpDtBase"
          LblValorMoeda22.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
          If Not IsNull(DatMoeda2.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda22 <> "" Then
            LblSimbolo2.Caption = DatMoeda2.Recordset.Fields("moed_tx_simbolo")
          End If
          If LblValorMoeda22.Caption = "" Then
            CboMoeda2.BoundText = ""
            LblSimbolo2.Caption = ""
          End If
        End If
      End If
    
    XFT_CONTROLE = ""
    End If
  XFB_ALTEROU = False
  End If

End Sub

Private Sub DtpVencimento_Change()

  DtpDtProrrogacao.Value = DtpVencimento.Value
  
End Sub

Private Sub Mskconta_LostFocus()
  If MskConta.Text <> "" Then
    CboPlanoContas.BoundText = MskConta.Text
    If CboPlanoContas.Text = "" Then
      MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
      MskConta.SetFocus
    End If
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

Private Sub CboTiposPlanos_Change()
    If CboTiposPlanos.Text <> "" Then
        DatTipoPlano.Recordset.bookmark = CboTiposPlanos.SelectedItem
'        If Not IsNull(DatTipoPlano.Recordset.Fields("tipl_nr_periodicidade")) Then
'            TxtPeriodicidade.Text = DatTipoPlano.Recordset.Fields("tipl_nr_periodicidade")
'        End If
    End If
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub cmdGravar_Click()
    Dim XLT_SQL As String, XLT_CODTITULO As String
    Dim XLD_VENCIMENTO As Date
    Dim XLT_PCR As String, XLT_CONTA As String
    Dim XLI_X As Integer, XLI_PARCELA As Integer
    Dim XLT_NATUREZA As String
    Dim XLO_PLANO As New ADODB.Recordset
    Dim XLO_CONTRATO As New ADODB.Recordset
    Dim XLF_TOTAL As Double
    Dim XLO_TITULO As ADODB.Recordset
    Dim XLT_STATUSQUALIFIN As String    'Indica a operação que será executada no Qualifin (compromissar ou alterar)
    Dim XLT_EXPORTOU As String
            
    'On Error GoTo RotuloErro
    
    'If CmdGravar.Caption = "&Gravar" Then
        If funChecarCampos Then Exit Sub
        
        If IsNull(DtpDtProrrogacao.Value) Then
          DtpDtProrrogacao.Value = DtpVencimento.Value
        End If
        
        If CboPcr.Text = "" Then
            XLT_PCR = ""
        Else
            XLT_PCR = DatPcr.Recordset.Fields("cere_cd_pcr")
        End If
        
        If CboPlanoContas.Text = "" Then
            XLT_CONTA = ""
        Else
            XLT_CONTA = DatConta.Recordset.Fields("plco_cd_conta")
        End If
        
        Select Case CboNaturezasPlanos.ListIndex
          Case 0
               XLT_NATUREZA = "S"
          Case 1
               XLT_NATUREZA = "P"
          Case 2
               XLT_NATUREZA = "F"
          Case 3
               XLT_NATUREZA = "G"
          Case 4
               XLT_NATUREZA = "O"
          Case 5
               XLT_NATUREZA = "V"
        End Select
        
        If ChkExportou.Value = 1 Then
          XLT_EXPORTOU = "S"
        Else
          XLT_EXPORTOU = "N"
        End If
      
        If XInserir = "I" Then
          SubQOpenRecordset XLO_PLANO, "SELECT CONVERT(decimal(18,2),SUM(titu_vl_Parcela)) as Total From Titulos " & _
          "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
          "AND imov_cd_Imovel='" & CboImovel.BoundText & "' AND cont_cd_Contrato='" & LblContrato2.Caption & "' " & _
          "AND titu_tx_NaturezaPlano='" & XLT_NATUREZA & "' " & _
          "AND titu_cd_Residuo='00'", Estatico
        ElseIf XInserir = "A" Then
          SubQOpenRecordset XLO_PLANO, "SELECT CONVERT(decimal(18,2),SUM(titu_vl_Parcela) - " & _
            " (SELECT titu_vl_parcela FROM conscaptitulos WHERE " & _
            " titulo='" & Chave & "' AND empr_cd_empresa=" & PCodEmpresa & ")) as Total From Titulos " & _
          "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
          "AND imov_cd_Imovel='" & CboImovel.BoundText & "' AND cont_cd_Contrato='" & LblContrato2.Caption & "' " & _
          "AND titu_tx_NaturezaPlano='" & XLT_NATUREZA & "' " & _
          "AND titu_cd_Residuo='00'", Estatico
        End If
        
        SubQOpenRecordset XLO_CONTRATO, "SELECT * From Contratos " & _
           "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
           "AND imov_cd_Imovel='" & CboImovel.BoundText & "' AND cont_cd_Contrato='" & LblContrato2.Caption & "'", Estatico
         
        'XLF_TOTAL = CDbl(FunNuloVal(XLO_PLANO!Total)) + (CInt(TxtParcelas.Text) * CDbl(TxtValorReal.Text))
        XLF_TOTAL = CDbl(FunNuloVal(XLO_PLANO!Total)) + CDbl(TxtValorReal.Text)
        
        Select Case XLT_NATUREZA
          Case "S":
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Sinal Then
              If MsgBox("O valor do Sinal está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Sinal, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
          Case "P"
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Poupanca Then
              If MsgBox("O valor da Poupança está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Poupanca, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
          Case "F"
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Financiamento Then
              If MsgBox("O valor do Financiamento está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Financiamento, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
          Case "G"
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Fgts Then
              If MsgBox("O valor do FGTS está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Fgts, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
          Case "O"
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Outros Then
              If MsgBox("O valor de Outros está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Outros, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
          Case "V"
            If XLF_TOTAL > XLO_CONTRATO!cont_vl_Avista Then
              If MsgBox("O valor A Visata está sendo excedido em  " & Format(XLF_TOTAL - XLO_CONTRATO!cont_vl_Avista, "standard") & "!" & vbCrLf + "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
                Exit Sub
              End If
            End If
        End Select
         
        XLO_PLANO.Close
        Set XLO_PLANO = Nothing
        XLO_CONTRATO.Close
        Set XLO_CONTRATO = Nothing
          
        Conexao.BeginTrans
        
        If XInserir = "I" Then
            XLT_SQL = "SELECT * FROM Titulos " & _
                 "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                       "empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
                       "imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
                       "cont_cd_Contrato='" & LblContrato2 & "' AND " & _
                       "titu_cd_Plano='" & MskPlano.Text & "'" & " AND " & _
                       "titu_cd_Parcela='" & LblParcela2.Caption & "' AND " & _
                       "titu_cd_Residuo='00'"
                       
            SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
        
            XLO_TITULO.AddNew
            XLT_STATUSQUALIFIN = "COMPROMISSAR"
            
            XLO_TITULO!titu_cd_residuo = "00"
            XLO_TITULO!titu_dt_Prorrogacao = Format(DtpVencimento.Value, "dd/MM/yy")
        Else
           
           
            XLT_SQL = "SELECT * FROM Titulos " & _
                 "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                        "empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
                        "imov_cd_imovel='" & XFT_CODIMOVEL & "' AND " & _
                        "cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND " & _
                        "titu_cd_Plano= '" & XFT_CODPLANO & "' AND " & _
                        "titu_cd_Parcela= '" & XFT_CODPARCELA & "' AND " & _
                        "titu_cd_Residuo= '" & XFT_CODRESIDUO & "'"
            SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
            XLT_STATUSQUALIFIN = "ALTERAR"
            XLO_TITULO!titu_cd_residuo = XFT_CODRESIDUO
            XLO_TITULO!titu_dt_Prorrogacao = Format(DtpDtProrrogacao.Value, "dd/MM/yy")
        End If
        
        XLO_TITULO!empr_cd_Empresa = PCodEmpresa
        XLO_TITULO!empd_cd_Empreendimento = MskEmpreendimento.Text
        XLO_TITULO!imov_cd_Imovel = CboImovel.BoundText
        XLO_TITULO!cont_cd_Contrato = LblContrato2.Caption
        XLO_TITULO!titu_cd_plano = MskPlano.Text
        XLO_TITULO!titu_cd_Parcela = LblParcela2.Caption
        XLO_TITULO!obse_cd_Observacao = funNull(CboObs.BoundText)
        XLO_TITULO!moed_cd_Moeda1 = funNull(CboMoeda1.BoundText)
        XLO_TITULO!moed_cd_Moeda2 = funNull(CboMoeda2.BoundText)
        XLO_TITULO!titu_vl_Parcela = TxtValorReal.Text
        XLO_TITULO!titu_vl_ParcelaIndex1 = FunNuloVal(LblValorMoeda12.Caption)
        XLO_TITULO!titu_vl_ParcelaIndex2 = FunNuloVal(LblValorMoeda22.Caption)
        XLO_TITULO!titu_dt_Vencimento = Format(DtpVencimento.Value, "dd/MM/yy")
        XLO_TITULO!tipl_cd_TipoPlano = CboTiposPlanos.BoundText
        XLO_TITULO!titu_vl_PercJurosFin = FunNuloVal(TxtJurosFin.Text)
        XLO_TITULO!titu_vl_PercMultaMora = FunNuloVal(TxtMulta.Text)
        XLO_TITULO!titu_vl_PercJurosMora = FunNuloVal(TxtJuros.Text)
        XLO_TITULO!cere_cd_Pcr = funNull(XLT_PCR)
        XLO_TITULO!plco_cd_Conta = funNull(XLT_CONTA)
        XLO_TITULO!coco_cd_codigo = funNull(CboCCorrente.BoundText)
        XLO_TITULO!titu_dt_Base = Format(DtpDtBase.Value, "dd/MM/yy")
        XLO_TITULO!titu_tx_NaturezaPlano = XLT_NATUREZA
        XLO_TITULO!titu_dt_Criacao = Format(Date, "dd/MM/yy")
        XLO_TITULO!iden_cd_Projeto = funNull(CboIdentificador.BoundText)
        XLO_TITULO!titu_tx_ExpBanco = XLT_EXPORTOU
        XLO_TITULO!titu_nr_ContratoBanco = FunNulo(txtContratoBanco.Text)
        
        
        'Se tiver sido exportado e está cancelando a exportação
        If ChkExportou.Visible = True And ChkExportou.Value = False Then
          XLO_TITULO!titu_vl_Seguro = 0
          XLO_TITULO!moed_cd_Moeda3 = Null
        End If
          
        XLO_TITULO.Update
        
        If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then 'Interligado com Qualifin e exp. compromissos
            subGravaFinanceiro XLO_TITULO, XLT_STATUSQUALIFIN, DatImovel.Recordset.Fields("focl_cd_FornCli"), CboTiposPlanos.Text, CboNaturezasPlanos.Text
        End If
        XLO_TITULO.Close
        
        Set XLO_TITULO = Nothing
'        If XInserir = "I" Then
'            XLT_SQL = "INSERT INTO Titulos (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel,cont_cd_Contrato, " & _
'                        "titu_cd_Plano,titu_cd_Parcela,titu_cd_Residuo,obse_cd_Observacao, " & _
'                        "moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, " & _
'                        "titu_vl_ParcelaIndex1, titu_vl_ParcelaIndex2, titu_dt_Vencimento, tipl_cd_TipoPlano, " & _
'                        "titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
'                        "cere_cd_Pcr,plco_cd_Conta,coco_cd_Codigo,titu_dt_Base," & _
'                        "titu_dt_Prorrogacao,titu_tx_NaturezaPlano,titu_dt_Criacao)" & _
'                  "VALUES (" & PCodEmpresa & ",'" & MskEmpreendimento.Text & "','" & CboImovel.BoundText & "','" & LblContrato2.Caption & "'," & _
'                        "'" & MskPlano.Text & "','" & LblParcela2.Caption & "','00'," & FunNuloBancoVal(CboObs.BoundText) & ", " & _
'                        "" & FunNuloBancoVal(CboMoeda1.BoundText) & "," & FunNuloBancoVal(CboMoeda2.BoundText) & "," & FunNuloVal(FunTrataFloat(TxtValorReal.Text)) & "," & _
'                        "" & FunNuloVal(FunTrataFloat(LblValorMoeda12.Caption)) & "," & FunNuloVal(FunTrataFloat(LblValorMoeda22.Caption)) & "," & FunNuloData(DtpVencimento.Value, NomeSgbd) & "," & CboTiposPlanos.BoundText & "," & _
'                        "" & FunNuloVal(FunTrataFloat(TxtJurosFin.Text)) & "," & FunNuloVal(FunTrataFloat(TxtMulta.Text)) & "," & FunNuloVal(FunTrataFloat(TxtJuros.Text)) & ", " & _
'                        "" & FunNuloBanco(XLT_PCR) & "," & FunNuloBanco(XLT_CONTA) & "," & FunNuloBancoVal(CboCCorrente.BoundText) & ", " & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
'                        "" & FunNuloData(DtpVencimento.Value, NomeSgbd) & ",'" & XLT_NATUREZA & "'," & FunNuloData(Date, NomeSgbd) & ")"
'            Conexao.Execute XLT_SQL
'
'            Chave = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & LblContrato2.Caption & "." & MskPlano.Text & "." & LblParcela2.Caption & "." & "00"
'        Else
'          XLT_SQL = "UPDATE Titulos SET " & _
'              "obse_cd_Observacao=" & FunNuloBancoVal(CboObs.BoundText) & "," & _
'              "titu_vl_Parcela=" & FunNuloVal(FunTrataFloat(TxtValorReal.Text)) & "," & _
'              "moed_cd_Moeda1= " & FunNuloBancoVal(CboMoeda1.BoundText) & "," & _
'              "moed_cd_Moeda2= " & FunNuloBancoVal(CboMoeda2.BoundText) & "," & _
'              "titu_vl_ParcelaIndex1=" & FunNuloVal(FunTrataFloat(LblValorMoeda12.Caption)) & ", " & _
'              "titu_vl_ParcelaIndex2=" & FunNuloVal(FunTrataFloat(LblValorMoeda22.Caption)) & ", " & _
'              "titu_vl_PercJurosFin=" & FunNuloVal(FunTrataFloat(TxtJurosFin.Text)) & "," & _
'              "titu_vl_PercMultaMora=" & FunNuloVal(FunTrataFloat(TxtMulta.Text)) & "," & _
'              "titu_vl_PercJurosMora=" & FunNuloVal(FunTrataFloat(TxtJuros.Text)) & "," & _
'              "cere_cd_Pcr=" & FunNuloBanco(XLT_PCR) & "," & _
'              "plco_cd_Conta= " & FunNuloBanco(XLT_CONTA) & "," & _
'              "coco_cd_Codigo= " & FunNuloBancoVal(CboCCorrente.BoundText) & "," & _
'              "titu_dt_Base=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
'              "titu_dt_Vencimento=" & FunNuloData(DtpVencimento.Value, NomeSgbd) & "," & _
'              "titu_dt_Prorrogacao=" & FunNuloData(DtpDtProrrogacao.Value, NomeSgbd) & "," & _
'              "titu_tx_NaturezaPlano='" & XLT_NATUREZA & "'" & _
'              " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND " & _
'              "empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
'              "imov_cd_imovel='" & XFT_CODIMOVEL & "' AND " & _
'              "cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND " & _
'              "titu_cd_Plano= '" & XFT_CODPLANO & "' AND " & _
'              "titu_cd_Parcela= '" & XFT_CODPARCELA & "' AND " & _
'              "titu_cd_Residuo= '" & XFT_CODRESIDUO & "'"
'          Conexao.Execute XLT_SQL
'        End If
                       
        Conexao.CommitTrans
        
        If XInserir = "I" Then
          Chave = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & LblContrato2.Caption & "." & MskPlano.Text & "." & LblParcela2.Caption & "." & "00"
        ElseIf XInserir = "A" Then
          Chave = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & LblContrato2.Caption & "." & MskPlano.Text & "." & LblParcela2.Caption & "." & XFT_CODRESIDUO
        End If
        
        CmdDesistir_Click
    
'        CmdGravar.Caption = "&Novo"
'    ElseIf CmdGravar.Caption = "&Novo" Then
'        SubPrepararInsercao
'        CmdGravar.Caption = "&Gravar"
'    End If
    Exit Sub
    
RotuloErro:
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFinTitulos
    If XInserir = "I" Then
      'MskPlano.SetFocus
    End If
    FrmFinTitulos.Top = 1005
    FrmFinTitulos.Left = 0
End Sub

Private Sub Form_Load()
    
    Dim XLT_SQL  As String
    
    Set Formulario = FrmFinTitulos
    
    ' Verifica se tem o Qualifin acoplado
    If XGB_PQUALIFIN = False Then
      FraConta.Enabled = False
    End If

   ' CliqueBarraFerramentas = False
    subConectarControleDadosNV DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    subConectarControleDadosNV DatMoeda1, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatMoeda2, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
    subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico

    DtpVencimento.Value = Date
    XFB_ALTEROU = False
        
    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Mid$(Chave, 6, 4)
    XFT_CODCONTRATO = Mid$(Chave, 11, 2)
    XFT_CODPLANO = Mid$(Chave, 14, 2)
    XFT_CODPARCELA = Mid$(Chave, 17, 3)
    XFT_CODRESIDUO = Right$(Chave, 2)
    
    If XInserir = "I" Then
      MskEmpreendimento.Text = XFT_CODEMPREENDIMENTO
      CboEmpreendimento.BoundText = XFT_CODEMPREENDIMENTO
      CboImovel.BoundText = XFT_CODIMOVEL
      LblContrato2.Caption = XFT_CODCONTRATO
      MskPlano.Text = XFT_CODPLANO
      'SubDesabHabCampos (False)
      LblProrrogacao.Visible = False
      DtpDtProrrogacao.Visible = False
      MskEmpreendimento.Enabled = False
      CboEmpreendimento.Enabled = False
      CboImovel.Enabled = False
      'MskPlano.Enabled = False
      'CboTiposPlanos.Enabled = False
      'CboNaturezasPlanos.Enabled = False
      CboIdentificador.Text = ""
    ElseIf XInserir = "A" Then
      
      'Vai para a observação pois se deixa em cima da data base quando
      'sai do foco acab recalculando os valores e se data base tiver sido mudada
      'os valores serão alterados
      'CboObs.SetFocus
      
      Me.Caption = "Alteração de Títulos"
      subPreencheCampos
      MskPlano.Enabled = False
      MskEmpreendimento.Enabled = False
      CboEmpreendimento.Enabled = False
      CboImovel.Enabled = False
      
      If ((UCase(PCliente) = "COSTA ANDRADE") And (func_cd_operador = 39 Or func_cd_operador = 57 Or func_cd_operador = 65)) Then
            CboTiposPlanos.Enabled = True
      Else
            CboTiposPlanos.Enabled = False
      End If
      
      CboNaturezasPlanos.Enabled = False
      XFB_ALTEROU = False
    ElseIf XInserir = "C" Then
      subPreencheCampos
      SubDesabHabCampos (False)
    End If
    
End Sub

Private Sub CboIdentificador_LostFocus()
    If TabAuxiliar.TabEnabled(1) = True Then
        TabAuxiliar.Tab = 1
    End If
End Sub

Private Sub CmdLimparIdentif_Click()
    CboIdentificador.BoundText = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmFinTitulos"
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
        'subConectarControleDados DatImovel, "SELECT * FROM ConsCapImoveis WHERE empd_cd_empreendimento='" & MskEmpreendimento.Text & "' and imov_tx_status='A' ORDER BY imov_cd_imovel", Estatico
    End If
End Sub
Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub

'Private Sub MskPlano_Change()
'  LblParcela2.Caption = FunGeraParcela
'End Sub

Private Sub MskPlano_GotFocus()
  Call subSelecionaMSK(MskPlano)
End Sub

Private Sub MskPlano_LostFocus()
  
  Dim XLO_TITULO As New ADODB.Recordset
  
  MskPlano.Text = Format(MskPlano, "00")
  LblParcela2.Caption = FunGeraParcela
  If LblParcela2.Caption = "" Then
    MsgBox "Esse plano não existe.", vbCritical, "ATENÇÃO !"
    subPrepararInsercao
    'SubDesabHabCampos (False)
    MskPlano.SetFocus
    Exit Sub
  Else
    LblParcela2.Caption = Format(LblParcela2.Caption, "000")
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND titu_cd_Plano='" & XFT_CODPLANO & "' AND titu_cd_Parcela='" & XFT_CODPARCELA & "' AND titu_cd_Residuo='" & XFT_CODRESIDUO & "'", Estatico
        
         Select Case XLO_TITULO!titu_tx_NaturezaPlano
          Case "S"
               CboNaturezasPlanos.ListIndex = 0
          Case "P"
               CboNaturezasPlanos.ListIndex = 1
          Case "F"
               CboNaturezasPlanos.ListIndex = 2
          Case "G"
               CboNaturezasPlanos.ListIndex = 3
          Case "O"
               CboNaturezasPlanos.ListIndex = 4
         End Select
         
         CboTiposPlanos.BoundText = XLO_TITULO!tipl_cd_TipoPlano
        
        XLO_TITULO.Close
        Set XLO_TITULO = Nothing
  End If

End Sub

Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtJuros_LostFocus()
    If IsNumeric(TxtJuros.Text) Then
        TxtJuros.Text = Format$(TxtJuros.Text, "##0.0000")
    Else
        TxtJuros.Text = ""
    End If
End Sub

Private Sub TxtJurosFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtJurosFin_LostFocus()
    If IsNumeric(TxtJurosFin.Text) Then
        TxtJurosFin.Text = Format$(TxtJurosFin.Text, "##0.0000")
    Else
        TxtJurosFin.Text = ""
    End If
End Sub

Private Sub TxtMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMulta_LostFocus()
    If IsNumeric(TxtMulta.Text) Then
        TxtMulta.Text = Format$(TxtMulta.Text, "##0.0000")
    Else
        TxtMulta.Text = ""
    End If
End Sub

Private Sub TxtValorReal_Change()
  
  XFT_CONTROLE = ""
  XFB_ALTEROU = True
  
End Sub

Private Sub TxtValorReal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValorReal_LostFocus()
    
  If XFB_ALTEROU = True Then
    
    If XFT_CONTROLE <> "DtpDtBase" Then
      If IsNumeric(TxtValorReal.Text) Then
          TxtValorReal.Text = Format$(TxtValorReal.Text, "##,##0.00")
          
          'Preenche valor indexado 1
          If CboMoeda1.BoundText <> "" Then
            XFT_CONTROLE = "TxtValorReal"
            LblValorMoeda12.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
            If Not IsNull(DatMoeda1.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda12 <> "" Then
              LblSimbolo1.Caption = DatMoeda1.Recordset.Fields("moed_tx_simbolo")
            End If
            If LblValorMoeda12.Caption = "" Then
              CboMoeda1.BoundText = ""
              LblSimbolo1.Caption = ""
            End If
          End If
          
          'Preenche valor indexado 2
          If CboMoeda2.BoundText <> "" Then
            XFT_CONTROLE = "TxtValorReal"
            LblValorMoeda22.Caption = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XFT_PRORATA, TelaCotacoes)
            If Not IsNull(DatMoeda2.Recordset.Fields("moed_tx_simbolo")) And LblValorMoeda22 <> "" Then
              LblSimbolo2.Caption = DatMoeda2.Recordset.Fields("moed_tx_simbolo")
            End If
            If LblValorMoeda22.Caption = "" Then
              CboMoeda2.BoundText = ""
              LblSimbolo2.Caption = ""
            End If
          End If
      Else
          'Limpa os campos quando o valor histórico não é numerico
          TxtValorReal.Text = ""
          LblValorMoeda12.Caption = ""
          LblSimbolo1.Caption = ""
          LblValorMoeda22.Caption = ""
          LblSimbolo2.Caption = ""
      End If
    
    XFT_CONTROLE = ""
    End If
  XFB_ALTEROU = False
  End If
    
End Sub


