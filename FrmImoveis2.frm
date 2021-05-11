VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManImoveis 
   Caption         =   "Cadastro de Imóveis"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   1350
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   9420
   Begin VB.Frame FraPrincipal 
      Height          =   5010
      Left            =   0
      TabIndex        =   41
      Top             =   -60
      Width           =   9345
      Begin Threed.SSPanel PanDadosGerais 
         Height          =   1095
         Left            =   90
         TabIndex        =   42
         Top             =   795
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1931
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
         Begin MSDataListLib.DataCombo CboTipoImovel 
            Bindings        =   "FrmImoveis.frx":0000
            Height          =   315
            Left            =   3060
            TabIndex        =   6
            Top             =   615
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "tiim_tx_descricao"
            BoundColumn     =   "tiim_cd_TipoImovel"
            Text            =   "CboTipoImovel"
         End
         Begin VB.TextBox TxtAndar 
            Height          =   315
            Left            =   795
            TabIndex        =   5
            Top             =   600
            Width           =   825
         End
         Begin VB.ComboBox CboStatus 
            Height          =   315
            ItemData        =   "FrmImoveis.frx":001C
            Left            =   795
            List            =   "FrmImoveis.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   165
            Width           =   2040
         End
         Begin VB.TextBox TxtGaragem 
            Height          =   315
            Left            =   6135
            TabIndex        =   4
            Top             =   195
            Width           =   2895
         End
         Begin VB.TextBox TxtArea 
            Height          =   315
            Left            =   3705
            TabIndex        =   3
            Top             =   180
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker DtpEntChaves 
            Height          =   315
            Left            =   7575
            TabIndex        =   7
            Top             =   615
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   52232193
            CurrentDate     =   37180
         End
         Begin MSAdodcLib.Adodc DatTipoImovel 
            Height          =   330
            Left            =   3810
            Top             =   660
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4101
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
            Caption         =   "DatTipoImovel"
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
         Begin VB.Label LblAndar 
            Alignment       =   1  'Right Justify
            Caption         =   "Andar:"
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
            Left            =   60
            TabIndex        =   48
            Top             =   630
            Width           =   690
         End
         Begin VB.Label LblTipoImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo do Imóvel:"
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
            Left            =   1650
            TabIndex        =   47
            Top             =   660
            Width           =   1365
         End
         Begin VB.Label LblStatus 
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
            Height          =   210
            Left            =   15
            TabIndex        =   46
            Top             =   210
            Width           =   750
         End
         Begin VB.Label LblEntChaves 
            Alignment       =   1  'Right Justify
            Caption         =   "Ent. Chaves:"
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
            Left            =   6390
            TabIndex        =   45
            Top             =   645
            Width           =   1140
         End
         Begin VB.Label LblGaragem 
            Alignment       =   1  'Right Justify
            Caption         =   "Garagem:"
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
            Left            =   5205
            TabIndex        =   44
            Top             =   225
            Width           =   885
         End
         Begin VB.Label LblArea 
            Alignment       =   1  'Right Justify
            Caption         =   "Área:"
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
            Left            =   3000
            TabIndex        =   43
            Top             =   210
            Width           =   675
         End
      End
      Begin Threed.SSPanel PanCodigo 
         Height          =   510
         Left            =   90
         TabIndex        =   49
         Top             =   225
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   900
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
         Begin VB.ComboBox CboImovel 
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox TxtEmpreendimento 
            Height          =   315
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   0
            Top             =   105
            Width           =   675
         End
         Begin MSDataListLib.DataCombo CboEmpreendimento 
            Bindings        =   "FrmImoveis.frx":003A
            Height          =   315
            Left            =   2535
            TabIndex        =   1
            Top             =   105
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "empd_tx_nome"
            BoundColumn     =   "empd_cd_empreendimento"
            Text            =   "CboEmpreendimento"
         End
         Begin MSAdodcLib.Adodc DatEmpreendimento 
            Height          =   330
            Left            =   3075
            Top             =   180
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
            Left            =   255
            TabIndex        =   51
            Top             =   150
            Width           =   1485
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
            Left            =   6405
            TabIndex        =   50
            Top             =   150
            Width           =   1605
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   2925
         Left            =   60
         TabIndex        =   52
         Top             =   1980
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   5159
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Dados da Venda"
         TabPicture(0)   =   "FrmImoveis.frx":005A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FraValores"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FraDesagio"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FraJuros"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "FraSeguro"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FraProRata"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Compradores"
         TabPicture(1)   =   "FrmImoveis.frx":0076
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LblCliTitular"
         Tab(1).Control(1)=   "LblCliente"
         Tab(1).Control(2)=   "CboComprador"
         Tab(1).Control(3)=   "TDBGrid1"
         Tab(1).Control(4)=   "DatComprador"
         Tab(1).Control(5)=   "CmdExcComprador"
         Tab(1).Control(6)=   "CmdInsComprador"
         Tab(1).Control(7)=   "CmdTitular"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Fiadores"
         TabPicture(2)   =   "FrmImoveis.frx":0092
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label1"
         Tab(2).Control(1)=   "DatFiador"
         Tab(2).Control(2)=   "CboFiador"
         Tab(2).Control(3)=   "TDBGrid3"
         Tab(2).Control(4)=   "CmdInsFiador"
         Tab(2).Control(5)=   "CmdExcFiador"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Corretores"
         TabPicture(3)   =   "FrmImoveis.frx":00AE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TDBGrid2"
         Tab(3).Control(1)=   "PanCorretor"
         Tab(3).Control(2)=   "CmdInsCorretor"
         Tab(3).Control(3)=   "CmdExcCorretor"
         Tab(3).Control(4)=   "CmdAltCorretor"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Dados da Venda 2"
         TabPicture(4)   =   "FrmImoveis.frx":00CA
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "LblIndexador"
         Tab(4).Control(1)=   "LblContratoOrigem"
         Tab(4).Control(2)=   "DatIndexador"
         Tab(4).Control(3)=   "CboIndexador"
         Tab(4).Control(4)=   "FraDatas"
         Tab(4).Control(5)=   "FraFinContabil"
         Tab(4).Control(6)=   "TxtContratoOrigem"
         Tab(4).ControlCount=   7
         Begin VB.TextBox TxtContratoOrigem 
            Height          =   315
            Left            =   -73380
            TabIndex        =   96
            Top             =   2400
            Width           =   975
         End
         Begin VB.Frame FraFinContabil 
            Height          =   1005
            Left            =   -74940
            TabIndex        =   84
            Top             =   1020
            Width           =   9045
            Begin MSMask.MaskEdBox Mskconta 
               Height          =   330
               Left            =   1395
               TabIndex        =   85
               Top             =   585
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas 
               Bindings        =   "FrmImoveis.frx":00E6
               Height          =   330
               Left            =   2340
               TabIndex        =   86
               Top             =   585
               Width           =   6615
               _ExtentX        =   11668
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
            Begin MSMask.MaskEdBox MskPcr 
               Height          =   330
               Left            =   1395
               TabIndex        =   87
               Top             =   180
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmImoveis.frx":00FD
               Height          =   330
               Left            =   2340
               TabIndex        =   88
               Top             =   180
               Width           =   6615
               _ExtentX        =   11668
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
               Left            =   60
               TabIndex        =   90
               Top             =   630
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
               Height          =   315
               Left            =   120
               TabIndex        =   89
               Top             =   225
               Width           =   1260
            End
         End
         Begin VB.Frame FraDatas 
            Height          =   555
            Left            =   -74940
            TabIndex        =   75
            Top             =   420
            Width           =   9045
            Begin MSComCtl2.DTPicker DtpDtBase 
               Height          =   315
               Left            =   1155
               TabIndex        =   76
               Top             =   165
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   52232193
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtDistrato 
               Height          =   315
               Left            =   7515
               TabIndex        =   77
               Top             =   165
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   52232193
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpRegContrato 
               Height          =   315
               Left            =   4380
               TabIndex        =   78
               Top             =   165
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   52232193
               CurrentDate     =   37180
            End
            Begin VB.Label LblRegContrato 
               Alignment       =   1  'Right Justify
               Caption         =   "Registro Contrato:"
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
               Left            =   2730
               TabIndex        =   81
               Top             =   210
               Width           =   1620
            End
            Begin VB.Label LblDtDistrato 
               Alignment       =   1  'Right Justify
               Caption         =   "Data de Distrato:"
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
               Left            =   5970
               TabIndex        =   80
               Top             =   210
               Width           =   1515
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
               Left            =   60
               TabIndex        =   79
               Top             =   210
               Width           =   1035
            End
         End
         Begin VB.Frame FraProRata 
            Caption         =   "Multa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   4560
            TabIndex        =   57
            Top             =   1440
            Width           =   2595
            Begin VB.TextBox TxtPercJurosMulta 
               Height          =   285
               Left            =   1140
               TabIndex        =   92
               Top             =   780
               Width           =   660
            End
            Begin VB.CheckBox ChkMultaProRata 
               Caption         =   "Multa Pró-Rata"
               Height          =   210
               Left            =   180
               TabIndex        =   14
               Top             =   240
               Width           =   1395
            End
            Begin VB.CheckBox ChkCorrecProRata 
               Caption         =   "Correção Monetária Pró-Rata"
               Height          =   210
               Left            =   180
               TabIndex        =   15
               Top             =   540
               Width           =   2370
            End
            Begin VB.Label LblPercJurosMulta 
               Alignment       =   1  'Right Justify
               Caption         =   "Taxa Multa:"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   840
               Width           =   915
            End
         End
         Begin VB.Frame FraSeguro 
            Caption         =   "Seguro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   7260
            TabIndex        =   55
            Top             =   1440
            Width           =   1905
            Begin VB.CheckBox ChkParcela 
               Caption         =   "Pela Parcela"
               Height          =   240
               Left            =   60
               TabIndex        =   23
               Top             =   900
               Width           =   1665
            End
            Begin VB.CheckBox ChkSlDevedor 
               Caption         =   "Pelo Saldo Devedor"
               Height          =   240
               Left            =   60
               TabIndex        =   22
               Top             =   600
               Width           =   1740
            End
            Begin VB.TextBox TxtTaxaSeguro 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1125
               MaxLength       =   20
               TabIndex        =   21
               Top             =   240
               Width           =   660
            End
            Begin VB.Label LblTaxaSeguro 
               Alignment       =   1  'Right Justify
               Caption         =   "Taxa Seguro:"
               Height          =   195
               Left            =   75
               TabIndex        =   56
               Top             =   300
               Width           =   960
            End
         End
         Begin VB.Frame FraJuros 
            Caption         =   "Juros Mora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   2640
            TabIndex        =   54
            Top             =   1440
            Width           =   1845
            Begin VB.TextBox TxtPercJurosMora 
               Height          =   285
               Left            =   1080
               TabIndex        =   91
               Top             =   840
               Width           =   660
            End
            Begin VB.OptionButton OptJurCheio 
               Caption         =   "Cheio"
               Height          =   255
               Left            =   165
               TabIndex        =   20
               Top             =   630
               Width           =   990
            End
            Begin VB.OptionButton OptJurMensal 
               Caption         =   "Mensal"
               Height          =   255
               Left            =   165
               TabIndex        =   19
               Top             =   420
               Width           =   990
            End
            Begin VB.OptionButton OptJurDiario 
               Caption         =   "Diário"
               Height          =   255
               Left            =   165
               TabIndex        =   18
               Top             =   210
               Width           =   990
            End
            Begin VB.Label LblPercJurosMOra 
               Alignment       =   1  'Right Justify
               Caption         =   "Juros Mora:"
               Height          =   255
               Left            =   180
               TabIndex        =   94
               Top             =   900
               Width           =   855
            End
         End
         Begin VB.Frame FraDesagio 
            Caption         =   "Deságio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   2445
            Begin VB.TextBox TxtPercJurosFin 
               Height          =   285
               Left            =   1680
               TabIndex        =   98
               Top             =   780
               Width           =   660
            End
            Begin VB.OptionButton OptDesDiario 
               Caption         =   "Diário"
               Height          =   255
               Left            =   180
               TabIndex        =   16
               Top             =   420
               Width           =   750
            End
            Begin VB.OptionButton OptDesMensal 
               Caption         =   "Mensal"
               Height          =   255
               Left            =   1260
               TabIndex        =   17
               Top             =   420
               Width           =   990
            End
            Begin VB.Label LblPercJurosFin 
               Caption         =   "Juros Financiamento:"
               Height          =   315
               Left            =   120
               TabIndex        =   97
               Top             =   840
               Width           =   1515
            End
         End
         Begin VB.CommandButton CmdExcFiador 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -67020
            TabIndex        =   30
            Top             =   1005
            Width           =   975
         End
         Begin VB.CommandButton CmdInsFiador 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -67020
            TabIndex        =   29
            Top             =   630
            Width           =   975
         End
         Begin VB.CommandButton CmdAltCorretor 
            Caption         =   "&Alterar"
            Height          =   255
            Left            =   -66990
            TabIndex        =   32
            Top             =   1005
            Width           =   975
         End
         Begin VB.CommandButton CmdExcCorretor 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -66990
            TabIndex        =   33
            Top             =   1350
            Width           =   975
         End
         Begin VB.CommandButton CmdInsCorretor 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -66990
            TabIndex        =   31
            Top             =   660
            Width           =   975
         End
         Begin VB.Frame FraValores 
            Height          =   930
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   8985
            Begin VB.TextBox TxtSinal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1260
               MaxLength       =   20
               TabIndex        =   9
               Top             =   555
               Width           =   1215
            End
            Begin VB.TextBox TxtPoupanca 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4380
               MaxLength       =   20
               TabIndex        =   10
               Top             =   165
               Width           =   1215
            End
            Begin VB.TextBox TxtTotVenda 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7500
               MaxLength       =   20
               TabIndex        =   12
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox TxtFgts 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7500
               MaxLength       =   20
               TabIndex        =   13
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtFinanciamento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4395
               MaxLength       =   20
               TabIndex        =   11
               Top             =   525
               Width           =   1215
            End
            Begin VB.TextBox TxtAVista 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1260
               MaxLength       =   20
               TabIndex        =   8
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label LblSinal 
               Alignment       =   1  'Right Justify
               Caption         =   "Sinal:"
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
               Left            =   540
               TabIndex        =   66
               Top             =   615
               Width           =   705
            End
            Begin VB.Label LblPoupanca 
               Alignment       =   1  'Right Justify
               Caption         =   "Poupança:"
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
               Left            =   3405
               TabIndex        =   65
               Top             =   195
               Width           =   960
            End
            Begin VB.Label LblTotVenda 
               Alignment       =   1  'Right Justify
               Caption         =   "Total a Prazo:"
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
               Left            =   6240
               TabIndex        =   64
               Top             =   240
               Width           =   1245
            End
            Begin VB.Label LblFgts 
               Alignment       =   1  'Right Justify
               Caption         =   "FGTS:"
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
               Left            =   6780
               TabIndex        =   63
               Top             =   600
               Width           =   705
            End
            Begin VB.Label LblFinanciamento 
               Alignment       =   1  'Right Justify
               Caption         =   "Financiamento:"
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
               Left            =   3075
               TabIndex        =   62
               Top             =   555
               Width           =   1305
            End
            Begin VB.Label LblAVista 
               Alignment       =   1  'Right Justify
               Caption         =   "À Vista:"
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
               Left            =   540
               TabIndex        =   61
               Top             =   195
               Width           =   705
            End
         End
         Begin VB.CommandButton CmdTitular 
            Caption         =   "&Titular"
            Height          =   255
            Left            =   -67020
            TabIndex        =   27
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton CmdInsComprador 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -67020
            TabIndex        =   25
            Top             =   570
            Width           =   975
         End
         Begin VB.CommandButton CmdExcComprador 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -67020
            TabIndex        =   26
            Top             =   945
            Width           =   975
         End
         Begin MSAdodcLib.Adodc DatComprador 
            Height          =   330
            Left            =   -71460
            Top             =   525
            Visible         =   0   'False
            Width           =   2100
            _ExtentX        =   3704
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
            Caption         =   "DatComprador"
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
         Begin TrueDBGrid70.TDBGrid TDBGrid1 
            Height          =   1650
            Left            =   -74820
            TabIndex        =   67
            Top             =   945
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   2910
            _LayoutType     =   4
            _RowHeight      =   12
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nome"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Titular"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1048848"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=1057296"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
            Splits(0)._ColumnProps(23)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(28)=   "Column(2)._ColStyle=1048848"
            Splits(0)._ColumnProps(29)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(30)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(31)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(32)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(33)=   "Column(2).Order=3"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&"
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
            _StyleDefs(18)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(35)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=16,.parent=21"
            _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=22"
            _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=75"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0"
            _StyleDefs(45)  =   "Named:id=34:Heading"
            _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   ":id=34,.wraptext=-1"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(52)  =   "Named:id=37:Caption"
            _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(54)  =   "Named:id=38:HighlightRow"
            _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(56)  =   "Named:id=39:EvenRow"
            _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(58)  =   "Named:id=40:OddRow"
            _StyleDefs(59)  =   ":id=40,.parent=33"
            _StyleDefs(60)  =   "Named:id=41:RecordSelector"
            _StyleDefs(61)  =   ":id=41,.parent=34"
            _StyleDefs(62)  =   "Named:id=42:FilterBar"
            _StyleDefs(63)  =   ":id=42,.parent=33"
         End
         Begin MSDataListLib.DataCombo CboComprador 
            Bindings        =   "FrmImoveis.frx":0112
            Height          =   315
            Left            =   -72810
            TabIndex        =   24
            Top             =   510
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "focl_tx_RazaoSocial"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   "CboComprador"
         End
         Begin Threed.SSPanel PanCorretor 
            Height          =   1500
            Left            =   -73500
            TabIndex        =   68
            Top             =   705
            Visible         =   0   'False
            Width           =   5340
            _Version        =   65536
            _ExtentX        =   9419
            _ExtentY        =   2646
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
            Begin VB.CommandButton CmdSairCorretor 
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
               Height          =   285
               Left            =   4230
               TabIndex        =   38
               Top             =   1095
               Width           =   975
            End
            Begin VB.CommandButton CmdConfirmaCorretor 
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
               Height          =   285
               Left            =   3195
               TabIndex        =   37
               Top             =   1095
               Width           =   975
            End
            Begin VB.TextBox TxtComissao 
               Height          =   315
               Left            =   1065
               MaxLength       =   20
               TabIndex        =   35
               Top             =   600
               Width           =   870
            End
            Begin MSDataListLib.DataCombo CboCorretor 
               Bindings        =   "FrmImoveis.frx":012D
               Height          =   315
               Left            =   1065
               TabIndex        =   34
               Top             =   165
               Width           =   4170
               _ExtentX        =   7355
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   "corr_tx_nome"
               BoundColumn     =   "corr_cd_corretor"
               Text            =   "CboCorretor"
            End
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3690
               TabIndex        =   36
               Top             =   600
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   52232193
               CurrentDate     =   37180
            End
            Begin MSAdodcLib.Adodc DatCorretor 
               Height          =   330
               Left            =   2295
               Top             =   195
               Visible         =   0   'False
               Width           =   2130
               _ExtentX        =   3757
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
               Caption         =   "DatCorretor"
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
            Begin VB.Label LblPagamento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Height          =   255
               Left            =   2610
               TabIndex        =   71
               Top             =   630
               Width           =   1035
            End
            Begin VB.Label LblComissao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comissão:"
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
               TabIndex        =   70
               Top             =   630
               Width           =   825
            End
            Begin VB.Label LblNomeCorretor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nome:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   405
               TabIndex        =   69
               Top             =   225
               Width           =   600
            End
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid2 
            Height          =   1935
            Left            =   -74820
            TabIndex        =   72
            Top             =   675
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   3413
            _LayoutType     =   4
            _RowHeight      =   12
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nome"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Comissão"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Pagamento"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "dd/mm/yy"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1048848"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=1057296"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
            Splits(0)._ColumnProps(22)=   "Column(2).Width=2143"
            Splits(0)._ColumnProps(23)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(2)._WidthInPix=2064"
            Splits(0)._ColumnProps(25)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(2)._ColStyle=1057296"
            Splits(0)._ColumnProps(27)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(28)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(30)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(31)=   "Column(2)._MinWidth=89769656"
            Splits(0)._ColumnProps(32)=   "Column(3).Width=2752"
            Splits(0)._ColumnProps(33)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(3)._WidthInPix=2672"
            Splits(0)._ColumnProps(35)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(3)._ColStyle=1057296"
            Splits(0)._ColumnProps(37)=   "Column(3).WrapText=1"
            Splits(0)._ColumnProps(38)=   "Column(3).FetchStyle=1"
            Splits(0)._ColumnProps(39)=   "Column(3).AllowFocus=0"
            Splits(0)._ColumnProps(40)=   "Column(3).Order=4"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&"
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
            _StyleDefs(18)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(35)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=90,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(40)  =   ":id=90,.wraptext=-1,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=22,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=23"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=75"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=94,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(45)  =   ":id=94,.wraptext=-1,.locked=-1"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=22,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=23"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=75"
            _StyleDefs(49)  =   "Named:id=33:Normal"
            _StyleDefs(50)  =   ":id=33,.parent=0"
            _StyleDefs(51)  =   "Named:id=34:Heading"
            _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   ":id=34,.wraptext=-1"
            _StyleDefs(54)  =   "Named:id=35:Footing"
            _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid3 
            Height          =   1650
            Left            =   -74820
            TabIndex        =   73
            Top             =   1005
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   2910
            _LayoutType     =   4
            _RowHeight      =   12
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nome"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1048848"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=1057296"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&"
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
            _StyleDefs(18)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(35)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(39)  =   "Named:id=33:Normal"
            _StyleDefs(40)  =   ":id=33,.parent=0"
            _StyleDefs(41)  =   "Named:id=34:Heading"
            _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(43)  =   ":id=34,.wraptext=-1"
            _StyleDefs(44)  =   "Named:id=35:Footing"
            _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(46)  =   "Named:id=36:Selected"
            _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(48)  =   "Named:id=37:Caption"
            _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(50)  =   "Named:id=38:HighlightRow"
            _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(52)  =   "Named:id=39:EvenRow"
            _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(54)  =   "Named:id=40:OddRow"
            _StyleDefs(55)  =   ":id=40,.parent=33"
            _StyleDefs(56)  =   "Named:id=41:RecordSelector"
            _StyleDefs(57)  =   ":id=41,.parent=34"
            _StyleDefs(58)  =   "Named:id=42:FilterBar"
            _StyleDefs(59)  =   ":id=42,.parent=33"
         End
         Begin MSDataListLib.DataCombo CboFiador 
            Bindings        =   "FrmImoveis.frx":0147
            Height          =   315
            Left            =   -72810
            TabIndex        =   28
            Top             =   570
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "fiad_tx_nome"
            BoundColumn     =   "fiad_cd_fiador"
            Text            =   "CboFiador"
         End
         Begin MSAdodcLib.Adodc DatFiador 
            Height          =   330
            Left            =   -72300
            Top             =   600
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
            Caption         =   "DatFiador"
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
         Begin MSDataListLib.DataCombo CboIndexador 
            Bindings        =   "FrmImoveis.frx":015F
            Height          =   315
            Left            =   -73920
            TabIndex        =   82
            Top             =   2040
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "moed_tx_descricao"
            BoundColumn     =   "moed_cd_codmoeda"
            Text            =   "CboIndexador"
         End
         Begin MSAdodcLib.Adodc DatIndexador 
            Height          =   330
            Left            =   -71040
            Top             =   2040
            Visible         =   0   'False
            Width           =   1710
            _ExtentX        =   3016
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
            Caption         =   "DatIndexador"
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
         Begin VB.Label LblContratoOrigem 
            Caption         =   "Contrato Origem:"
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
            Left            =   -74880
            TabIndex        =   95
            Top             =   2460
            Width           =   1455
         End
         Begin VB.Label LblIndexador 
            Alignment       =   1  'Right Justify
            Caption         =   "Indexador:"
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
            Left            =   -74940
            TabIndex        =   83
            Top             =   2100
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Selecione o Fiador:"
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
            Left            =   -74850
            TabIndex        =   74
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label LblCliente 
            Alignment       =   1  'Right Justify
            Caption         =   "Selecione o Cliente:"
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
            Left            =   -74850
            TabIndex        =   59
            Top             =   540
            Width           =   1995
         End
         Begin VB.Label LblCliTitular 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente Titular"
            ForeColor       =   &H00008000&
            Height          =   225
            Left            =   -67035
            TabIndex        =   58
            Top             =   2400
            Width           =   1005
         End
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5580
      TabIndex        =   40
      Top             =   4980
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
      Left            =   2400
      TabIndex        =   39
      Top             =   4980
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
Attribute VB_Name = "FrmManImoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VFV_VETOR1 As New XArray 'Vetor de Compradores
Dim VFV_VETOR2 As New XArray 'Vetor de Corretores
Dim VFV_VETOR3 As New XArray 'Vetor de Fiadores
Dim XFB_ENTROUCOMP As Boolean 'Controla a entrada na orelha de Compradores
Dim XFB_ENTROUFIAD As Boolean 'Controla a entrada na orelha de Fiadores
Dim XFB_ENTROUCORR As Boolean 'Controla a entrada na orelha de Corretores
Dim XFI_CODCLIENTE As Integer, XFI_CODCORRETOR As Integer
Dim XFT_CODIMOVEL As String, XFT_INSCORRETOR As String


Dim errLoop As ADODB.Error
Dim StrError As String

Sub SubDesabilitaBotoes()
    If TabAuxiliar.Tab = 1 Then
        CmdInsComprador.Enabled = False
        CmdExcComprador.Enabled = False
    ElseIf TabAuxiliar.Tab = 2 Then
        CmdInsFiador.Enabled = False
        CmdExcFiador.Enabled = False
    ElseIf TabAuxiliar.Tab = 3 Then
        CmdInsCorretor.Enabled = False
        CmdAltCorretor.Enabled = False
        CmdExcCorretor.Enabled = False
    End If
End Sub
                                                        
Sub SubHabilitaBotoes()
    If TabAuxiliar.Tab = 1 Then
        CmdInsComprador.Enabled = True
        CmdExcComprador.Enabled = True
    ElseIf TabAuxiliar.Tab = 2 Then
        CmdInsFiador.Enabled = True
        CmdExcFiador.Enabled = True
    ElseIf TabAuxiliar.Tab = 3 Then
        CmdInsCorretor.Enabled = True
        CmdAltCorretor.Enabled = True
        CmdExcCorretor.Enabled = True
    End If
End Sub
Sub SubPreencheCampos()
    Dim XLT_SQL As String
    
    'Dados Gerais
    XFI_CODCLIENTE = TelaManImoveis.Adodc1.Recordset.Fields("focl_cd_forncli")
    TxtEmpreendimento.Text = TelaManImoveis.Adodc1.Recordset.Fields("empd_cd_empreendimento")
    CboEmpreendimento.BoundText = TelaManImoveis.Adodc1.Recordset.Fields("empd_cd_empreendimento")
    'MskCodigo.Mask = ""
    CboImovel.Text = TelaManImoveis.Adodc1.Recordset.Fields("imov_cd_imovel")
    XFT_CODIMOVEL = CboImovel.Text
    'MskCodigo.Mask = "####-#"
    LblEmpreendimento.Enabled = False
    TxtEmpreendimento.Enabled = False
    CboEmpreendimento.Enabled = False
    LblCodigo.Enabled = False
    CboImovel.Enabled = False
    PanCodigo.Enabled = False
    TabAuxiliar.Enabled = True
    Select Case TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_status")
        Case "D"
            CboStatus.ListIndex = 0
            TabAuxiliar.Enabled = False
        Case "A"
            CboStatus.ListIndex = 1
        Case "I"
            CboStatus.ListIndex = 2
        Case Else
            CboStatus.ListIndex = 3
    End Select
    
    TxtArea.Text = TelaManImoveis.Adodc1.Recordset.Fields("imov_nr_area")
    TxtGaragem.Text = TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_garagem")
    TxtAndar.Text = TelaManImoveis.Adodc1.Recordset.Fields("imov_nr_andar")
    CboTipoImovel.BoundText = TelaManImoveis.Adodc1.Recordset.Fields("tiim_cd_tipoimovel")
    CboIndexador.BoundText = TelaManImoveis.Adodc1.Recordset.Fields("moed_cd_CodMoeda")
    
    If Not IsNull(TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_EntregaChaves")) Then DtpEntChaves.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_EntregaChaves")
    If Not IsNull(TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_Base")) Then DtpDtBase.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_Base")
    If Not IsNull(TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_RegistroContrato")) Then DtpRegContrato.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_RegistroContrato")
    If Not IsNull(TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_Distrato")) Then DtpDtDistrato.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_dt_Distrato")
    
    'Dados da Venda
    TxtAVista.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_AVista"), "standard")
    TxtSinal.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_Sinal"), "standard")
    TxtPoupanca.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_Poupanca"), "standard")
    TxtFinanciamento.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_Financiamento"), "standard")
    TxtTotVenda.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_Total"), "standard")
    TxtFgts.Text = Format(TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_Fgts"), "standard")
    
    ChkMultaProRata.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_MultaProRata")
    ChkCorrecProRata.Value = TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_CorrecProRata")
    Select Case TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_TipoJurosMora")
        Case "D"
            OptJurDiario.Value = True
        Case "M"
            OptJurMensal.Value = True
        Case "C"
            OptJurCheio.Value = True
        Case Else
            OptJurDiario.Value = False
            OptJurMensal.Value = False
            OptJurCheio.Value = False
    End Select
    
    Select Case TelaManImoveis.Adodc1.Recordset.Fields("imov_tx_Desagio")
        Case "D"
            OptDesDiario.Value = True
        Case "M"
            OptDesMensal.Value = True
        Case Else
            OptDesDiario.Value = False
            OptDesMensal.Value = False
    End Select
    
    TxtTaxaSeguro.Text = TelaManImoveis.Adodc1.Recordset.Fields("imov_vl_TaxaSegVida")
    If TelaManImoveis.Adodc1.Recordset.Fields("imov_nr_TipoSeguro") = 1 Then
        ChkSlDevedor.Value = 1
    ElseIf TelaManImoveis.Adodc1.Recordset.Fields("imov_nr_TipoSeguro") = 2 Then
        ChkParcela.Value = 1
    End If
End Sub


Sub SubPrepararInsercao()
    TxtEmpreendimento.Text = ""
    CboEmpreendimento.Text = ""
    'MskCodigo.Mask = ""
    CboImovel.Text = ""
    'MskCodigo.Mask = "####-#"
    CboImovel.Enabled = False
    CboStatus.ListIndex = 0
    TxtArea.Text = ""
    TxtGaragem.Text = ""
    TxtAndar.Text = ""
    CboTipoImovel.Text = ""
    CboIndexador.Text = ""
    
    'TabAuxiliar.Tab(0)
    TxtAVista.Text = ""
    TxtSinal.Text = ""
    TxtPoupanca.Text = ""
    TxtFinanciamento = ""
    TxtTotVenda.Text = ""
    TxtFgts.Text = ""
    
    ChkMultaProRata.Value = 0
    ChkCorrecProRata.Value = 0
    OptDesDiario.Value = True
    OptJurDiario.Value = True
    TxtTaxaSeguro.Text = ""
    ChkSlDevedor.Value = 0
    ChkParcela.Value = 0
    
    DtpEntChaves.Value = ""
    DtpDtBase.Value = ""
    DtpDtDistrato.Value = ""
    DtpRegContrato.Value = ""
        
    'Compradores
    CboComprador.Text = ""
    
    'Fiadores
    CboFiador.Text = ""
End Sub

Private Sub CboComprador_DblClick(Area As Integer)
    FrmFornecedores.Show
    If FunLocalizaJanelasAtivas("FrmFornecedores") <> -1 Then
        FrmFornecedopres.SetFocus
    End If
End Sub

Private Sub CboComprador_GotFocus()
    TabAuxiliar.Tab = 1
End Sub

Private Sub CboCorretor_Change()
    TxtComissao.Text = Format(((CDbl(TxtTotVenda.Text) * DatCorretor.Recordset.Fields("corr_vl_percentcomissao")) / 100), "standard")
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.Bookmark = CboEmpreendimento.SelectedItem
        TxtEmpreendimento.Text = CboEmpreendimento.BoundText
        CboImovel.Enabled = True
        
'        If MskCodigo.Text <> "     - " And MskCodigo.Text <> XFT_CODIMOVEL Then
'            Dim XLO_IMOVEL As Object
'            SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM Imoveis WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento= '" & TxtEmpreendimento.Text & "'", Estatico
'            If Not XLO_IMOVEL.EOF Then
'               MsgBox "Imóvel já cadastrado para este Empreendimento.", vbCritical + vbOKOnly, "ATENÇÃO"
'               MskCodigo.SetFocus
'               XLO_IMOVEL.Close
'               Exit Sub
'            End If
'            XLO_IMOVEL.Close
'        End If
    End If
End Sub

Private Sub CboEmpreendimento_DblClick(Area As Integer)
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub

Private Sub CboFiador_DblClick(Area As Integer)
    TelaManFiadores.Show
    If FunLocalizaJanelasAtivas("TelaManFiadores") <> -1 Then
        TelaManFiadores.SetFocus
    End If
End Sub

Private Sub CboFiador_GotFocus()
    TabAuxiliar.Tab = 2
End Sub

Private Sub CboStatus_LostFocus()
    If CboStatus.ListIndex <> 0 Then
        TabAuxiliar.Enabled = True
    Else
        TabAuxiliar.Enabled = False
    End If
End Sub

Private Sub ChkParcela_Click()
    If ChkParcela.Value = 1 Then ChkSlDevedor.Value = 0
End Sub

Private Sub ChkSlDevedor_Click()
    If ChkSlDevedor.Value = 1 Then ChkParcela.Value = 0
End Sub

Private Sub CmdAltCorretor_Click()
    CboCorretor.BoundText = TDBGrid2.Columns("Codigo")
    XFI_CODCORRETOR = TDBGrid2.Columns("Codigo")
    TxtComissao.Text = TDBGrid2.Columns("Comissão")
    DtpPagamento.Value = TDBGrid2.Columns("Pagamento")
    XFT_INSCORRETOR = "A"
    PanCorretor.Enabled = True
    PanCorretor.Visible = True
End Sub

Private Sub CmdConfirmaCorretor_Click()
    If CboCorretor.Text <> "" Then
        DatCorretor.Recordset.Bookmark = CboCorretor.SelectedItem
        
        If XDistribuicao = "A" And XFI_CORRETOR = CboCorretor.BoundText Then
            TDBGrid2.Columns(0).Text = CboCorretor.BoundText
            TDBGrid2.Columns(1).Text = CboCorretor.Text
            TDBGrid2.Columns(2).Text = TxtComissao.Text
            TDBGrid2.Columns(3).Text = DtpPagamento.Value
            TDBGrid2.Update
        Else
            If funExisteChave(1, Array("0"), Array(DatCorretor.Recordset.Fields("corr_cd_corretor")), VFV_VETOR2, TDBGrid2) Then
                MsgBox "Esse Corretor já foi cadastrado", vbCritical, "ATENÇÃO !"
                Exit Sub
            Else
                If XFT_INSCORRETOR = "I" Then
                    subInsereRegistroVetor Array(CboCorretor.BoundText, CboCorretor.Text, TxtComissao.Text, DtpPagamento.Value), VFV_VETOR2, TDBGrid2
                ElseIf XDistribuicao = "A" Then
                    TDBGrid2.Columns(0).Text = CboCorretor.BoundText
                    TDBGrid2.Columns(1).Text = CboCorretor.Text
                    TDBGrid2.Columns(2).Text = TxtComissao.Text
                    TDBGrid2.Columns(3).Text = DtpPagamento.Value
                    TDBGrid2.Update
                End If
            End If
        End If
        
        If Not CmdExcCorretor.Enabled Then CmdExcCorretor.Enabled = True
        If Not CmdAltCorretor.Enabled Then CmdAltCorretor.Enabled = True
    End If
    
    PanCorretor.Enabled = False
    PanCorretor.Visible = False
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub CmdExcComprador_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR1, TDBGrid1
        If VFV_VETOR1.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR1(0, 0)) Then
                CmdExcComprador.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub CmdExcCorretor_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR2, TDBGrid2
        If VFV_VETOR2.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR2(0, 0)) Then
                CmdExcCorretor.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub CmdExcFiador_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR3, TDBGrid3
        If VFV_VETOR3.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR3(0, 0)) Then
                CmdExcFiador.Enabled = False
            End If
        End If
    End If
End Sub


Private Sub CmdGravar_Click()
    Dim XLT_STATUS As String, XLT_SQL As String
    Dim XLT_JUROS As String, XLT_DESAGIO As String, XLT_SEGURO As String
            
    'On Error GoTo TrataErro
    
    If CmdGravar.Caption = "&Gravar" Then
        'Se a taxa do seguro for digitada é obrigatório optar por 1 tipo de seguro
        If TxtTaxaSeguro.Text <> "" And ChkSlDevedor.Value = 0 And ChkSlDevedor.Value = 0 Then
            MsgBox "É preciso definir o tipo de Seguro.", vbCritical, "ATENÇÃO !"
            TabAuxiliar.Tab = 0
            ChkSlDevedor.SetFocus
            Exit Sub
        End If
        
        If XFI_CODCLIENTE = 0 Then
            MsgBox "É preciso definir um Cliente como Titular.", vbCritical, "ATENÇÃO !"
            TabAuxiliar.Tab = 1
            CmdTitular.SetFocus
            Exit Sub
        End If
        If FunObrigatorioTXT(TxtEmpreendimento, "O Empreendimento é Obrigatório") Then Exit Sub
        If FunObrigatorioMSK(MskCodigo, "O Código do Imóvel é Obrigatório") Then Exit Sub
        If FunObrigatorioCBO(CboTipoImovel, "O Tipo do Imóvel é Obrigatório") Then Exit Sub
            
        Select Case CboStatus.ListIndex
            Case 0
                 XLT_STATUS = "A"
            Case 1
                 XLT_STATUS = "I"
        End Select
                
        If CboStatus.ListIndex <> 0 Then
            If OptJurDiario.Value = True Then
                XLT_JUROS = "D"
            ElseIf OptJurMensal.Value = True Then
                XLT_JUROS = "M"
            Else
                XLT_JUROS = "C"
            End If
            
            If OptDesDiario.Value = True Then
                XLT_DESAGIO = "D"
            Else
                XLT_DESAGIO = "M"
            End If
        Else
            XLT_JUROS = ""
            XLT_DESAGIO = ""
        End If
            
        If ChkSlDevedor.Value = 1 Then
            XLT_SEGURO = 1
        ElseIf ChkParcela.Value = 1 Then
            XLT_SEGURO = 2
        Else
            XLT_SEGURO = ""
        End If
        
        Conexao.BeginTrans
        If XInserir = "I" Then
           
           ' XLT_SQL = "INSERT INTO Imoveis (imov_cd_imovel, empd_cd_empreendimento, imov_nr_area, imov_tx_garagem, imov_nr_andar, " & _
                        "tiim_cd_TipoImovel, moed_cd_CodMoeda, imov_dt_EntregaChaves, imov_dt_Base, " & _
                        "imov_dt_RegistroContrato, imov_dt_Distrato, imov_vl_AVista, imov_vl_Sinal, " & _
                        "imov_vl_Poupanca, imov_vl_Financiamento, imov_vl_Total, imov_vl_Fgts, " & _
                        "imov_tx_TipoJurosMora, imov_tx_Desagio,imov_vl_TaxaSegVida, imov_nr_TipoSeguro, " & _
                        "imov_tx_MultaProRata, imov_tx_CorrecProRata, imov_tx_status, focl_cd_forncli) " & _
                  "VALUES ('" & MskCodigo.Text & "','" & TxtEmpreendimento.Text & "','" & TxtArea.Text & "','" & TxtGaragem.Text & "','" & TxtAndar.Text & "'," & _
                        "" & CboTipoImovel.BoundText & "," & CboIndexador.BoundText & "," & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
                        "" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & ", " & _
                        "" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & ", " & _
                        "'" & FunNulo(XLT_JUROS) & "','" & FunNulo(XLT_DESAGIO) & "', " & FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & ", '" & XLT_SEGURO & "', " & _
                        "'" & ChkMultaProRata.Value & "', '" & ChkCorrecProRata.Value & "','" & XLT_STATUS & "'," & XFI_CODCLIENTE & ")"
            
            XLT_SQL = "INSERT INTO Contratos (cont_cd_Contrato,imov_cd_imovel, empd_cd_empreendimento, cere_cd_Pcr, plco_cd_Conta" & _
                        "moed_cd_CodMoeda, cont_dt_EntregaChaves, cont_dt_Base, " & _
                        "cont_dt_RegistroContrato, cont_dt_Distrato, cont_vl_AVista, cont_vl_Sinal, " & _
                        "cont_vl_Poupanca, cont_vl_Financiamento, cont_vl_Total, cont_vl_Fgts, " & _
                        "cont_tx_TipoJurosMora, cont_tx_Desagio,cont_vl_TaxaSegVida, cont_nr_TipoSeguro, " & _
                        "cont_tx_MultaProRata, cont_tx_CorrecProRata, cont_tx_status, focl_cd_FornCli, " & _
                        "cont_vl_PercJurosFin, cont_vl_PercJurosMulta, cont_vl_PercJurosMora, cont_cd_ContratoOrigem) " & _
                  "VALUES ('" & MskCodigo.Text & "','" & TxtEmpreendimento.Text & "','" & TxtArea.Text & "','" & TxtGaragem.Text & "','" & TxtAndar.Text & "'," & _
                        "" & CboTipoImovel.BoundText & "," & CboIndexador.BoundText & "," & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
                        "" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & ", " & _
                        "" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & ", " & _
                        "'" & FunNulo(XLT_JUROS) & "','" & FunNulo(XLT_DESAGIO) & "', " & FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & ", '" & XLT_SEGURO & "', " & _
                        "'" & ChkMultaProRata.Value & "', '" & ChkCorrecProRata.Value & "','" & XLT_STATUS & "'," & XFI_CODCLIENTE & "" & _
                        "" & TxtPercJurosFin.Text & ", " & TxtPercJurosMulta.Text & "," & TxtPercJurosMora.Text & "," & XFI_CODCLIENTE & ")"
            Conexao.Execute XLT_SQL
        Else    'Alteração
            'XLT_SQL = "UPDATE Imoveis SET imov_nr_area='" & TxtArea.Text & "'," & _
                        "imov_tx_garagem='" & TxtGaragem.Text & "'," & _
                        "imov_nr_andar='" & TxtAndar.Text & "'," & _
                        "tiim_cd_TipoImovel=" & CboTipoImovel.BoundText & "," & _
                        "moed_cd_CodMoeda= " & CboIndexador.BoundText & "," & _
                        "imov_dt_EntregaChaves=" & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & _
                        "imov_dt_Base=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
                        "imov_dt_RegistroContrato=" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & _
                        "imov_dt_Distrato=" & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & "," & _
                        "imov_vl_AVista=" & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & "," & _
                        "imov_vl_Sinal=" & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & "," & _
                        "imov_vl_Poupanca=" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & "," & _
                        "imov_vl_Financiamento= " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & _
                        "imov_vl_Total=" & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & _
                        "imov_vl_Fgts=" & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & ", " & _
                        "imov_tx_TipoJurosMora='" & XLT_JUROS & "'," & _
                        "imov_tx_Desagio='" & XLT_DESAGIO & "'," & _
                        "imov_vl_TaxaSegVida=" & FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & "," & _
                        "imov_nr_TipoSeguro= '" & XLT_SEGURO & "'," & _
                        "imov_tx_MultaProRata='" & ChkMultaProRata.Value & "'," & _
                        "imov_tx_CorrecProRata= '" & ChkCorrecProRata.Value & "'," & _
                        "imov_tx_status='" & XLT_STATUS & "'," & _
                        "focl_cd_forncli=" & XFI_CODCLIENTE & "" & _
                    " WHERE imov_cd_imovel='" & MskCodigo.Text & "' and empd_cd_empreendimento='" & TxtEmpreendimento.Text & "'"
            
            XLT_SQL = "UPDATE Contratos SET cere_cd_Pcr=" & X & ", plco_cd_Conta= '" & X & "'," & _
                        "moed_cd_CodMoeda= " & CboIndexador.BoundText & "," & _
                        "cont_dt_EntregaChaves=" & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & _
                        "cont_dt_Base=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "," & _
                        "cont_dt_RegistroContrato=" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & _
                        "cont_dt_Distrato=" & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & "," & _
                        "cont_vl_AVista=" & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & "," & _
                        "cont_vl_Sinal=" & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & "," & _
                        "cont_vl_Poupanca=" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & "," & _
                        "cont_vl_Financiamento= " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & _
                        "cont_vl_Total=" & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & _
                        "cont_vl_Fgts=" & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & ", " & _
                        "cont_tx_TipoJurosMora='" & XLT_JUROS & "'," & _
                        "cont_tx_Desagio='" & XLT_DESAGIO & "'," & _
                        "cont_vl_TaxaSegVida=" & FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & "," & _
                        "cont_nr_TipoSeguro= '" & XLT_SEGURO & "'," & _
                        "cont_tx_MultaProRata='" & ChkMultaProRata.Value & "'," & _
                        "cont_tx_CorrecProRata= '" & ChkCorrecProRata.Value & "'," & _
                        "cont_tx_status='" & XLT_STATUS & "'," & _
                        "focl_cd_forncli=" & XFI_CODCLIENTE & "" & _
                        "cont_vl_PercJurosFin=" & TxtPercJurosFin.Text & "" & _
                        "cont_vl_PercJurosMulta=" & TxtPercJurosMulta.Text & "" & _
                        "cont_vl_PercJurosMora=" & TxtPercJurosMora.Text & "" & _
                        "cont_cd_ContratoOrigem=" & X & "" & _
                        " WHERE cont_cd_Contrato='" & X & "' AND imov_cd_imovel='" & MskCodigo.Text & "' AND empd_cd_empreendimento='" & TxtEmpreendimento.Text & "'"
            
            Conexao.Execute XLT_SQL
        End If
                  
        'Compradores
        If XFB_ENTROUCOMP Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Compradores WHERE imov_cd_imovel = '" & MskCodigo.Text & "'"
            End If
            If VFV_VETOR1(0, 0) <> Empty Then
                subGravaVetor "Compradores", "focl_cd_FornCli, imov_cd_imovel, empd_cd_empreendimento", Array("N", "T", "T"), 1, Array(0), "'" & MskCodigo.Text & "', '" & TxtEmpreendimento.Text & "'", VFV_VETOR1, TDBGrid1
            End If
        End If
        
        'Fiadores
        If XFB_ENTROUFIAD Then
            If XInserir = "A" Then
                Conexao.Execute "DELETE from FiadoresImoveis WHERE imov_cd_imovel = '" & MskCodigo.Text & "'"
            End If
            If VFV_VETOR3(0, 0) <> Empty Then
                subGravaVetor "FiadoresImoveis", "fiad_cd_Fiador, imov_cd_imovel, empd_cd_empreendimento", Array("N", "T", "T"), 1, Array(0), "'" & MskCodigo.Text & "', '" & TxtEmpreendimento.Text & "'", VFV_VETOR3, TDBGrid3
            End If
        End If
        
        'Corretores
        If XFB_ENTROUCORR Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Comissoes WHERE imov_cd_imovel = '" & MskCodigo.Text & "'"
            End If
            If VFV_VETOR2(0, 0) <> Empty Then
                subGravaVetor "Comissoes", "corr_cd_Corretor, comi_vl_Comissao, comi_dt_Pagamento, imov_cd_imovel, empd_cd_empreendimento", Array("N", "N", "D", "T", "T"), 3, Array(0, 2, 3), "'" & MskCodigo.Text & "', '" & TxtEmpreendimento.Text & "'", VFV_VETOR2, TDBGrid2
            End If
        End If
        Conexao.CommitTrans
        
        CmdGravar.Caption = "&Novo"
        If XInserir = "A" Then
            Call CmdDesistir_Click
        End If
    
    ElseIf CmdGravar.Caption = "&Novo" Then
        Chave = -1
        XInserir = "I"
        TabAuxiliar.Tab = 0
        SubPrepararInsercao

        XFB_ENTROUCOMP = False
        XFB_ENTROUFIAD = False
        XFB_ENTROUCORR = False
        subLimpaVetor VFV_VETOR1, TDBGrid1
        subLimpaVetor VFV_VETOR2, TDBGrid2
        subLimpaVetor VFV_VETOR3, TDBGrid3
        
        CmdGravar.Caption = "&Gravar"
    End If
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

Private Sub CmdInsComprador_Click()
    If CboComprador.Text <> "" Then
        DatComprador.Recordset.Bookmark = CboComprador.SelectedItem
        
        If funExisteChave(1, Array("0"), Array(DatComprador.Recordset.Fields("focl_cd_forncli")), VFV_VETOR1, TDBGrid1) Then
            MsgBox "Esse Cliente já foi cadastrado", vbCritical, "ATENÇÃO !"
            Exit Sub
        End If
        subInsereRegistroVetor Array(CboComprador.BoundText, CboComprador.Text, ""), VFV_VETOR1, TDBGrid1
        If Not CmdExcComprador.Enabled Then CmdExcComprador.Enabled = True
    End If
End Sub

Private Sub CmdInsCorretor_Click()
    If TxtTotVenda.Text = "" Then
        MsgBox "É necessário os valores da venda para o cálculo da comissão. Favor preenchê-los.", vbCritical, "CAMPO OBRIGATÓRIO"
        TabAuxiliar.Tab = 0
        TxtAVista.SetFocus
        Exit Sub
    End If
    CboCorretor.Text = ""
    TxtComissao.Text = ""
    DtpPagamento.Value = ""
    XFT_INSCORRETOR = "I"
    PanCorretor.Visible = True
    PanCorretor.Enabled = True
End Sub

Private Sub CmdInsCorretor_GotFocus()
    TabAuxiliar.Tab = 3
End Sub

Private Sub CmdInsFiador_Click()
    If CboFiador.Text <> "" Then
        DatFiador.Recordset.Bookmark = CboFiador.SelectedItem
        
        If funExisteChave(1, Array("0"), Array(DatFiador.Recordset.Fields("fiad_cd_fiador")), VFV_VETOR3, TDBGrid3) Then
            MsgBox "Esse Fiador já foi cadastrado", vbCritical, "ATENÇÃO !"
            Exit Sub
        End If
        subInsereRegistroVetor Array(CboFiador.BoundText, CboFiador.Text), VFV_VETOR3, TDBGrid3
        If Not CmdExcFiador.Enabled Then CmdExcFiador.Enabled = True
    End If
End Sub

Private Sub CmdSairCorretor_Click()
    PanCorretor.Enabled = False
    PanCorretor.Visible = False
End Sub

Private Sub CmdTitular_Click()
    Dim XLI_POS As Integer
    Dim XLI_POSPADRAO As Integer
    
    If VFV_VETOR1(0, 0) <> Empty Then
        XLI_POS = TDBGrid1.Bookmark
        TDBGrid1.MoveFirst
        Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Titular").Value = "S" Then
                XLI_POSPADRAO = TDBGrid1.Bookmark
                If XLI_POSPADRAO = XLI_POS Then
                    MsgBox "Este banco Cliente já é o Titular!", vbCritical, "ATENÇÃO"
                    XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
                    Exit Sub
                Else
                    If MsgBox("Já existe um Cliente Titular. Deseja alterar?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma alteração
                        TDBGrid1.Columns("Titular").Text = "N"
                        TDBGrid1.Update
                        TDBGrid1.Bookmark = XLI_POS
                        TDBGrid1.Columns("Titular").Text = "S"
                        TDBGrid1.Update
                        XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            End If
            TDBGrid1.MoveNext
        Loop
        TDBGrid1.Bookmark = XLI_POS
        TDBGrid1.Columns("Titular").Text = "S"
        TDBGrid1.Update
        XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
    End If

End Sub

Private Sub Form_Activate()
    Set Formulario = FrmManImoveis
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
    'On Error GoTo TrataErro
      
    Set Formulario = FrmManImoveis
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If

   ' CliqueBarraFerramentas = False
 
    subConectarControleDados DatComprador, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
    subConectarControleDados DatCorretor, "SELECT * FROM Corretores ORDER BY corr_tx_nome", Estatico
    subConectarControleDados DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    subConectarControleDados DatFiador, "SELECT fiad_cd_Fiador, fiad_tx_Nome FROM Fiadores ORDER BY fiad_tx_Nome", Estatico
    subConectarControleDados DatIndexador, "SELECT * FROM Moedas ORDER BY moed_tx_Descricao", Estatico
    subConectarControleDados DatTipoImovel, "SELECT * FROM TiposImoveis ORDER BY tiim_tx_Descricao", Estatico
        
    XFB_ENTROUCOMP = False
    XFB_ENTROUFIAD = False
    XFB_ENTROUCORR = False
    
    If XInserir = "I" Then
        SubPrepararInsercao
    ElseIf XInserir = "A" Then
        SubPreencheCampos
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
        'TrataErros (ComMensagem)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmManImoveis"
End Sub

Private Sub MskCodigo_GotFocus()
    Call subSelecionaMSK(MskCodigo)
End Sub

Private Sub MskCodigo_LostFocus()
    'Checar se o Imóvel já foi cadastrado
    'If MskNumdocumento.Text <> "" And XInserir = "I" Then
    If MskCodigo.Text <> "    - " And MskCodigo.Text <> XFT_CODIMOVEL Then
        Dim XLO_IMOVEL As Object
        SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM Imoveis WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento= '" & TxtEmpreendimento.Text & "'", Estatico
        If Not XLO_IMOVEL.EOF Then
           MsgBox "Imóvel já cadastrado.", vbCritical + vbOKOnly, "ATENÇÃO"
           MskCodigo.SetFocus
           XLO_IMOVEL.Close
           Exit Sub
        End If
        XLO_IMOVEL.Close
    End If
    
    MskCodigo.PromptInclude = False
    If MskCodigo.Text <> "" And Len(MskCodigo.Text) <> 5 Then
        MsgBox "Código inválido! O código deve possuir 5 dígitos.", vbCritical + vbOKOnly, "ATENÇÃO"
        MskCodigo.SetFocus
    End If
    MskCodigo.PromptInclude = True
End Sub

Private Sub TabAuxiliar_Click(PreviousTab As Integer)
    Dim XLT_SQL As String
    
    If TabAuxiliar.Tab = 1 And Not XFB_ENTROUCOMP Then
        'Compradores
        XLT_SQL = "SELECT focl_cd_FornCli, focl_tx_RazaoSocial, focl_tx_RazaoSocial FROM ConsCAPCompradores WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento='" & TxtEmpreendimento.Text & "'"
        subCarregaVetorCompradores XLT_SQL, Array("focl_cd_FornCli", "focl_tx_RazaoSocial", "focl_tx_RazaoSocial"), VFV_VETOR1, TDBGrid1
        XFB_ENTROUCOMP = True
    ElseIf TabAuxiliar.Tab = 2 And Not XFB_ENTROUFIAD Then
        'Fiadores
        XLT_SQL = "SELECT fiad_cd_Fiador, fiad_tx_Nome FROM ConsCAPFiadoresImoveis WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento='" & TxtEmpreendimento.Text & "'"
        subCarregaVetor XLT_SQL, Array("fiad_cd_Fiador", "fiad_tx_Nome"), VFV_VETOR3, TDBGrid3
        XFB_ENTROUFIAD = True
    ElseIf TabAuxiliar.Tab = 3 And Not XFB_ENTROUCORR Then
        'Corretores
        XLT_SQL = "SELECT * FROM ConsCAPCorretores WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento='" & TxtEmpreendimento.Text & "'"
        subCarregaVetor XLT_SQL, Array("corr_cd_Corretor", "corr_tx_Nome", "comi_vl_Comissao", "comi_dt_Pagamento"), VFV_VETOR2, TDBGrid2
        XFB_ENTROUCORR = True
    End If
End Sub

Sub subCarregaVetorCompradores(Sql As String, VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)   'Carrega o vetores
    Dim XLO_COMPRADOR As Object
    Dim XLI_CONT As Integer
    Dim XLI_X As Integer
    
    'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
    'VetorCampos - vetor simples com os campos na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    SubQOpenRecordset XLO_COMPRADOR, Sql, Estatico
    Vetor.ReDim 0, 0, 0, Grid.Columns.Count - 1
    Vetor.Clear
    Grid.Array = Vetor
    If Not (XLO_COMPRADOR.EOF And XLO_COMPRADOR.BOF) Then
        XLO_COMPRADOR.MoveFirst
        Do While Not XLO_COMPRADOR.EOF
            If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
                Vetor.Insert 1, Vetor.UpperBound(1) + 1
            End If
            XLI_CONT = Vetor.UpperBound(1)
            For XLI_X = 0 To Grid.Columns.Count - 2
                If VetorCampos(XLI_X) <> "" Then
                    subSetaVal Vetor, XLI_CONT, XLI_X, FunNulo(XLO_COMPRADOR(VetorCampos(XLI_X)))
                End If
            Next
            'O próximo if preenche a coluna titular com "S" para o comprador Titular
            If XFI_CODCLIENTE = XLO_COMPRADOR!focl_cd_forncli And Grid.Columns.Count = 3 Then
                subSetaVal Vetor, XLI_CONT, XLI_X, "S"
            End If
            XLO_COMPRADOR.MoveNext
        Loop
        
        With Grid
            .ReBind
        End With
        
    End If
    XLO_COMPRADOR.Close
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    If TDBGrid1.Columns("Titular").CellText(Bookmark) = "S" Then
        CellStyle.ForeColor = &H8000&
    End If
End Sub

Private Sub TxtAVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAVista_LostFocus()
    If IsNumeric(TxtAVista.Text) Then
        TxtAVista.Text = Format$(TxtAVista.Text, "##,##0.00")
    Else
        TxtAVista.Text = ""
    End If
End Sub

Private Sub TxtComissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtComissao_LostFocus()
    If IsNumeric(TxtComissao.Text) Then
        TxtComissao.Text = Format$(TxtComissao.Text, "##,##0.00")
    Else
        TxtComissao.Text = ""
    End If
End Sub


Private Sub TxtEmpreendimento_GotFocus()
    Call subSelecionaTXT(TxtEmpreendimento)
End Sub

Private Sub TxtEmpreendimento_LostFocus()
    If TxtEmpreendimento.Text <> "" Then
        CboEmpreendimento.BoundText = TxtEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtEmpreendimento.SetFocus
            MskCodigo.Enabled = True
        End If
    End If
End Sub

Private Sub TxtFgts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFgts_LostFocus()
    If IsNumeric(TxtFgts.Text) Then
        TxtFgts.Text = Format$(TxtFgts.Text, "##,##0.00")
    Else
        TxtFgts.Text = ""
    End If
End Sub

Private Sub TxtFinanciamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFinanciamento_LostFocus()
    If IsNumeric(TxtFinanciamento.Text) Then
        TxtFinanciamento.Text = Format$(TxtFinanciamento.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)), "standard")
    Else
        TxtFinanciamento.Text = ""
    End If
End Sub

Private Sub TxtPoupanca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPoupanca_LostFocus()
    If IsNumeric(TxtPoupanca.Text) Then
        TxtPoupanca.Text = Format$(TxtPoupanca.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(TxtPoupanca.Text) + CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text)), "standard")
    Else
        TxtPoupanca.Text = ""
    End If
End Sub

Private Sub TxtSinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSinal_LostFocus()
    If IsNumeric(TxtSinal.Text) Then
        TxtSinal.Text = Format$(TxtSinal.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(TxtSinal.Text) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text)), "standard")
    Else
        TxtSinal.Text = ""
    End If
End Sub

Private Sub TxtTotVenda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTotVenda_LostFocus()
    If IsNumeric(TxtTotVenda.Text) Then
        TxtTotVenda.Text = Format$(TxtTotVenda.Text, "##,##0.00")
    Else
        TxtTotVenda.Text = ""
    End If
End Sub


