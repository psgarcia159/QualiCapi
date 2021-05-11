VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmManContratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contratos"
   ClientHeight    =   7065
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   9390
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleMode       =   0  'User
   ScaleWidth      =   9390
   Begin VB.Frame FraPrincipal 
      Height          =   6555
      Left            =   0
      TabIndex        =   48
      Top             =   -60
      Width           =   9345
      Begin Threed.SSPanel PanCodigo 
         Height          =   870
         Left            =   120
         TabIndex        =   49
         Top             =   180
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1535
         _StockProps     =   15
         ForeColor       =   -2147483630
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
            ItemData        =   "FrmContratos.frx":0000
            Left            =   6600
            List            =   "FrmContratos.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   480
            Width           =   2040
         End
         Begin VB.Label LblCodContrato 
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
            Left            =   1140
            TabIndex        =   78
            Top             =   540
            Width           =   315
         End
         Begin VB.Label LblCodImovel 
            Caption         =   "Imovel"
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
            Left            =   8040
            TabIndex        =   77
            Top             =   120
            Width           =   555
         End
         Begin VB.Label LblCodContratoOrigem 
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
            Left            =   3960
            TabIndex        =   76
            Top             =   540
            Width           =   315
         End
         Begin VB.Label LblDescEmpreendimento 
            Caption         =   "Empreendimento"
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
            Left            =   2475
            TabIndex        =   75
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label LblCodEmpreendimento 
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
            Left            =   1860
            TabIndex        =   74
            Top             =   120
            Width           =   555
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
            Left            =   2520
            TabIndex        =   73
            Top             =   540
            Width           =   1455
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
            Left            =   5760
            TabIndex        =   54
            Top             =   540
            Width           =   750
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
            Left            =   300
            TabIndex        =   53
            Top             =   540
            Width           =   855
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
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label LblImovel 
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
            Top             =   120
            Width           =   1605
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   4665
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   8229
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Compradores"
         TabPicture(0)   =   "FrmContratos.frx":001E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "CmdTitular"
         Tab(0).Control(1)=   "CmdInsComprador"
         Tab(0).Control(2)=   "CmdExcComprador"
         Tab(0).Control(3)=   "DatComprador"
         Tab(0).Control(4)=   "TDBGrid1"
         Tab(0).Control(5)=   "CboComprador"
         Tab(0).Control(6)=   "LblCliente"
         Tab(0).Control(7)=   "LblCliTitular"
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "&Fiadores"
         TabPicture(1)   =   "FrmContratos.frx":003A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "DatFiador"
         Tab(1).Control(2)=   "TDBGrid3"
         Tab(1).Control(3)=   "CboFiador"
         Tab(1).Control(4)=   "CmdExcFiador"
         Tab(1).Control(5)=   "CmdInsFiador"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "&Dados Financeiros"
         TabPicture(2)   =   "FrmContratos.frx":0056
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FraProRata"
         Tab(2).Control(1)=   "FraValores"
         Tab(2).Control(2)=   "FraDesagio"
         Tab(2).Control(3)=   "FraJuros"
         Tab(2).Control(4)=   "FraSeguro"
         Tab(2).Control(5)=   "FraMulta"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Dados da &Venda"
         TabPicture(3)   =   "FrmContratos.frx":0072
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FraFinContabil"
         Tab(3).Control(1)=   "FraDatas"
         Tab(3).Control(2)=   "DatConta"
         Tab(3).Control(3)=   "DatPcr"
         Tab(3).Control(4)=   "DatIndexador"
         Tab(3).Control(5)=   "DatIdentificador"
         Tab(3).ControlCount=   6
         TabCaption(4)   =   "C&orretores"
         TabPicture(4)   =   "FrmContratos.frx":008E
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "TDBGrid2"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "PanCorretor"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "CmdAltCorretor"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "CmdExcCorretor"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "CmdInsCorretor"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).ControlCount=   5
         Begin VB.Frame FraMulta 
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
            Height          =   1820
            Left            =   -70320
            TabIndex        =   125
            Top             =   1440
            Width           =   2295
            Begin VB.OptionButton OptMultaCrescente 
               Caption         =   "Crescente"
               Height          =   255
               Left            =   1140
               TabIndex        =   131
               ToolTipText     =   "Juros simples da Caixa Econômica Federal"
               Top             =   420
               Width           =   1050
            End
            Begin VB.OptionButton OptMultaProRata 
               Caption         =   "Pró-Rata"
               Height          =   255
               Left            =   60
               TabIndex        =   130
               ToolTipText     =   "Juros simples da Caixa Econômica Federal"
               Top             =   840
               Width           =   930
            End
            Begin VB.OptionButton OptMultaNormal 
               Caption         =   "Normal"
               Height          =   255
               Left            =   60
               TabIndex        =   129
               ToolTipText     =   "Juros simples da Caixa Econômica Federal"
               Top             =   420
               Value           =   -1  'True
               Width           =   810
            End
            Begin VB.TextBox TxtPercMulta 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   127
               Top             =   1380
               Width           =   780
            End
            Begin VB.TextBox TxtLimiteDias 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   126
               ToolTipText     =   "Limite de dias para o crescimento"
               Top             =   780
               Width           =   480
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00808080&
               X1              =   1080
               X2              =   1080
               Y1              =   240
               Y2              =   1260
            End
            Begin VB.Label LblLimiteDias 
               Caption         =   "Dias:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1200
               TabIndex        =   132
               Top             =   840
               Width           =   405
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00808080&
               X1              =   -60
               X2              =   2280
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00808080&
               X1              =   -60
               X2              =   2100
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Label LblPercJurosMulta 
               Alignment       =   1  'Right Justify
               Caption         =   "Percentagem:"
               Height          =   255
               Left            =   120
               TabIndex        =   128
               Top             =   1440
               Width           =   1035
            End
         End
         Begin MSAdodcLib.Adodc DatIdentificador 
            Height          =   330
            Left            =   -69000
            Top             =   2700
            Visible         =   0   'False
            Width           =   2490
            _ExtentX        =   4392
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
         Begin MSAdodcLib.Adodc DatIndexador 
            Height          =   330
            Left            =   -68520
            Top             =   2700
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
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
         Begin MSAdodcLib.Adodc DatPcr 
            Height          =   330
            Left            =   -68340
            Top             =   2700
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
            Left            =   -68040
            Top             =   2700
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
         Begin VB.Frame FraDatas 
            Height          =   1335
            Left            =   -74880
            TabIndex        =   104
            Top             =   345
            Width           =   9045
            Begin MSComCtl2.DTPicker DtpDtBase 
               Height          =   315
               Left            =   7500
               TabIndex        =   25
               Top             =   180
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   63766531
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpRegContrato 
               Height          =   315
               Left            =   4260
               TabIndex        =   24
               Top             =   180
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   63766531
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtQuitacao 
               Height          =   315
               Left            =   1080
               TabIndex        =   26
               Top             =   600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yy"
               DateIsNull      =   -1  'True
               Format          =   63766531
               CurrentDate     =   37686.4345601852
            End
            Begin MSComCtl2.DTPicker DtpDtVenda 
               Height          =   315
               Left            =   1080
               TabIndex        =   23
               Top             =   180
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   63766531
               CurrentDate     =   37180
            End
            Begin MSDataListLib.DataCombo CboIndexador 
               Bindings        =   "FrmContratos.frx":00AA
               Height          =   315
               Left            =   4260
               TabIndex        =   27
               Top             =   600
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboIndexador"
            End
            Begin Threed.SSCommand CmdLimparIndexador 
               Height          =   315
               Left            =   5730
               TabIndex        =   28
               Top             =   600
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483633
               MouseIcon       =   "FrmContratos.frx":00C5
               Picture         =   "FrmContratos.frx":0517
            End
            Begin MSComCtl2.DTPicker DtpDistrato 
               Height          =   315
               Left            =   7500
               TabIndex        =   29
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yy"
               DateIsNull      =   -1  'True
               Format          =   63766531
               CurrentDate     =   37686.4345601852
            End
            Begin VB.Label LblDistrato 
               Caption         =   "Data Distrato:"
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
               Left            =   6240
               TabIndex        =   113
               Top             =   615
               Visible         =   0   'False
               Width           =   1275
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
               Left            =   3255
               TabIndex        =   109
               Top             =   645
               Width           =   975
            End
            Begin VB.Label LblDtVenda 
               Caption         =   "Venda:"
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
               Left            =   360
               TabIndex        =   108
               Top             =   240
               Width           =   615
            End
            Begin VB.Label LblDtQuitacao 
               Caption         =   "Quitação:"
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
               Left            =   135
               TabIndex        =   107
               Top             =   645
               Width           =   855
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
               Left            =   2610
               TabIndex        =   106
               Top             =   225
               Width           =   1620
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
               Left            =   6435
               TabIndex        =   105
               Top             =   225
               Width           =   1035
            End
         End
         Begin VB.Frame FraFinContabil 
            Height          =   1785
            Left            =   -74880
            TabIndex        =   101
            Top             =   1680
            Width           =   9045
            Begin MSMask.MaskEdBox MskConta 
               Height          =   330
               Left            =   1440
               TabIndex        =   30
               Top             =   180
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
               Left            =   1440
               TabIndex        =   33
               Top             =   580
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmContratos.frx":0629
               Height          =   330
               Left            =   2340
               TabIndex        =   34
               Top             =   580
               Width           =   6255
               _ExtentX        =   11033
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
               Bindings        =   "FrmContratos.frx":063E
               Height          =   330
               Left            =   2340
               TabIndex        =   31
               Top             =   180
               Width           =   6255
               _ExtentX        =   11033
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
               Left            =   8625
               TabIndex        =   35
               Top             =   580
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmContratos.frx":0655
               Picture         =   "FrmContratos.frx":0AA7
            End
            Begin Threed.SSCommand CmdLimparplanoContas 
               Height          =   330
               Left            =   8625
               TabIndex        =   32
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               ForeColor       =   -2147483633
               MouseIcon       =   "FrmContratos.frx":0BB9
               Picture         =   "FrmContratos.frx":100B
            End
            Begin MSDataListLib.DataCombo CboIdentificador 
               Bindings        =   "FrmContratos.frx":111D
               Height          =   315
               Left            =   1440
               TabIndex        =   110
               Top             =   1320
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
               TabIndex        =   111
               Top             =   1320
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmContratos.frx":113C
               Picture         =   "FrmContratos.frx":158E
            End
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "FrmContratos.frx":16A0
               Height          =   330
               Left            =   1440
               TabIndex        =   118
               Top             =   980
               Width           =   7150
               _ExtentX        =   12621
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
               Left            =   8625
               TabIndex        =   119
               Top             =   980
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmContratos.frx":16BF
               Picture         =   "FrmContratos.frx":1B11
            End
            Begin MSAdodcLib.Adodc DatContaCorrente 
               Height          =   330
               Left            =   6540
               Top             =   1380
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
               TabIndex        =   120
               Top             =   1018
               Width           =   1365
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
               Left            =   255
               TabIndex        =   112
               Top             =   1425
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
               Left            =   90
               TabIndex        =   103
               Top             =   240
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
               Left            =   150
               TabIndex        =   102
               Top             =   648
               Width           =   1260
            End
         End
         Begin VB.CommandButton CmdInsFiador 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   5
            Top             =   465
            Width           =   975
         End
         Begin VB.CommandButton CmdExcFiador 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   6
            Top             =   840
            Width           =   975
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
            Height          =   855
            Left            =   -67920
            TabIndex        =   96
            Top             =   2400
            Width           =   2115
            Begin VB.CheckBox ChkParcela 
               Caption         =   "P/ Parcela"
               Height          =   240
               Left            =   120
               TabIndex        =   22
               Top             =   585
               Width           =   1605
            End
            Begin VB.CheckBox ChkSlDevedor 
               Caption         =   "P/ Saldo Devedor"
               Height          =   240
               Left            =   120
               TabIndex        =   21
               Top             =   390
               Width           =   1620
            End
            Begin VB.TextBox TxtTaxaSeguro 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   840
               MaxLength       =   20
               TabIndex        =   20
               Top             =   135
               Width           =   870
            End
            Begin VB.Label LblTaxaSeguro 
               Caption         =   "Indice:"
               Height          =   195
               Left            =   120
               TabIndex        =   97
               Top             =   195
               Width           =   540
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
            Height          =   1820
            Left            =   -72600
            TabIndex        =   94
            Top             =   1440
            Width           =   2235
            Begin VB.OptionButton OptJurCheioSimples 
               Caption         =   "Simples"
               Height          =   255
               Left            =   1140
               TabIndex        =   122
               Top             =   960
               Width           =   870
            End
            Begin VB.OptionButton OptJurCheioComposto 
               Caption         =   "Composto"
               Height          =   255
               Left            =   1140
               TabIndex        =   121
               Top             =   660
               Width           =   1050
            End
            Begin VB.TextBox TxtPercJurosMora 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   18
               Top             =   1380
               Width           =   765
            End
            Begin VB.OptionButton OptJurMensal 
               Caption         =   "Mensal"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   960
               Width           =   810
            End
            Begin VB.OptionButton OptJurDiario 
               Caption         =   "Diário"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   630
               Value           =   -1  'True
               Width           =   810
            End
            Begin VB.Label LblNormal 
               Caption         =   "Normal"
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
               TabIndex        =   124
               Top             =   300
               Width           =   585
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   2220
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Label LblCheio 
               Caption         =   "Cheio"
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
               Left            =   1080
               TabIndex        =   123
               Top             =   300
               Width           =   585
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   1020
               X2              =   1020
               Y1              =   240
               Y2              =   1260
            End
            Begin VB.Label LblPercJurosMOra 
               Alignment       =   1  'Right Justify
               Caption         =   "Percentagem:"
               Height          =   255
               Left            =   180
               TabIndex        =   95
               Top             =   1440
               Width           =   975
            End
         End
         Begin VB.Frame FraDesagio 
            Caption         =   "Juros Financeiros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1820
            Left            =   -74880
            TabIndex        =   91
            Top             =   1440
            Width           =   2205
            Begin VB.OptionButton OptDesSimples 
               Caption         =   "Simples"
               Height          =   255
               Left            =   1260
               TabIndex        =   115
               ToolTipText     =   "Juros simples da Caixa Econômica Federal"
               Top             =   600
               Width           =   870
            End
            Begin VB.TextBox TxtPercJurosFin 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   840
               TabIndex        =   15
               Top             =   1380
               Width           =   660
            End
            Begin VB.OptionButton OptDesDiario 
               Caption         =   "Diário"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   750
            End
            Begin VB.OptionButton OptDesMensal 
               Caption         =   "Mensal"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Value           =   -1  'True
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "C.E.F"
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
               Left            =   1200
               TabIndex        =   117
               Top             =   300
               Width           =   465
            End
            Begin VB.Label LblDesagio 
               Alignment       =   1  'Right Justify
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
               Height          =   195
               Left            =   60
               TabIndex        =   116
               Top             =   300
               Width           =   705
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   1080
               X2              =   1080
               Y1              =   240
               Y2              =   1260
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   2170
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Label LblMes 
               Caption         =   "a. m"
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
               Height          =   195
               Left            =   1560
               TabIndex        =   93
               Top             =   1440
               Width           =   435
            End
            Begin VB.Label LblPercJurosFin 
               Caption         =   "Percent.:"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               Top             =   1440
               Width           =   675
            End
         End
         Begin VB.Frame FraValores 
            Height          =   1050
            Left            =   -74880
            TabIndex        =   82
            Top             =   360
            Width           =   9045
            Begin VB.TextBox TxtOutros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6000
               TabIndex        =   12
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtSinal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   8
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtPoupanca 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               MaxLength       =   20
               TabIndex        =   9
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox TxtTotVenda 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7680
               MaxLength       =   20
               TabIndex        =   83
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtFgts 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6000
               MaxLength       =   20
               TabIndex        =   10
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox TxtFinanciamento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               MaxLength       =   20
               TabIndex        =   11
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtAVista 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   7
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label LblOutros 
               Caption         =   "Outros:"
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
               Left            =   5340
               TabIndex        =   90
               Top             =   600
               Width           =   675
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
               Left            =   240
               TabIndex        =   89
               Top             =   600
               Width           =   585
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
               Left            =   2700
               TabIndex        =   88
               Top             =   240
               Width           =   960
            End
            Begin VB.Label LblTotVenda 
               Alignment       =   1  'Right Justify
               Caption         =   "Total a Prazo"
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
               Left            =   7680
               TabIndex        =   87
               Top             =   300
               Width           =   1185
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
               Left            =   5280
               TabIndex        =   86
               Top             =   240
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
               Left            =   2400
               TabIndex        =   85
               Top             =   600
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
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   705
            End
         End
         Begin VB.CommandButton CmdTitular 
            Caption         =   "&Titular"
            Height          =   255
            Left            =   -66960
            TabIndex        =   3
            Top             =   1215
            Width           =   975
         End
         Begin VB.CommandButton CmdInsComprador 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   1
            Top             =   465
            Width           =   975
         End
         Begin VB.CommandButton CmdExcComprador 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   2
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton CmdInsCorretor 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   8040
            TabIndex        =   36
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton CmdExcCorretor 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   8040
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton CmdAltCorretor 
            Caption         =   "&Alterar"
            Height          =   255
            Left            =   8040
            TabIndex        =   37
            Top             =   840
            Width           =   975
         End
         Begin Threed.SSPanel PanCorretor 
            Height          =   1800
            Left            =   1680
            TabIndex        =   55
            Top             =   1080
            Visible         =   0   'False
            Width           =   5400
            _Version        =   65536
            _ExtentX        =   9525
            _ExtentY        =   3175
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
            Begin VB.TextBox TxtPercComissao 
               Height          =   315
               Left            =   1500
               TabIndex        =   41
               Top             =   1020
               Width           =   855
            End
            Begin VB.TextBox TxtComissao 
               Height          =   315
               Left            =   4260
               MaxLength       =   20
               TabIndex        =   42
               Top             =   1020
               Width           =   930
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
               TabIndex        =   43
               Top             =   1440
               Width           =   975
            End
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
               Left            =   4200
               TabIndex        =   44
               Top             =   1440
               Width           =   975
            End
            Begin MSDataListLib.DataCombo CboCorretor 
               Bindings        =   "FrmContratos.frx":1C23
               Height          =   330
               Left            =   825
               TabIndex        =   39
               Top             =   165
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "Corretor"
               BoundColumn     =   "corr_cd_corretor"
               Text            =   "CboCorretor"
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
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3660
               TabIndex        =   40
               Top             =   600
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yy"
               DateIsNull      =   -1  'True
               Format          =   63766531
               CurrentDate     =   37630.3585069444
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
            Begin VB.Label LblValorDoImovel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor"
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
               Left            =   1380
               TabIndex        =   72
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label LblPercComissao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comissão (%): "
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
               TabIndex        =   71
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label LblValorImovel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor Imóvel"
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
               Left            =   240
               TabIndex        =   70
               Top             =   600
               Width           =   1095
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
               Left            =   180
               TabIndex        =   58
               Top             =   225
               Width           =   600
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
               Left            =   3300
               TabIndex        =   57
               Top             =   1080
               Width           =   825
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
               Left            =   2640
               TabIndex        =   56
               Top             =   630
               Width           =   1035
            End
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid2 
            Height          =   3735
            Left            =   255
            TabIndex        =   59
            Top             =   555
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   6588
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
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
            Splits(0)._ColumnProps(22)=   "Column(2).Width=2143"
            Splits(0)._ColumnProps(23)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(2)._WidthInPix=2064"
            Splits(0)._ColumnProps(25)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(2)._ColStyle=74258"
            Splits(0)._ColumnProps(27)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(28)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(30)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(31)=   "Column(2)._MinWidth=89769656"
            Splits(0)._ColumnProps(32)=   "Column(3).Width=2752"
            Splits(0)._ColumnProps(33)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(3)._WidthInPix=2672"
            Splits(0)._ColumnProps(35)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(3)._ColStyle=74257"
            Splits(0)._ColumnProps(37)=   "Column(3).WrapText=1"
            Splits(0)._ColumnProps(38)=   "Column(3).FetchStyle=1"
            Splits(0)._ColumnProps(39)=   "Column(3).AllowFocus=0"
            Splits(0)._ColumnProps(40)=   "Column(3).Order=4"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(41)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=90,.parent=21,.alignment=1,.valignment=2"
            _StyleDefs(46)  =   ":id=90,.wraptext=-1,.locked=-1"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=22,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=23"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=75"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=94,.parent=21,.alignment=2,.valignment=2"
            _StyleDefs(51)  =   ":id=94,.wraptext=-1,.locked=-1"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=22,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=23"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=75"
            _StyleDefs(55)  =   "Named:id=33:Normal"
            _StyleDefs(56)  =   ":id=33,.parent=0"
            _StyleDefs(57)  =   "Named:id=34:Heading"
            _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   ":id=34,.wraptext=-1"
            _StyleDefs(60)  =   "Named:id=35:Footing"
            _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   "Named:id=36:Selected"
            _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=37:Caption"
            _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(66)  =   "Named:id=38:HighlightRow"
            _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(68)  =   "Named:id=39:EvenRow"
            _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(70)  =   "Named:id=40:OddRow"
            _StyleDefs(71)  =   ":id=40,.parent=33"
            _StyleDefs(72)  =   "Named:id=41:RecordSelector"
            _StyleDefs(73)  =   ":id=41,.parent=34"
            _StyleDefs(74)  =   "Named:id=42:FilterBar"
            _StyleDefs(75)  =   ":id=42,.parent=33"
         End
         Begin MSAdodcLib.Adodc DatComprador 
            Height          =   330
            Left            =   -71460
            Top             =   360
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
            Height          =   3390
            Left            =   -74760
            TabIndex        =   79
            Top             =   945
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   5980
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
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
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
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
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
            Splits(0)._ColumnProps(28)=   "Column(2)._ColStyle=65808"
            Splits(0)._ColumnProps(29)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(30)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(31)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(32)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(33)=   "Column(2).Order=3"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(41)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=16,.parent=21"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=22"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=23"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=75"
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
         Begin MSDataListLib.DataCombo CboComprador 
            Bindings        =   "FrmContratos.frx":1C3D
            Height          =   315
            Left            =   -72780
            TabIndex        =   0
            Top             =   540
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "focl_tx_RazaoSocial"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   "CboComprador"
         End
         Begin MSDataListLib.DataCombo CboFiador 
            Bindings        =   "FrmContratos.frx":1C58
            Height          =   315
            Left            =   -72840
            TabIndex        =   4
            Top             =   540
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "fiad_tx_nome"
            BoundColumn     =   "fiad_cd_fiador"
            Text            =   "CboFiador"
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid3 
            Height          =   3390
            Left            =   -74760
            TabIndex        =   99
            Top             =   945
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   5980
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
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(41)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(45)  =   "Named:id=33:Normal"
            _StyleDefs(46)  =   ":id=33,.parent=0"
            _StyleDefs(47)  =   "Named:id=34:Heading"
            _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(49)  =   ":id=34,.wraptext=-1"
            _StyleDefs(50)  =   "Named:id=35:Footing"
            _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   "Named:id=36:Selected"
            _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
         Begin MSAdodcLib.Adodc DatFiador 
            Height          =   330
            Left            =   -70080
            Top             =   420
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
         Begin VB.Frame FraProRata 
            Caption         =   "Correção Monetária"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   -67920
            TabIndex        =   98
            Top             =   1440
            Width           =   2115
            Begin VB.CheckBox ChkCorrecProRataMista 
               Caption         =   "Pró-Rata"
               Height          =   210
               Left            =   120
               TabIndex        =   133
               Top             =   690
               Width           =   1950
            End
            Begin VB.CheckBox ChkCorrecProRataPosterior 
               Caption         =   "Pró-Rata Mês Posterior"
               Height          =   210
               Left            =   120
               TabIndex        =   114
               Top             =   470
               Width           =   1950
            End
            Begin VB.CheckBox ChkCorrecProRataAnterior 
               Caption         =   "Pró-Rata Mês Anterior"
               Height          =   210
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   1890
            End
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
            Left            =   -74880
            TabIndex        =   100
            Top             =   585
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
            Left            =   -74820
            TabIndex        =   81
            Top             =   585
            Width           =   1995
         End
         Begin VB.Label LblCliTitular 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente Titular"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   -66975
            TabIndex        =   80
            Top             =   2295
            Width           =   1005
         End
      End
      Begin VB.Frame FraDadosGerais 
         Height          =   675
         Left            =   120
         TabIndex        =   60
         Top             =   1020
         Width           =   9135
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
            Left            =   6360
            TabIndex        =   68
            Top             =   240
            Width           =   885
         End
         Begin VB.Label LblDescAndar 
            Caption         =   "Andar"
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
            Left            =   5820
            TabIndex        =   67
            Top             =   240
            Width           =   375
         End
         Begin VB.Label LblDescArea 
            Alignment       =   1  'Right Justify
            Caption         =   "Area"
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
            Left            =   3960
            TabIndex        =   66
            Top             =   240
            Width           =   435
         End
         Begin VB.Label LblDescTipoImovel 
            Caption         =   "Tipo Imovel"
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
            Left            =   1500
            TabIndex        =   65
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Lblm2 
            Caption         =   "m²"
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
            Left            =   4455
            TabIndex        =   64
            Top             =   240
            Width           =   255
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
            Left            =   3420
            TabIndex        =   63
            Top             =   240
            Width           =   495
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
            Left            =   60
            TabIndex        =   62
            Top             =   240
            Width           =   1365
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
            Height          =   210
            Left            =   5100
            TabIndex        =   61
            Top             =   240
            Width           =   690
         End
         Begin VB.Label LblDescGaragem 
            Caption         =   "Garagem"
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
            Left            =   7260
            TabIndex        =   69
            Top             =   240
            Width           =   1755
         End
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5580
      TabIndex        =   46
      Top             =   6600
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   -2147483630
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
      TabIndex        =   45
      Top             =   6600
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Gravar"
      ForeColor       =   -2147483630
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
Attribute VB_Name = "FrmManContratos"
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
Dim XFI_CODCLIENTE As Integer 'Código do Comprador
Dim XFI_CODCORRETOR As Integer 'Código do Corretor
Dim XFT_INSCORRETOR As String
Dim XFT_CODEMPREENDIMENTO As String 'Código do Empreendimento
Dim XFT_CODIMOVEL As String  'Código do Imóvel
Dim XFT_CODCONTRATO As String  'Código do Contrato
Dim XFT_STATUS As String  'Armazena o status do contrato

Dim errLoop As ADODB.Error
Dim StrError As String

Sub SubDesabHabCampos(XHab As Boolean)
    
    'Desabilita/Habilita os campos em caso de inserção
    
    CboStatus.Enabled = XHab
    
    CboComprador.Enabled = XHab
    CmdInsComprador.Enabled = XHab
    CmdExcComprador.Enabled = XHab
    CmdTitular.Enabled = XHab
    
    CboFiador.Enabled = XHab
    CmdInsFiador.Enabled = XHab
    CmdExcFiador.Enabled = XHab
    
    CboCorretor.Enabled = XHab
    CmdInsCorretor.Enabled = XHab
    CmdAltCorretor.Enabled = XHab
    CmdExcCorretor.Enabled = XHab
    
    FraValores.Enabled = XHab
    FraDesagio.Enabled = XHab
    FraMulta.Enabled = XHab
    FraJuros.Enabled = XHab
    FraProRata.Enabled = XHab
    FraSeguro.Enabled = XHab
    FraDatas.Enabled = XHab
    FraFinContabil.Enabled = XHab
    CmdGravar.Enabled = XHab
    
End Sub
Function funChecarCampos() As Boolean
    
    funChecarCampos = False
    
    If CboStatus.Text = "" Then
      MsgBox "É preciso definir um Status.", vbCritical, "ATENÇÃO !"
      CboStatus.SetFocus
      Exit Function
    End If
    
    'Se a taxa do seguro for digitada é obrigatório optar por 1 tipo de seguro
    If TxtTaxaSeguro.Text <> "" And ChkSlDevedor.Value = 0 And ChkParcela.Value = 0 Then
        MsgBox "É preciso definir o tipo de Seguro.", vbCritical, "ATENÇÃO !"
        TabAuxiliar.Tab = 2
        ChkSlDevedor.SetFocus
        Exit Function
    End If
    
    'If XFI_CODCLIENTE = 0 Then
    If IsEmpty(VFV_VETOR1(0, 0)) Then
        MsgBox "É preciso definir um Comprador.", vbCritical, "ATENÇÃO !"
        TabAuxiliar.Tab = 0
        CmdTitular.SetFocus
        Exit Function
    End If
    
    If DtpDtQuitacao.Value < DtpDtVenda.Value Then
      MsgBox "A data da quitação tem que ser Posterior ou Igual a data da venda", vbCritical, "ATENÇÃO"
      DtpDtQuitacao.SetFocus
      Exit Function
    End If
    
    funChecarCampos = True

End Function

Sub subDesabilitaBotoes()
    If TabAuxiliar.Tab = 0 Then
        CmdInsComprador.Enabled = False
        CmdExcComprador.Enabled = False
    ElseIf TabAuxiliar.Tab = 1 Then
        CmdInsFiador.Enabled = False
        CmdExcFiador.Enabled = False
    ElseIf TabAuxiliar.Tab = 4 Then
        CmdInsCorretor.Enabled = False
        CmdAltCorretor.Enabled = False
        CmdExcCorretor.Enabled = False
    End If
End Sub
                                                        
Sub subHabilitaBotoes()
    If TabAuxiliar.Tab = 0 Then
        CmdInsComprador.Enabled = True
        CmdExcComprador.Enabled = True
    ElseIf TabAuxiliar.Tab = 1 Then
        CmdInsFiador.Enabled = True
        CmdExcFiador.Enabled = True
    ElseIf TabAuxiliar.Tab = 4 Then
        CmdInsCorretor.Enabled = True
        CmdAltCorretor.Enabled = True
        CmdExcCorretor.Enabled = True
    End If
End Sub

Sub subPreencheCampos()
    Dim XLT_SQL As String
    Dim XLO_Contrato As New ADODB.Recordset
    
    SubQOpenRecordset XLO_Contrato, "SELECT * FROM Contratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "'", "Estatico"
    If Not (XLO_Contrato.BOF And XLO_Contrato.EOF) Then
      
        TabAuxiliar.Enabled = True
        Select Case XLO_Contrato!cont_tx_Status
            Case "A"
                CboStatus.ListIndex = 0
                XFT_STATUS = "A"
                'DtpDistrato.Visible = False
                'DtpDistrato.CheckBox = True
                'DtpDistrato = Null
            Case "R"
                CboStatus.ListIndex = 0
                XFT_STATUS = "R"
            Case "I"
                CboStatus.ListIndex = 1
                DtpDistrato.Visible = True
                DtpDistrato.CheckBox = False
                LblDistrato.Visible = True
                DtpDistrato = XLO_Contrato!cont_dt_Distrato
        End Select
        
        If Not IsNull(XLO_Contrato!cont_dt_Base) Then
            DtpDtBase.Value = XLO_Contrato!cont_dt_Base
        Else
            'Quando ativa um contrato reservado
            DtpDtBase.Value = Date
        End If
        
        If Not IsNull(XLO_Contrato!cont_dt_RegistroContrato) Then
            DtpRegContrato.Value = XLO_Contrato!cont_dt_RegistroContrato
        Else
            'Quando ativa um contrato reservado
            DtpRegContrato.Value = Date
        End If
        
        If Not IsNull(XLO_Contrato!cont_dt_Quitacao) Then DtpDtQuitacao.Value = XLO_Contrato!cont_dt_Quitacao
        If Not IsNull(XLO_Contrato!cont_dt_Venda) Then DtpDtVenda.Value = XLO_Contrato!cont_dt_Venda
        If Not IsNull(XLO_Contrato!moed_cd_CodMoeda) Then CboIndexador.BoundText = XLO_Contrato!moed_cd_CodMoeda
        
        'Dados da Venda
        TxtAVista.Text = Format(XLO_Contrato!cont_vl_Avista, "standard")
      
        If CDbl(TxtAVista.Text) = 0# Then
            TxtAVista.Text = ""
            TxtSinal.Text = Format(XLO_Contrato!cont_vl_Sinal, "standard")
            TxtPoupanca.Text = Format(XLO_Contrato!cont_vl_Poupanca, "standard")
            TxtFinanciamento.Text = Format(XLO_Contrato!cont_vl_Financiamento, "standard")
            TxtTotVenda.Text = Format(XLO_Contrato!cont_vl_Total, "standard")
            TxtFgts.Text = Format(XLO_Contrato!cont_vl_Fgts, "standard")
            TxtOutros.Text = Format(XLO_Contrato!cont_vl_Outros, "standard")
        End If
        
        TxtPercJurosFin.Text = Format(XLO_Contrato!cont_vl_PercJurosFin, "##,##0.0000")
        TxtPercMulta.Text = Format(XLO_Contrato!cont_vl_PercJurosMulta, "##,##0.0000")
        TxtPercJurosMora.Text = Format(XLO_Contrato!cont_vl_PercJurosMora, "##,##0.0000")
      
        Select Case XLO_Contrato!cont_tx_MultaProRata
          Case "0"
             OptMultaNormal.Value = True
          Case "1"
             OptMultaProRata.Value = True
          Case "2"
             OptMultaCrescente.Value = True
             TxtLimiteDias.Text = XLO_Contrato!cont_nr_LimiteDias
        End Select
      
        If XLO_Contrato!cont_tx_CorrecProRata <> "0" Then
            If XLO_Contrato!cont_tx_CorrecProRata = "A" Then
                ChkCorrecProRataAnterior.Value = 1
            ElseIf XLO_Contrato!cont_tx_CorrecProRata = "P" Then
                ChkCorrecProRataPosterior.Value = 1
            Else
                ChkCorrecProRataMista.Value = 1
            End If
        End If
      
        Select Case XLO_Contrato!cont_tx_TipoJurosMora
            Case "D"
                OptJurDiario.Value = True
            Case "M"
                OptJurMensal.Value = True
            Case "C"
                OptJurCheioComposto.Value = True
            Case "S"
                OptJurCheioSimples.Value = True
        End Select
        
        Select Case XLO_Contrato!cont_tx_Desagio
            Case "D"
                OptDesDiario.Value = True
            Case "M"
                OptDesMensal.Value = True
            Case "S"
                OptDesSimples = True
        End Select
      
        TxtTaxaSeguro.Text = FunTrataZero(XLO_Contrato!cont_vl_TaxaSegVida)
        If XLO_Contrato!cont_nr_TipoSeguro = 1 Then
            ChkSlDevedor.Value = 1
        ElseIf XLO_Contrato!cont_nr_TipoSeguro = 2 Then
            ChkParcela.Value = 1
        End If
              
        If Not DatConta.Recordset.EOF Then
            DatConta.Recordset.MoveFirst
            If (XLO_Contrato!plco_cd_Conta <> "") And Not IsNull(XLO_Contrato!plco_cd_Conta) Then
                DatConta.Recordset.Find "plco_cd_Conta =" & XLO_Contrato!plco_cd_Conta
                MskConta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
                CboPlanoContas.BoundText = MskConta.Text
            End If
        End If
    
        If Not DatPcr.Recordset.EOF Then
            DatPcr.Recordset.MoveFirst
            If (XLO_Contrato!cere_cd_Pcr <> "") And Not IsNull(XLO_Contrato!cere_cd_Pcr) Then
                DatPcr.Recordset.Find "cere_cd_pcr =" & XLO_Contrato!cere_cd_Pcr
                MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
                CboPcr.BoundText = MskPcr.Text
            End If
        End If
          
        If Not DatContaCorrente.Recordset.EOF Then
            DatContaCorrente.Recordset.MoveFirst
            If Not IsNull(XLO_Contrato!coco_cd_codigo) Then
                DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & XLO_Contrato!coco_cd_codigo
                CboCCorrente.BoundText = FunNulo(XLO_Contrato!coco_cd_codigo)
            Else
                CboCCorrente.BoundText = ""
            End If
        End If
          
        If Not IsNull(XLO_Contrato!iden_cd_Projeto) Then
          CboIdentificador.BoundText = XLO_Contrato!iden_cd_Projeto
        End If
    
        XFI_CODCLIENTE = XLO_Contrato!focl_cd_FornCli
    End If
    XLO_Contrato.Close ' Inserido em 13/10/08 - Patrícia
End Sub
Sub subPrepararInsercao()
    DtpDtVenda.Value = Date
    DtpRegContrato.Value = Date
    DtpDtBase.Value = Date
End Sub

Private Sub CboComprador_Change()
  If CboComprador.Text <> "" Then
    CmdInsComprador.Enabled = True
  End If
End Sub

Private Sub CboComprador_DblClick(Area As Integer)
    Dim XLV_CHAVETEMP As Variant
    Dim XLT_INSERIRTEMP As String

    XLV_CHAVETEMP = Chave
    XLT_INSERIRTEMP = XInserir
    XInserir = "I"
    FornClie = "C" 'Cliente
    Pessoa = "F" 'Pessoa Física
    Chave = -1
    
    XFormForn = "FrmManContratos"
    FrmFornecedores.Show 1
    Chave = XLV_CHAVETEMP
    XInserir = XLT_INSERIRTEMP
    Set Formulario = FrmManContratos
    FrmManContratos.DatComprador.Recordset.Requery adCmdText
    FrmManContratos.CboComprador.BoundText = XGV_FornNotaFiscal
End Sub

Private Sub CboComprador_GotFocus()
    TabAuxiliar.Tab = 0
End Sub

Private Sub CboCorretor_Change()
  If CboCorretor.BoundText <> "" Then
    If CDbl(LblValorDoImovel.Caption) > 0 Then
      DatCorretor.Recordset.bookmark = CboCorretor.SelectedItem
      TxtPercComissao.Text = Format(DatCorretor.Recordset.Fields("corr_vl_percentcomissao"), "standard")
      TxtComissao.Text = Format(((CDbl(LblValorDoImovel.Caption) * DatCorretor.Recordset.Fields("corr_vl_percentcomissao")) / 100), "standard")
    Else
      TxtPercComissao.Text = Format(0, "standard")
      TxtComissao.Text = Format(0, "standard")
    End If
  End If
End Sub

Private Sub CboFiador_Change()
  If CboFiador.Text <> "" Then
    CmdInsFiador.Enabled = True
  End If
End Sub

Private Sub CboFiador_DblClick(Area As Integer)
    TelaManFiadores.Show vbModal
    If FunLocalizaJanelasAtivas("TelaManFiadores") <> -1 Then
        TelaManFiadores.SetFocus
    End If
End Sub

Private Sub CboFiador_GotFocus()
    TabAuxiliar.Tab = 1
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

Private Sub CboStatus_Click()
  If CboStatus.ListIndex = 1 Then
    DtpDistrato.CheckBox = False
    DtpDistrato.CustomFormat = "dd/MM/yy"
    DtpDistrato.Value = Date
    DtpDistrato.Visible = True
    LblDistrato.Visible = True
  Else
    DtpDistrato.CheckBox = True
    DtpDistrato.Value = Null
    DtpDistrato.Visible = False
    LblDistrato.Visible = False
  End If
End Sub

Private Sub ChkCorrecProRataAnterior_Click()
  If ChkCorrecProRataAnterior.Value = 1 Then
    ChkCorrecProRataPosterior.Value = 0
    ChkCorrecProRataMista.Value = 0
  End If
End Sub

Private Sub ChkCorrecProRataMista_Click()
  If ChkCorrecProRataMista.Value = 1 Then
    ChkCorrecProRataAnterior.Value = 0
    ChkCorrecProRataPosterior.Value = 0
  End If
End Sub

Private Sub ChkCorrecProRataPosterior_Click()
  If ChkCorrecProRataPosterior.Value = 1 Then
    ChkCorrecProRataAnterior.Value = 0
    ChkCorrecProRataMista.Value = 0
  End If
End Sub

Private Sub ChkParcela_Click()
    If ChkParcela.Value = 1 Then ChkSlDevedor.Value = 0
End Sub

Private Sub ChkSlDevedor_Click()
    If ChkSlDevedor.Value = 1 Then ChkParcela.Value = 0
End Sub

Private Sub CmdAltCorretor_Click()
    LblValorDoImovel.Caption = IIf(TxtTotVenda.Text <> "", TxtTotVenda.Text, TxtAVista.Text)
    CboCorretor.BoundText = TDBGrid2.Columns("Codigo")
    XFI_CODCORRETOR = TDBGrid2.Columns("Codigo")
    TxtComissao.Text = TDBGrid2.Columns("Comissão")
    
    If CDbl(LblValorDoImovel.Caption) > 0 Then
        TxtPercComissao.Text = Format(CDbl(TDBGrid2.Columns("Comissão")) / CDbl(LblValorDoImovel.Caption) * 100, "standard")
    Else
        TxtPercComissao.Text = Format(0, "standard")
        TxtComissao.Text = Format(0, "standard")
    End If
      
    If TDBGrid2.Columns("Pagamento") <> Null Or TDBGrid2.Columns("Pagamento") <> "" Then
        DtpPagamento.Value = TDBGrid2.Columns("Pagamento")
    Else
        DtpPagamento.Value = Null
    End If
    XFT_INSCORRETOR = "A"
    PanCorretor.Enabled = True
    PanCorretor.Visible = True
    'LblValorDoImovel.Caption = TxtTotVenda.Text
End Sub

Private Sub CmdConfirmaCorretor_Click()
    If CboCorretor.Text <> "" Then
        DatCorretor.Recordset.bookmark = CboCorretor.SelectedItem
        
        If XFT_INSCORRETOR = "A" And XFI_CODCORRETOR = CboCorretor.BoundText Then
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
                ElseIf XFT_INSCORRETOR = "A" Then
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
    Else
      MsgBox "O corretor não foi escolhido!", vbCritical, "ATENÇÃO"
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
            Else
              If VFV_VETOR1(0, 2) <> "S" Then
                Call CmdTitular_Click
              End If
            End If
        End If
    End If
End Sub

Private Sub CmdExcCorretor_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR2, TDBGrid2
        If VFV_VETOR2.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR2(0, 0)) Then
                CmdAltCorretor.Enabled = False
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

'Rotina utilizada para gravar o log de alterações nos contratos. Pega os dados originais antes da alteração.
Private Sub PreencheInformacoesOriginaisLog()
    ReDim XGM_MATRIZLOG(32, 3) As Variant
    
    'Dados originais
    XGM_MATRIZLOG(0, 1) = XFT_CODEMPREENDIMENTO
    XGM_MATRIZLOG(1, 1) = XFT_CODIMOVEL
    XGM_MATRIZLOG(2, 1) = LblCodContrato.Caption
    
    XGM_MATRIZLOG(3, 1) = IIf(CboStatus.ListIndex = 1, "I", "A")

    XGM_MATRIZLOG(4, 1) = XFI_CODCLIENTE
    XGM_MATRIZLOG(5, 1) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
    XGM_MATRIZLOG(6, 1) = FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil"))
    XGM_MATRIZLOG(7, 1) = FunNulo(CboCCorrente.BoundText)
    XGM_MATRIZLOG(8, 1) = FunNulo(CboIndexador.BoundText)
    XGM_MATRIZLOG(9, 1) = FunNulo(CboIdentificador.BoundText)
    XGM_MATRIZLOG(10, 1) = FunNulo(DtpDtBase.Value)
    XGM_MATRIZLOG(11, 1) = FunNulo(DtpDistrato.Value)
    XGM_MATRIZLOG(12, 1) = FunNulo(DtpDtQuitacao.Value)
    XGM_MATRIZLOG(13, 1) = FunNulo(DtpRegContrato.Value)
    XGM_MATRIZLOG(14, 1) = FunNulo(DtpDtVenda.Value)
    XGM_MATRIZLOG(15, 1) = FunNuloVal(TxtAVista.Text)
    XGM_MATRIZLOG(16, 1) = FunNuloVal(TxtSinal.Text)
    XGM_MATRIZLOG(17, 1) = FunNuloVal(TxtPoupanca.Text)
    XGM_MATRIZLOG(18, 1) = FunNuloVal(TxtFinanciamento.Text)
    XGM_MATRIZLOG(19, 1) = FunNuloVal(TxtFgts.Text)
    XGM_MATRIZLOG(20, 1) = FunNuloVal(TxtOutros.Text)
    XGM_MATRIZLOG(21, 1) = FunNuloVal(TxtTotVenda.Text)
        
    If OptDesDiario.Value = True Then
      XGM_MATRIZLOG(22, 1) = "D"
    ElseIf OptDesMensal = True Then
      XGM_MATRIZLOG(22, 1) = "M"
    Else
      XGM_MATRIZLOG(22, 1) = "S"
    End If
            
    XGM_MATRIZLOG(23, 1) = FunNuloVal(TxtPercJurosFin.Text)
    
    If OptJurDiario.Value = True Then
        XGM_MATRIZLOG(24, 1) = "D"
    ElseIf OptJurMensal.Value = True Then
        XGM_MATRIZLOG(24, 1) = "M"
    ElseIf OptJurCheioComposto.Value = True Then
        XGM_MATRIZLOG(24, 1) = "C"
    Else
        XGM_MATRIZLOG(24, 1) = "S"
    End If
        
    XGM_MATRIZLOG(25, 1) = FunNuloVal(TxtPercJurosMora.Text)
    
    If OptMultaNormal.Value = True Then
      XGM_MATRIZLOG(26, 1) = "0"
    ElseIf OptMultaProRata.Value = True Then
      XGM_MATRIZLOG(26, 1) = "1"
    ElseIf OptMultaCrescente.Value = True Then
      XGM_MATRIZLOG(26, 1) = "2"
    End If
        
    XGM_MATRIZLOG(27, 1) = FunNuloVal(TxtLimiteDias.Text)
    XGM_MATRIZLOG(28, 1) = FunNuloVal(TxtPercMulta.Text)
    
    If ChkCorrecProRataAnterior = 1 Then
      XGM_MATRIZLOG(29, 1) = "A"
    ElseIf ChkCorrecProRataPosterior = 1 Then
      XGM_MATRIZLOG(29, 1) = "P"
    ElseIf ChkCorrecProRataMista = 1 Then
      XGM_MATRIZLOG(29, 1) = "M"
    Else
      XGM_MATRIZLOG(29, 1) = "0"
    End If
        
    If ChkSlDevedor.Value = 1 Then
        XGM_MATRIZLOG(30, 1) = 1
    ElseIf ChkParcela.Value = 1 Then
        XGM_MATRIZLOG(30, 1) = 2
    Else
        XGM_MATRIZLOG(30, 1) = ""
    End If
        
    XGM_MATRIZLOG(31, 1) = FunNuloVal(TxtTaxaSeguro.Text)
        
End Sub

Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
    Dim Indice As Integer
        
    If XLT_OPERACAO <> ALTERAR Then
        ReDim XGM_MATRIZLOG(32, 2) As Variant
        Indice = 1
    Else
        Indice = 2
    End If
        
    'Campos
    XGM_MATRIZLOG(0, 0) = "Empreendimento"
    XGM_MATRIZLOG(1, 0) = "Imóvel"
    XGM_MATRIZLOG(2, 0) = "Contrato"
    XGM_MATRIZLOG(3, 0) = "Status"
    XGM_MATRIZLOG(4, 0) = "Cliente"
    XGM_MATRIZLOG(5, 0) = "Centro de Custo"
    XGM_MATRIZLOG(6, 0) = "Conta Contábil"
    XGM_MATRIZLOG(7, 0) = "Conta Corrente"
    XGM_MATRIZLOG(8, 0) = "Indexador"
    XGM_MATRIZLOG(9, 0) = "Identificador"
    XGM_MATRIZLOG(10, 0) = "Data Base"
    XGM_MATRIZLOG(11, 0) = "Data Distrato"
    XGM_MATRIZLOG(12, 0) = "Data Quitação"
    XGM_MATRIZLOG(13, 0) = "Data Registro Contrato"
    XGM_MATRIZLOG(14, 0) = "Data Venda"
    XGM_MATRIZLOG(15, 0) = "Valor a Vista"
    XGM_MATRIZLOG(16, 0) = "Valor Sinal"
    XGM_MATRIZLOG(17, 0) = "Valor Poupança"
    XGM_MATRIZLOG(18, 0) = "Valor Financiamento"
    XGM_MATRIZLOG(19, 0) = "Valor FGTS"
    XGM_MATRIZLOG(20, 0) = "Valor Outros"
    XGM_MATRIZLOG(21, 0) = "Valor Total"
    XGM_MATRIZLOG(22, 0) = "Tipo Deságio"
    XGM_MATRIZLOG(23, 0) = "Percentual Deságio"
    XGM_MATRIZLOG(24, 0) = "Tipo Juros Mora"
    XGM_MATRIZLOG(25, 0) = "Percentual Juros Mora"
    XGM_MATRIZLOG(26, 0) = "Tipo Multa ProRata"
    XGM_MATRIZLOG(27, 0) = "Limite Dias Multa ProRata"
    XGM_MATRIZLOG(28, 0) = "Percentual Multa ProRata"
    XGM_MATRIZLOG(29, 0) = "Tipo Correção ProRata"
    XGM_MATRIZLOG(30, 0) = "Tipo Seguro"
    XGM_MATRIZLOG(31, 0) = "Taxa Seguro"
                
    'Dados informados
    XGM_MATRIZLOG(0, Indice) = XFT_CODEMPREENDIMENTO
    XGM_MATRIZLOG(1, Indice) = XFT_CODIMOVEL
    XGM_MATRIZLOG(2, Indice) = LblCodContrato.Caption
    XGM_MATRIZLOG(3, Indice) = IIf(CboStatus.ListIndex = 1, "I", "A")
    XGM_MATRIZLOG(4, Indice) = XFI_CODCLIENTE
    XGM_MATRIZLOG(5, Indice) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
    XGM_MATRIZLOG(6, Indice) = FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil"))
    XGM_MATRIZLOG(7, Indice) = FunNulo(CboCCorrente.BoundText)
    XGM_MATRIZLOG(8, Indice) = FunNulo(CboIndexador.BoundText)
    XGM_MATRIZLOG(9, Indice) = FunNulo(CboIdentificador.BoundText)
    XGM_MATRIZLOG(10, Indice) = FunNulo(DtpDtBase.Value)
    XGM_MATRIZLOG(11, Indice) = FunNulo(DtpDistrato.Value)
    XGM_MATRIZLOG(12, Indice) = FunNulo(DtpDtQuitacao.Value)
    XGM_MATRIZLOG(13, Indice) = FunNulo(DtpRegContrato.Value)
    XGM_MATRIZLOG(14, Indice) = FunNulo(DtpDtVenda.Value)
    XGM_MATRIZLOG(15, Indice) = FunNuloVal(TxtAVista.Text)
    XGM_MATRIZLOG(16, Indice) = FunNuloVal(TxtSinal.Text)
    XGM_MATRIZLOG(17, Indice) = FunNuloVal(TxtPoupanca.Text)
    XGM_MATRIZLOG(18, Indice) = FunNuloVal(TxtFinanciamento.Text)
    XGM_MATRIZLOG(19, Indice) = FunNuloVal(TxtFgts.Text)
    XGM_MATRIZLOG(20, Indice) = FunNuloVal(TxtOutros.Text)
    XGM_MATRIZLOG(21, Indice) = FunNuloVal(TxtTotVenda.Text)
    
    If OptDesDiario.Value = True Then
      XGM_MATRIZLOG(22, Indice) = "D"
    ElseIf OptDesMensal = True Then
      XGM_MATRIZLOG(22, Indice) = "M"
    Else
      XGM_MATRIZLOG(22, Indice) = "S"
    End If
        
    XGM_MATRIZLOG(23, Indice) = FunNuloVal(TxtPercJurosFin.Text)
    
    If OptJurDiario.Value = True Then
        XGM_MATRIZLOG(24, Indice) = "D"
    ElseIf OptJurMensal.Value = True Then
        XGM_MATRIZLOG(24, Indice) = "M"
    ElseIf OptJurCheioComposto.Value = True Then
        XGM_MATRIZLOG(24, Indice) = "C"
    Else
        XGM_MATRIZLOG(24, Indice) = "S"
    End If
        
    XGM_MATRIZLOG(25, Indice) = FunNuloVal(TxtPercJurosMora.Text)
    
    If OptMultaNormal.Value = True Then
      XGM_MATRIZLOG(26, Indice) = "0"
    ElseIf OptMultaProRata.Value = True Then
      XGM_MATRIZLOG(26, Indice) = "1"
    ElseIf OptMultaCrescente.Value = True Then
      XGM_MATRIZLOG(26, Indice) = "2"
    End If
        
    XGM_MATRIZLOG(27, Indice) = FunNuloVal(TxtLimiteDias.Text)
    XGM_MATRIZLOG(28, Indice) = FunNuloVal(TxtPercMulta.Text)
    
    If ChkCorrecProRataAnterior = 1 Then
      XGM_MATRIZLOG(29, Indice) = "A"
    ElseIf ChkCorrecProRataPosterior = 1 Then
      XGM_MATRIZLOG(29, Indice) = "P"
    ElseIf ChkCorrecProRataMista = 1 Then
      XGM_MATRIZLOG(29, Indice) = "M"
    Else
      XGM_MATRIZLOG(29, Indice) = "0"
    End If
        
    If ChkSlDevedor.Value = 1 Then
        XGM_MATRIZLOG(30, Indice) = 1
    ElseIf ChkParcela.Value = 1 Then
        XGM_MATRIZLOG(30, Indice) = 2
    Else
        XGM_MATRIZLOG(30, Indice) = ""
    End If
        
    XGM_MATRIZLOG(31, Indice) = FunNuloVal(TxtTaxaSeguro.Text)
    
        
    Call subRegistraLog("TelaManContratos", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, CONTRATO_CAPI))
    
            
    Dim i As Integer
    
    ReDim XGM_MATRIZLOG(8, VFV_VETOR2.Count(2)) As Variant

    For i = 0 To VFV_VETOR2.Count(1) - 1
        
            XGM_MATRIZLOG(i, 0) = "Cód Corretor: "
            XGM_MATRIZLOG(i, Indice) = VFV_VETOR2(i, 0)
            
            XGM_MATRIZLOG(i, 0) = "Corretor: "
            XGM_MATRIZLOG(i, Indice) = VFV_VETOR2(i, 1)
            
            XGM_MATRIZLOG(i, 0) = "Valor: "
            XGM_MATRIZLOG(i, Indice) = Format(VFV_VETOR2(i, 2), "0.00")
            
            XGM_MATRIZLOG(i, 0) = "Pagamento: "
            XGM_MATRIZLOG(i, Indice) = VFV_VETOR2(i, 3)
        
    Next i
    
    
    Call subRegistraLog("TelaManContratos", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, CONTRATO_CAPI))

    
End Sub

Private Sub CmdGravar_Click()
    
    Dim XLT_STATUS As String, XLT_SQL As String
    Dim XLT_JUROS As String, XLT_DESAGIO As String, XLT_SEGURO As String
    Dim XLT_PCR As String, XLT_CONTA As String
    Dim XLO_IMOVEL As ADODB.Recordset 'Usado para recuperar o Status do imóvel
    Dim XLO_Contrato As New ADODB.Recordset  'Usado para verificar se o status do contrato está ativo
    Dim XLO_QUITACAO As New ADODB.Recordset
    Dim XLT_CORRECAOPRORATA As String
    Dim XLT_MULTA As String
    
        
    On Error GoTo RotuloErro
    
    If funChecarCampos Then
    
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
        
        Select Case CboStatus.ListIndex
            Case 1
                 XLT_STATUS = "I"
            Case Else
                 XLT_STATUS = "A"
        End Select
        
        If OptJurDiario.Value = True Then
            XLT_JUROS = "D"
        ElseIf OptJurMensal.Value = True Then
            XLT_JUROS = "M"
        ElseIf OptJurCheioComposto.Value = True Then
            XLT_JUROS = "C"
        Else
            XLT_JUROS = "S"
        End If
        
        If OptDesDiario.Value = True Then
          XLT_DESAGIO = "D"
        ElseIf OptDesMensal = True Then
          XLT_DESAGIO = "M"
        Else
          XLT_DESAGIO = "S"
        End If
        
        If OptMultaNormal.Value = True Then
          XLT_MULTA = "0"
        ElseIf OptMultaProRata.Value = True Then
          XLT_MULTA = "1"
        ElseIf OptMultaCrescente.Value = True Then
          XLT_MULTA = "2"
        End If
            
        If ChkSlDevedor.Value = 1 Then
            XLT_SEGURO = 1
        ElseIf ChkParcela.Value = 1 Then
            XLT_SEGURO = 2
        Else
            XLT_SEGURO = ""
        End If
        
        If ChkCorrecProRataAnterior = 1 Then
          XLT_CORRECAOPRORATA = "A"
        ElseIf ChkCorrecProRataPosterior = 1 Then
          XLT_CORRECAOPRORATA = "P"
        ElseIf ChkCorrecProRataMista = 1 Then ' Criado para atender à Franisa
          XLT_CORRECAOPRORATA = "M"
        Else
          XLT_CORRECAOPRORATA = "0"
        End If
        
        'Quando é chamado do Painel de Cessao o Conexao.BeginTrans já foi aberto, então pula.
        If XOrigem <> "PanCessao" Then
          Conexao.BeginTrans
        End If
        
        If Not IsNull(DtpDtQuitacao.Value) Then
          SubQOpenRecordset XLO_QUITACAO, "SELECT titu_dt_Vencimento FROM Titulos " & _
            " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & LblCodContrato.Caption & "'" & _
            " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'" & _
            " AND titu_dt_Pagamento is null", Estatico
          If Not XLO_QUITACAO.EOF Then
            MsgBox "O Contrato não pode ser quitado, pois ainda há parcelas pendentes!", vbCritical, "ATENÇÃO"
            Conexao.RollbackTrans
            XLO_QUITACAO.Close
            Set XLO_QUITACAO = Nothing
            Exit Sub
          End If
        End If
        
        If XInserir = "I" Then
            
            'Quando o status do contrato inserido é ATIVO se não tiver uma reserva
            'então soma 1 ao número de unidades vendidas do empreendimento e o staus do imóvel passa a ser Vendido
            If XLT_STATUS = "A" Then
            
              SubQOpenRecordset XLO_Contrato, "SELECT cont_tx_status, cont_cd_Contrato FROM Contratos WHERE " & _
                "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
                "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_tx_Status ='R'", "Estatico"
            
              If Not (XLO_Contrato.BOF And XLO_Contrato.EOF) Then
                MsgBox "Esse imóvel já está reservado.", vbCritical, "ATENÇÃO !"
                XLO_Contrato.Close 'Inserido em 13/10/08 - Patrícia
                Set XLO_Contrato = Nothing
                Conexao.RollbackTrans
                Exit Sub
              Else
               Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas +1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
               Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='V' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
               XLO_Contrato.Close 'Inserido em 13/10/08 - Patrícia
               Set XLO_Contrato = Nothing
              End If
              
            End If
            
            XLT_SQL = _
                "INSERT INTO " & _
                "contratos " & _
                "(" & _
                "empr_cd_Empresa, "
            XLT_SQL = XLT_SQL & _
                "cont_cd_Contrato, " & _
                "imov_cd_imovel, " & _
                "empd_cd_empreendimento, " & _
                "cere_cd_Pcr, " & _
                "plco_cd_Conta, " & _
                "iden_cd_Projeto, "
            XLT_SQL = XLT_SQL & _
                "moed_cd_CodMoeda, " & _
                "cont_dt_Base, " & _
                "cont_dt_Venda, " & _
                "cont_dt_Quitacao, " & _
                "cont_dt_RegistroContrato, " & _
                "cont_vl_AVista, " & _
                "cont_vl_Sinal, " & _
                "cont_vl_Poupanca, " & _
                "cont_vl_Financiamento, "
            XLT_SQL = XLT_SQL & _
                "cont_vl_Total, " & _
                "cont_vl_Fgts, " & _
                "cont_vl_Outros, " & _
                "cont_tx_TipoJurosMora, " & _
                "cont_tx_Desagio, " & _
                "cont_vl_TaxaSegVida, " & _
                "cont_nr_TipoSeguro, " & _
                "cont_tx_MultaProRata, " & _
                "cont_tx_CorrecProRata, " & _
                "cont_tx_status, " & _
                "focl_cd_FornCli, " & _
                "cont_vl_PercJurosFin, "
            XLT_SQL = XLT_SQL & _
                "cont_vl_PercJurosMulta, " & _
                "cont_vl_PercJurosMora, " & _
                "cont_cd_ContratoOrigem, " & _
                "cont_dt_Distrato, " & _
                "coco_cd_Codigo, " & _
                "cont_nr_LimiteDias "
            XLT_SQL = XLT_SQL & _
                ") VALUES (" & _
                PCodEmpresa & ",'" & _
                LblCodContrato.Caption & "','" & _
                XFT_CODIMOVEL & "','" & _
                XFT_CODEMPREENDIMENTO & "', " & _
                "" & FunNuloBanco(XLT_PCR) & "," _
                & FunNuloBanco(XLT_CONTA) & "," & _
                FunNuloBancoVal(CboIdentificador.BoundText) & "," & _
                FunNuloBancoVal(CboIndexador.BoundText) & ", " & _
                "" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ","
            XLT_SQL = XLT_SQL & _
                FunNuloData(DtpDtVenda.Value, NomeSgbd) & "," & _
                FunNuloData(DtpDtQuitacao.Value, NomeSgbd) & ", " & _
                "" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & _
                FunNuloVal(FunTrataFloat(TxtAVista.Text)) & ", " & _
                FunNuloVal(FunTrataFloat(TxtSinal.Text)) & ", " & _
                "" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & ", " & _
                FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & _
                FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & _
                FunNuloVal(FunTrataFloat(TxtFgts.Text)) & "," & _
                FunNuloVal(FunTrataFloat(TxtOutros.Text)) & "," & _
                "'" & FunNulo(XLT_JUROS) & "','" & _
                FunNulo(XLT_DESAGIO) & "', " & _
                FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & ", '" & _
                XLT_SEGURO & "', " & _
                "'" & XLT_MULTA & "', '" & _
                XLT_CORRECAOPRORATA & "','" & _
                XLT_STATUS & "'," & _
                XFI_CODCLIENTE & ", " & _
                "" & FunNuloVal(FunTrataFloat(TxtPercJurosFin.Text)) & ", " & _
                FunNuloVal(FunTrataFloat(TxtPercMulta.Text)) & "," & _
                FunNuloVal(FunTrataFloat(TxtPercJurosMora.Text)) & "," & _
                "'" & LblCodContratoOrigem.Caption & "',"
            XLT_SQL = XLT_SQL & _
                FunNuloData(DtpDistrato.Value, NomeSgbd) & _
                "," & FunNuloBancoVal(CboCCorrente.BoundText) & _
                "," & FunNuloVal(TxtLimiteDias.Text)
            XLT_SQL = XLT_SQL & _
                ")"
            
            Conexao.Execute XLT_SQL
            
            'Atualiza o último contrato na tabela de imóveis
            Conexao.Execute "UPDATE Imoveis SET imov_cd_UltimoContrato='" & LblCodContrato.Caption & "' " & _
            "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            
            'Grava Log
            GravaLog CADASTRAR, "2"
            
        'ALTERAÇÃO
        Else
            'Ao alterar um contrato verificar se ele está sendo alterado XFT_STATUS = "A" ou ativado XFT_STATUS = "R"
            If XLT_STATUS = "A" Then
              If XFT_STATUS = "R" Then
                'De Reservado para ativo só altera o status do imóvel
                Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='V' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
              End If
            End If
            
            XLT_SQL = "UPDATE Contratos SET cere_cd_Pcr = " & FunNuloBanco(XLT_PCR) & "," & _
                      "plco_cd_Conta = " & FunNuloBanco(XLT_CONTA) & "," & _
                      "iden_cd_Projeto = " & FunNuloBancoVal(CboIdentificador.BoundText) & "," & _
                      "moed_cd_CodMoeda = " & FunNuloBancoVal(CboIndexador.BoundText) & "," & _
                      "cont_dt_Base = " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ",cont_dt_Venda= " & FunNuloData(DtpDtVenda.Value, NomeSgbd) & "," & _
                      "cont_dt_Quitacao = " & FunNuloData(DtpDtQuitacao.Value, NomeSgbd) & ",cont_dt_RegistroContrato=" & FunNuloData(DtpRegContrato.Value, NomeSgbd) & "," & _
                      "cont_vl_AVista = " & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & "," & _
                      "cont_vl_Sinal = " & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & "," & _
                      "cont_vl_Poupanca = " & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & "," & _
                      "cont_vl_Financiamento = " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & _
                      "cont_vl_Total = " & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & _
                      "cont_vl_Fgts = " & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & ",cont_vl_Outros=" & FunNuloVal(FunTrataFloat(TxtOutros.Text)) & "," & _
                      "cont_tx_TipoJurosMora = '" & XLT_JUROS & "'"
              
            XLT_SQL = XLT_SQL & "," & _
                      "cont_tx_Desagio = '" & XLT_DESAGIO & "'," & _
                      "cont_vl_TaxaSegVida = " & FunNuloVal(FunTrataFloat(TxtTaxaSeguro.Text)) & "," & _
                      "cont_nr_TipoSeguro = '" & XLT_SEGURO & "'," & _
                      "cont_tx_MultaProRata = '" & XLT_MULTA & "'," & _
                      "cont_tx_CorrecProRata = '" & XLT_CORRECAOPRORATA & "'," & _
                      "cont_tx_status = '" & XLT_STATUS & "'," & _
                      "focl_cd_forncli = " & XFI_CODCLIENTE & "," & _
                      "cont_vl_PercJurosFin = " & FunNuloVal(FunTrataFloat(TxtPercJurosFin.Text)) & "," & _
                      "cont_vl_PercJurosMulta = " & FunNuloVal(FunTrataFloat(TxtPercMulta.Text)) & "," & _
                      "cont_vl_PercJurosMora = " & FunNuloVal(FunTrataFloat(TxtPercJurosMora.Text)) & "," & _
                      "cont_cd_ContratoOrigem = '" & LblCodContratoOrigem.Caption & "', " & _
                      "cont_dt_Distrato =" & FunNuloData(DtpDistrato.Value, NomeSgbd) & ", " & _
                      "coco_cd_Codigo = " & FunNuloBancoVal(CboCCorrente.BoundText) & ", " & _
                      "cont_nr_LimiteDias = " & FunNuloVal(TxtLimiteDias.Text)
            XLT_SQL = XLT_SQL & _
                      " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & LblCodContrato.Caption & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            
            Conexao.Execute XLT_SQL
            
            'Grava Log
            GravaLog ALTERAR, "3"
        End If
        
        'Compradores
        If XFB_ENTROUCOMP Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Compradores WHERE empr_cd_Empresa = " & PCodEmpresa & _
                "AND empd_cd_Empreendimento = '" & XFT_CODEMPREENDIMENTO & "' AND " & _
                "imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato = '" & XFT_CODCONTRATO & "'"
            End If
            If VFV_VETOR1(0, 0) <> Empty Then
                subGravaVetor "Compradores", "focl_cd_FornCli, imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa,cont_cd_Contrato", Array("N", "T", "T", "T", "N"), 1, Array(0), "'" & XFT_CODIMOVEL & "', '" & XFT_CODEMPREENDIMENTO & "', " & PCodEmpresa & ",'" & LblCodContrato.Caption & "'", VFV_VETOR1, TDBGrid1
            End If
        End If
        
        'Fiadores
        If XFB_ENTROUFIAD Then
            If XInserir = "A" Then
                Conexao.Execute "DELETE from FiadoresContratos WHERE empr_cd_Empresa=" & PCodEmpresa & _
                "AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
                "imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "'"
            End If
            If VFV_VETOR3(0, 0) <> Empty Then
                subGravaVetor "FiadoresContratos", "fiad_cd_Fiador, imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa,cont_cd_Contrato", Array("N", "T", "T", "T", "N"), 1, Array(0), "'" & XFT_CODIMOVEL & "', '" & XFT_CODEMPREENDIMENTO & "', " & PCodEmpresa & ",'" & LblCodContrato.Caption & "'", VFV_VETOR3, TDBGrid3
            End If
        End If
        
        'Corretores
        If XFB_ENTROUCORR Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Comissoes WHERE empr_cd_Empresa=" & PCodEmpresa & _
                "AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
                "imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "'"
            End If
            If VFV_VETOR2(0, 0) <> Empty Then
                subGravaVetor "Comissoes", "corr_cd_Corretor, comi_vl_Comissao, comi_dt_Pagamento, imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa,cont_cd_Contrato", Array("N", "N", "D", "T", "T", "T", "N"), 3, Array(0, 2, 3), "'" & XFT_CODIMOVEL & "', '" & XFT_CODEMPREENDIMENTO & "', " & PCodEmpresa & ",'" & LblCodContrato.Caption & "'", VFV_VETOR2, TDBGrid2
            End If
        End If
                
        'Quando é chamado do Painel de Cessao o Conexao.CommitTrans
         'já foi aberto, então pula.
        If XOrigem <> "PanCessao" Then
          Conexao.CommitTrans
        Else
          Chave = LblCodContrato.Caption
        End If
        
        Call CmdDesistir_Click
    
    End If
    Exit Sub
    
RotuloErro:
  funTrataErros (ComMensagem)
  Resume Next
End Sub

Private Sub CmdInsComprador_Click()
  If CboComprador.Text <> "" Then
    DatComprador.Recordset.bookmark = CboComprador.SelectedItem
    
    If funExisteChave(1, Array("0"), Array(DatComprador.Recordset.Fields("focl_cd_forncli")), VFV_VETOR1, TDBGrid1) Then
        MsgBox "Esse Cliente já foi cadastrado", vbCritical, "ATENÇÃO !"
        Exit Sub
    End If
    
    subInsereRegistroVetor Array(CboComprador.BoundText, CboComprador.Text, ""), VFV_VETOR1, TDBGrid1
    If VFV_VETOR1.UpperBound(1) = 0 Then
       Call CmdTitular_Click
    End If
    
    If Not CmdExcComprador.Enabled Then CmdExcComprador.Enabled = True
  End If
End Sub

Private Sub CmdInsCorretor_Click()
    
    CboCorretor.Text = ""
    TxtPercComissao.Text = ""
    TxtComissao.Text = ""
    DtpPagamento.Value = ""
    XFT_INSCORRETOR = "I"
    PanCorretor.Visible = True
    PanCorretor.Enabled = True
    
    LblValorDoImovel.Caption = IIf(TxtTotVenda.Text <> "", TxtTotVenda.Text, TxtAVista.Text)
    
End Sub

Private Sub CmdInsCorretor_GotFocus()
    TabAuxiliar.Tab = 4
End Sub

Private Sub CmdInsFiador_Click()
    If CboFiador.Text <> "" Then
        DatFiador.Recordset.bookmark = CboFiador.SelectedItem
        
        If funExisteChave(1, Array("0"), Array(DatFiador.Recordset.Fields("fiad_cd_fiador")), VFV_VETOR3, TDBGrid3) Then
            MsgBox "Esse Fiador já foi cadastrado", vbCritical, "ATENÇÃO !"
            Exit Sub
        End If
        subInsereRegistroVetor Array(CboFiador.BoundText, CboFiador.Text), VFV_VETOR3, TDBGrid3
        If Not CmdExcFiador.Enabled Then CmdExcFiador.Enabled = True
    End If
End Sub

Private Sub CmdLimparCCorrente_Click()
  CboCCorrente.BoundText = ""
End Sub

Private Sub CmdLimparIdentif_Click()
    CboIdentificador.BoundText = ""
End Sub

Private Sub CboIdentificador_LostFocus()
'    If TabAuxiliar.TabEnabled(1) = True Then
'        TabAuxiliar.Tab = 1
'    End If
End Sub

Private Sub CmdLimparIndexador_Click()
  CboIndexador.BoundText = ""
End Sub

Private Sub CmdLimparplanoContas_Click()
  CboPlanoContas.BoundText = ""
  MskConta.Text = ""
End Sub

Private Sub CmdLimparPcr_Click()
  CboPcr.BoundText = ""
  MskPcr.Text = ""
End Sub

Private Sub CmdSairCorretor_Click()
    PanCorretor.Enabled = False
    PanCorretor.Visible = False
End Sub

Private Sub CmdTitular_Click()
    Dim XLI_POS As Integer
    Dim XLI_POSPADRAO As Integer
    
    If VFV_VETOR1(0, 0) <> Empty Then
        XLI_POS = TDBGrid1.bookmark
        TDBGrid1.MoveFirst
        Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Titular").Value = "S" Then
                XLI_POSPADRAO = TDBGrid1.bookmark
                If XLI_POSPADRAO = XLI_POS Then
                    MsgBox "Este Cliente já é o Titular!", vbCritical, "ATENÇÃO"
                    XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
                    Exit Sub
                Else
                    If MsgBox("Já existe um Cliente Titular. Deseja alterar?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma alteração
                        TDBGrid1.Columns("Titular").Text = "N"
                        TDBGrid1.Update
                        TDBGrid1.bookmark = XLI_POS
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
        TDBGrid1.bookmark = XLI_POS
        TDBGrid1.Columns("Titular").Text = "S"
        TDBGrid1.Update
        XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
    End If

End Sub

Private Function funUltimoContrato() As String
  Dim XLO_ULTIMOCONTRATO As New ADODB.Recordset

   'Preenche o numero do contrato automaticamente
   'fazendo uma busca no contrato com codigo de maior valor
  SubQOpenRecordset XLO_ULTIMOCONTRATO, "SELECT MAX(cont_cd_Contrato) as Maior FROM Contratos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'", "Estatico"
  If IsNull(XLO_ULTIMOCONTRATO!Maior) Then
    funUltimoContrato = "00"
  Else
    funUltimoContrato = XLO_ULTIMOCONTRATO!Maior
  End If
  
  XLO_ULTIMOCONTRATO.Close
  Set XLO_ULTIMOCONTRATO = Nothing

End Function

Private Sub Form_Activate()
    Set Formulario = FrmManContratos
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String 'Armazena comaandos SQL
    Dim XLO_IMOVEL As Object 'Usado para preencher os dados do imóvel
    Dim XLO_ULTIMOCONTRATO As New ADODB.Recordset 'Pega o último contrato desta imóvel
    
    VFV_VETOR2.ReDim 0, 0, 0, TDBGrid2.Columns.Count - 1
      
    If XGB_PQUALIFIN = False Then
      FraFinContabil.Enabled = False
    End If
    
    Set Formulario = FrmManContratos
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If
    
    'desabilita botao de compradores
    CmdInsComprador.Enabled = False
    CmdInsFiador.Enabled = False
    
    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Mid$(Chave, 6, 4)
    XFT_CODCONTRATO = Right$(Chave, 2)
    
    XFB_ENTROUCOMP = False
    XFB_ENTROUFIAD = False
    XFB_ENTROUCORR = False
    
    subConectarControleDadosNV DatComprador, "SELECT focl_cd_forncli, focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
    subConectarControleDadosNV DatCorretor, "SELECT corr_cd_Corretor,corr_vl_PercentComissao,corr_tx_Nome + space(32 - len(corr_tx_Nome)) + '- ' + convert(varchar,corr_vl_PercentComissao) as Corretor FROM Corretores ORDER BY corr_tx_nome", Estatico
    subConectarControleDadosNV DatFiador, "SELECT fiad_cd_Fiador, fiad_tx_Nome FROM Fiadores ORDER BY fiad_tx_Nome", Estatico
    subConectarControleDadosNV DatIndexador, "SELECT * FROM Moedas ORDER BY moed_tx_Descricao", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico
            
    'Preenche dados
    LblCodEmpreendimento.Caption = XFT_CODEMPREENDIMENTO
    LblCodImovel.Caption = XFT_CODIMOVEL
    
    'Quando o formulário é chamado pelo Painel de Cessão
      'o contrato a ser criado só pode se ATIVO
    If XOrigem = "PanCessao" Then
      XInserir = "I"
    End If
    
    SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM ConsCapImoveis WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'", "Estatico"
    If Not (XLO_IMOVEL.BOF Or XLO_IMOVEL.EOF) Then
        LblDescEmpreendimento.Caption = XLO_IMOVEL!Empreendimento
        LblDescTipoImovel.Caption = XLO_IMOVEL!Tipo
        LblDescGaragem.Caption = XLO_IMOVEL!imov_tx_Garagem
        LblDescArea.Caption = XLO_IMOVEL!imov_nr_Area
        LblDescAndar.Caption = FunNuloVal(XLO_IMOVEL!imov_nr_Andar)
        
        If XInserir = "I" Then
            If XOrigem <> "PanCessao" Then
                Call subPrepararInsercao
                
                If Not DatConta.Recordset.EOF Then
                    DatConta.Recordset.MoveFirst
                    If Not IsNull(XLO_IMOVEL!plco_cd_Conta) Then
                        DatConta.Recordset.Find "plco_cd_Conta =" & XLO_IMOVEL!plco_cd_Conta
                        MskConta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
                        CboPlanoContas.BoundText = MskConta.Text
                    Else
                        MskConta.Text = ""
                        CboPlanoContas.BoundText = ""
                    End If
                End If
                If Not DatPcr.Recordset.EOF Then
                    DatPcr.Recordset.MoveFirst
                    If XInserir <> "I" Then
                        If Not IsNull(XLO_IMOVEL!cere_cd_Pcr) Then
                            DatPcr.Recordset.Find "cere_cd_pcr =" & XLO_IMOVEL!cere_cd_Pcr
                            MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
                            CboPcr.BoundText = MskPcr.Text
                        Else
                          MskPcr.Text = ""
                          CboPcr.BoundText = ""
                        End If
                    End If
                End If
                If Not DatContaCorrente.Recordset.EOF Then
                    DatContaCorrente.Recordset.MoveFirst
                    If Not IsNull(XLO_IMOVEL!coco_cd_codigo) Then
                        DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & XLO_IMOVEL!coco_cd_codigo
                        CboCCorrente.BoundText = FunNulo(XLO_IMOVEL("coco_cd_Codigo"))
                    Else
                      CboCCorrente.BoundText = ""
                    End If
                End If
                
                TxtAVista.Text = Format(XLO_IMOVEL!imov_vl_Avista, "standard")
                LblCodContrato.Caption = Format(CInt(funUltimoContrato) + 1, "00")
            Else
                'Preenche o contrato de origem
                If XOrigem = "PanCessao" Then
                    LblCodContrato.Caption = Format(CInt(funUltimoContrato) + 1, "00")
                    LblCodContratoOrigem.Caption = funUltimoContrato
                    subPreencheCampos
                    CboStatus.ListIndex = 0
                    CboStatus.Enabled = False
                End If
            End If
        ElseIf XInserir = "A" Then
            Me.Caption = "Alteração de Contratos"
            LblCodContrato.Caption = XFT_CODCONTRATO
            CboStatus.Enabled = False
            subPreencheCampos
            'Preenche informações para log
            PreencheInformacoesOriginaisLog
        ElseIf XInserir = "C" Then
            Me.Caption = "Consulta de Contratos"
            LblCodContrato.Caption = XFT_CODCONTRATO
            CboStatus.Enabled = False
            subPreencheCampos
            SubDesabHabCampos (False)
        End If
    End If
    
    'Compradores
    XLT_SQL = "SELECT focl_cd_FornCli, focl_tx_RazaoSocial, focl_tx_RazaoSocial FROM ConsCAPCompradores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & LblCodContrato.Caption & "'"
    
    subCarregaVetorCompradores XLT_SQL, Array("focl_cd_FornCli", "focl_tx_RazaoSocial", "focl_tx_RazaoSocial"), VFV_VETOR1, TDBGrid1
    XFB_ENTROUCOMP = True
    
    XLO_IMOVEL.Close
    Set XLO_IMOVEL = Nothing
    
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
    subManutencaoJanelasAtivas "R", "FrmManContratos"
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

Private Sub OptJurCheioComposto_Click()
   
  'Quando o juros é o cheio composto (sobre o valor histório) a multa só pode ser a normal
  OptMultaNormal.Value = True
  OptMultaProRata.Enabled = False
  OptMultaCrescente.Enabled = False
    
End Sub

Private Sub OptJurCheioSimples_Click()
  
  OptMultaProRata.Enabled = True
  OptMultaCrescente.Enabled = True
  
End Sub

Private Sub OptJurDiario_Click()
  
  OptMultaProRata.Enabled = True
  OptMultaCrescente.Enabled = True
  
End Sub

Private Sub OptJurMensal_Click()
   
  OptMultaProRata.Enabled = True
  OptMultaCrescente.Enabled = True
    
End Sub

Private Sub OptMultaCrescente_Click()
  
  TxtLimiteDias.Enabled = True
  LblLimiteDias.Enabled = True
  
End Sub

Private Sub OptMultaNormal_Click()
  
  TxtLimiteDias.Enabled = False
  LblLimiteDias.Enabled = False
  TxtLimiteDias.Text = ""
  
End Sub

Private Sub OptMultaProRata_Click()
  
  TxtLimiteDias.Enabled = False
  LblLimiteDias.Enabled = False
  TxtLimiteDias.Text = ""
  
End Sub

Private Sub TabAuxiliar_Click(PreviousTab As Integer)
    Dim XLT_SQL As String
    
    If TabAuxiliar.Tab = 0 And Not XFB_ENTROUCOMP Then
        'Compradores
        XLT_SQL = "SELECT focl_cd_FornCli, focl_tx_RazaoSocial, focl_tx_RazaoSocial FROM ConsCAPCompradores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & LblCodContrato.Caption & "'"
        subCarregaVetorCompradores XLT_SQL, Array("focl_cd_FornCli", "focl_tx_RazaoSocial", "focl_tx_RazaoSocial"), VFV_VETOR1, TDBGrid1
        XFB_ENTROUCOMP = True
    ElseIf TabAuxiliar.Tab = 1 And Not XFB_ENTROUFIAD Then
        'Fiadores
        XLT_SQL = "SELECT fiad_cd_Fiador, fiad_tx_Nome FROM ConsCAPFiadoresContratos WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & LblCodContrato.Caption & "'"
        subCarregaVetor XLT_SQL, Array("fiad_cd_Fiador", "fiad_tx_Nome"), VFV_VETOR3, TDBGrid3
        XFB_ENTROUFIAD = True
        If IsEmpty(VFV_VETOR3(0, 0)) Then
          CmdExcFiador.Enabled = False
        End If
    ElseIf TabAuxiliar.Tab = 4 And Not XFB_ENTROUCORR Then
        'Corretores
        XLT_SQL = "SELECT * FROM ConsCAPCorretores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & LblCodContrato.Caption & "'"
        subCarregaVetor XLT_SQL, Array("corr_cd_Corretor", "corr_tx_Nome", "comi_vl_Comissao", "comi_dt_Pagamento"), VFV_VETOR2, TDBGrid2
        XFB_ENTROUCORR = True
        If IsEmpty(VFV_VETOR2(0, 0)) Then
          CmdAltCorretor.Enabled = False
          CmdExcCorretor.Enabled = False
        End If
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
            If XFI_CODCLIENTE = XLO_COMPRADOR!focl_cd_FornCli And Grid.Columns.Count = 3 Then
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

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    If TDBGrid1.Columns("Titular").CellText(bookmark) = "S" Then
        CellStyle.ForeColor = vbRed
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
        TxtSinal.Text = ""
        TxtPoupanca.Text = ""
        TxtFinanciamento.Text = ""
        TxtFgts.Text = ""
        TxtOutros.Text = ""
        TxtTotVenda.Text = ""
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
    
    If CDbl(LblValorDoImovel.Caption) > 0 Then
        TxtComissao.Text = Format$(TxtComissao.Text, "##,##0.00")
        TxtPercComissao.Text = Format(CDbl(TxtComissao.Text) * 100 / (CDbl(LblValorDoImovel.Caption)), "standard")
    Else
        TxtComissao.Text = Format(0, "standard")
        TxtPercComissao.Text = Format(0, "standard")
    End If
    
End Sub

Private Sub TxtPercComissao_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
  End If
  
End Sub

Private Sub TxtPercComissao_LostFocus()
  
  If CDbl(LblValorDoImovel.Caption) > 0 And TxtPercComissao.Text <> "" Then
    TxtPercComissao.Text = Format(TxtPercComissao.Text, "standard")
    TxtComissao.Text = Format(CDbl(LblValorDoImovel.Caption) * CDbl(TxtPercComissao.Text) / 100, "standard")
  Else
    TxtComissao.Text = Format(0, "standard")
    TxtPercComissao.Text = Format(0, "standard")
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
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtFgts.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtFgts.Text = ""
        End If
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
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtFinanciamento.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtFinanciamento.Text = ""
        End If
    End If
End Sub

Private Sub TxtOutros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
  End If
End Sub

Private Sub TxtOutros_LostFocus()
   If IsNumeric(TxtOutros.Text) Then
        TxtOutros.Text = Format$(TxtOutros.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtOutros.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtOutros.Text = ""
        End If
    End If
End Sub

Private Sub TxtPercJurosFin_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtPercJurosFin_LostFocus()
   
   If IsNumeric(TxtPercJurosFin.Text) Then
     TxtPercJurosFin.Text = Format$(TxtPercJurosFin.Text, "##,##0.0000")
   Else
     TxtPercJurosFin.Text = ""
   End If
   
End Sub

Private Sub TxtPercJurosMora_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtPercJurosMora_LostFocus()
   
   If IsNumeric(TxtPercJurosMora.Text) Then
     TxtPercJurosMora.Text = Format$(TxtPercJurosMora.Text, "##,##0.0000")
   Else
     TxtPercJurosMora.Text = ""
   End If
   
End Sub

Private Sub TxtPercMulta_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtPercMulta_LostFocus()
   
   If IsNumeric(TxtPercMulta.Text) Then
     TxtPercMulta.Text = Format$(TxtPercMulta.Text, "##,##0.0000")
   Else
     TxtPercMulta.Text = ""
   End If
   
End Sub

Private Sub TxtTaxaSeguro_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtTaxaSeguro_LostFocus()
   
   If IsNumeric(TxtTaxaSeguro.Text) Then
     TxtTaxaSeguro.Text = Format$(TxtTaxaSeguro.Text, "##,##0.00000")
   Else
     TxtTaxaSeguro.Text = ""
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
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtPoupanca.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtPoupanca.Text = ""
        End If
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
        TxtTotVenda.Text = Format(CDbl(TxtSinal.Text) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtSinal.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtSinal.Text = ""
        End If
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
