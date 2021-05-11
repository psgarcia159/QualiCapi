VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinPlanos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Planos de Pagamentos"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   1335
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9390
   Begin VB.Frame FraPrincipal 
      Height          =   5040
      Left            =   45
      TabIndex        =   30
      Top             =   -60
      Width           =   9345
      Begin VB.Frame FraPlano1 
         Height          =   780
         Left            =   120
         TabIndex        =   48
         Top             =   1020
         Width           =   9135
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmPlanos.frx":0000
            Left            =   4860
            List            =   "FrmPlanos.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   1875
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmPlanos.frx":0051
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   120
            Width           =   2430
            _ExtentX        =   4286
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
            Left            =   1680
            Top             =   120
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
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
         Begin VB.Label LblVlCadastrado 
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   4800
            TabIndex        =   78
            Top             =   495
            Width           =   1155
         End
         Begin VB.Label LblValorCadastrado 
            Caption         =   "Valor Cadastrado:"
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
            Left            =   3180
            TabIndex        =   77
            Top             =   495
            Width           =   1575
         End
         Begin VB.Label LblVlCadastrar 
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   7860
            TabIndex        =   76
            Top             =   495
            Width           =   1155
         End
         Begin VB.Label LblValorCadastrar 
            Caption         =   "Valor a Cadastrar:"
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
            Left            =   6120
            TabIndex        =   75
            Top             =   495
            Width           =   1575
         End
         Begin VB.Label LblVlNatureza 
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1800
            TabIndex        =   74
            Top             =   495
            Width           =   1155
         End
         Begin VB.Label LblValorNatureza 
            Caption         =   "Valor da Natureza:"
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
            TabIndex        =   73
            Top             =   495
            Width           =   1605
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
            Height          =   210
            Left            =   4035
            TabIndex        =   70
            Top             =   195
            Width           =   855
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
            Left            =   1020
            TabIndex        =   54
            Top             =   180
            Width           =   435
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
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   180
            Width           =   570
         End
         Begin VB.Label LblPlano2 
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
            Height          =   225
            Left            =   720
            TabIndex        =   52
            Top             =   180
            Width           =   285
         End
         Begin VB.Label LblPeriodicidade 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   6840
            TabIndex        =   51
            Top             =   180
            Width           =   1260
         End
         Begin VB.Label LblMes 
            Alignment       =   1  'Right Justify
            Caption         =   "Mês(es)"
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
            Left            =   8340
            TabIndex        =   50
            Top             =   180
            Width           =   690
         End
         Begin VB.Label LblPeriodicidade2 
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
            Left            =   8100
            TabIndex        =   49
            Top             =   180
            Width           =   255
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   3135
         Left            =   120
         TabIndex        =   31
         Top             =   1860
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5530
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Dados Financeiros"
         TabPicture(0)   =   "FrmPlanos.frx":006C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblObs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FraVlIndexador"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FraPlano2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CmdLimparObs"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DatObs"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CboObs"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "FraPlano3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Dados Contábeis"
         TabPicture(1)   =   "FrmPlanos.frx":0088
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DatIdentificador"
         Tab(1).Control(1)=   "DatContaCorrente"
         Tab(1).Control(2)=   "DatConta"
         Tab(1).Control(3)=   "DatPcr"
         Tab(1).Control(4)=   "FraConta"
         Tab(1).ControlCount=   5
         Begin VB.Frame FraConta 
            Height          =   1770
            Left            =   -74850
            TabIndex        =   43
            Top             =   360
            Width           =   8850
            Begin VB.TextBox txtContratoBanco 
               Height          =   315
               Left            =   7800
               MaxLength       =   3
               TabIndex        =   80
               Top             =   1380
               Width           =   735
            End
            Begin MSMask.MaskEdBox Mskconta 
               Height          =   330
               Left            =   1470
               TabIndex        =   18
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
               Bindings        =   "FrmPlanos.frx":00A4
               Height          =   330
               Left            =   2370
               TabIndex        =   19
               Top             =   180
               Width           =   6000
               _ExtentX        =   10583
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
               Left            =   1470
               TabIndex        =   21
               Top             =   570
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmPlanos.frx":00BB
               Height          =   330
               Left            =   2370
               TabIndex        =   22
               Top             =   570
               Width           =   6000
               _ExtentX        =   10583
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
               Bindings        =   "FrmPlanos.frx":00D0
               Height          =   330
               Left            =   1470
               TabIndex        =   24
               Top             =   975
               Width           =   6885
               _ExtentX        =   12144
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
               Left            =   8385
               TabIndex        =   25
               Top             =   975
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":00EF
               Picture         =   "FrmPlanos.frx":0541
            End
            Begin Threed.SSCommand CmdLimparPcr 
               Height          =   315
               Left            =   8400
               TabIndex        =   23
               Top             =   570
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":0653
               Picture         =   "FrmPlanos.frx":0AA5
            End
            Begin Threed.SSCommand CmdLimparplanoContas 
               Height          =   315
               Left            =   8400
               TabIndex        =   20
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":0BB7
               Picture         =   "FrmPlanos.frx":1009
            End
            Begin MSDataListLib.DataCombo CboIdentificador 
               Bindings        =   "FrmPlanos.frx":111B
               Height          =   315
               Left            =   1470
               TabIndex        =   26
               Top             =   1380
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
               Left            =   5070
               TabIndex        =   27
               Top             =   1380
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":113A
               Picture         =   "FrmPlanos.frx":158C
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
               TabIndex        =   81
               Top             =   1425
               Width           =   2295
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
               Left            =   285
               TabIndex        =   79
               Top             =   1425
               Width           =   1155
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
               Left            =   75
               TabIndex        =   46
               Top             =   1035
               Width           =   1365
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
               Left            =   180
               TabIndex        =   45
               Top             =   615
               Width           =   1260
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
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   1320
            End
         End
         Begin VB.Frame FraPlano3 
            Height          =   585
            Left            =   120
            TabIndex        =   32
            Top             =   345
            Width           =   8835
            Begin VB.TextBox TxtParcelas 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1620
               MaxLength       =   3
               TabIndex        =   5
               Top             =   180
               Width           =   570
            End
            Begin MSComCtl2.DTPicker DtpVencimento 
               Height          =   315
               Left            =   7200
               TabIndex        =   7
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   74448899
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtBase 
               Height          =   315
               Left            =   3780
               TabIndex        =   6
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   74448899
               CurrentDate     =   37589
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
               Left            =   2700
               TabIndex        =   42
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label LblVencimnto 
               Alignment       =   1  'Right Justify
               Caption         =   "Venc. 1ª Parcela:"
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
               TabIndex        =   41
               Top             =   240
               Width           =   1560
            End
            Begin VB.Label LblParcelas 
               Alignment       =   1  'Right Justify
               Caption         =   "Nº de Parcelas:"
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
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   1410
            End
         End
         Begin MSAdodcLib.Adodc DatPcr 
            Height          =   330
            Left            =   -71340
            Top             =   2280
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
            Left            =   -74400
            Top             =   2280
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
            Left            =   -68880
            Top             =   2280
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
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
            Bindings        =   "FrmPlanos.frx":169E
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   2730
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
            Left            =   7080
            Top             =   2715
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
            Left            =   6180
            TabIndex        =   17
            Top             =   2730
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmPlanos.frx":16B3
            Picture         =   "FrmPlanos.frx":1B05
         End
         Begin MSAdodcLib.Adodc DatIdentificador 
            Height          =   330
            Left            =   -75000
            Top             =   2640
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
         Begin VB.Frame FraPlano2 
            Height          =   915
            Left            =   120
            TabIndex        =   33
            Top             =   915
            Width           =   8835
            Begin VB.TextBox TxtValorReal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6480
               TabIndex        =   11
               Top             =   540
               Width           =   1515
            End
            Begin VB.TextBox TxtJurosFin 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   8
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtMulta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6480
               TabIndex        =   9
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtJuros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   10
               Top             =   540
               Width           =   960
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
               Left            =   8100
               TabIndex        =   63
               Top             =   600
               Width           =   315
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
               Left            =   4020
               TabIndex        =   62
               Top             =   600
               Width           =   2445
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
               TabIndex        =   39
               Top             =   240
               Width           =   390
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
               Left            =   360
               TabIndex        =   38
               Top             =   240
               Width           =   1755
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
               Left            =   7620
               TabIndex        =   37
               Top             =   240
               Visible         =   0   'False
               Width           =   780
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
               Left            =   5670
               TabIndex        =   36
               Top             =   240
               Width           =   795
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
               Left            =   480
               TabIndex        =   35
               Top             =   600
               Width           =   1620
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
               TabIndex        =   34
               Top             =   600
               Width           =   735
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
            Height          =   870
            Left            =   120
            TabIndex        =   55
            Top             =   1815
            Width           =   8835
            Begin MSDataListLib.DataCombo CboMoeda2 
               Bindings        =   "FrmPlanos.frx":1C17
               Height          =   315
               Left            =   6060
               TabIndex        =   14
               Top             =   180
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboMoeda2"
            End
            Begin MSDataListLib.DataCombo CboMoeda1 
               Bindings        =   "FrmPlanos.frx":1C2F
               Height          =   315
               Left            =   1800
               TabIndex        =   12
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
               TabIndex        =   13
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":1C47
               Picture         =   "FrmPlanos.frx":2099
            End
            Begin Threed.SSCommand CmdLimparIndex2 
               Height          =   315
               Left            =   7860
               TabIndex        =   15
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmPlanos.frx":21AB
               Picture         =   "FrmPlanos.frx":25FD
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
               TabIndex        =   72
               Top             =   600
               Width           =   1695
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
               Height          =   210
               Left            =   1800
               TabIndex        =   71
               Top             =   600
               Width           =   1695
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
               Height          =   195
               Left            =   7875
               TabIndex        =   61
               Top             =   600
               Width           =   810
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
               Height          =   165
               Left            =   3600
               TabIndex        =   60
               Top             =   600
               Width           =   720
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
               Left            =   615
               TabIndex        =   59
               Top             =   240
               Width           =   1155
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
               TabIndex        =   58
               Top             =   240
               Width           =   1095
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
               Height          =   210
               Left            =   1095
               TabIndex        =   57
               Top             =   585
               Width           =   690
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
               Height          =   195
               Left            =   5325
               TabIndex        =   56
               Top             =   600
               Width           =   690
            End
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
            Left            =   180
            TabIndex        =   47
            Top             =   2790
            Width           =   1140
         End
      End
      Begin Threed.SSPanel PanCodigo 
         Height          =   825
         Left            =   120
         TabIndex        =   65
         Top             =   180
         Width           =   9120
         _Version        =   65536
         _ExtentX        =   16087
         _ExtentY        =   1455
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
         Begin MSMask.MaskEdBox MskEmpreendimento 
            Height          =   315
            Left            =   1620
            TabIndex        =   0
            Top             =   60
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboEmpreendimento 
            Bindings        =   "FrmPlanos.frx":270F
            Height          =   315
            Left            =   2340
            TabIndex        =   1
            Top             =   60
            Width           =   5445
            _ExtentX        =   9604
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
         Begin MSDataListLib.DataCombo CboImovel 
            Bindings        =   "FrmPlanos.frx":272F
            Height          =   315
            Left            =   1620
            TabIndex        =   2
            Top             =   420
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
         Begin MSAdodcLib.Adodc DatImovel 
            Height          =   330
            Left            =   1740
            Top             =   420
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
         Begin VB.Label LblCodigo 
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
            Left            =   720
            TabIndex        =   69
            Top             =   480
            Width           =   825
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
            Left            =   120
            TabIndex        =   68
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label LblContrato2 
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
            Left            =   8760
            TabIndex        =   67
            Top             =   480
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
            TabIndex        =   66
            Top             =   480
            Width           =   795
         End
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   6720
      TabIndex        =   29
      Top             =   5010
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
      Left            =   4140
      TabIndex        =   28
      Top             =   5010
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
   Begin Threed.SSCommand CmdNovo 
      Height          =   330
      Left            =   1620
      TabIndex        =   64
      Top             =   5010
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
   End
End
Attribute VB_Name = "FrmFinPlanos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errLoop As ADODB.Error
Dim StrError As String
Dim XFT_PRORATA As String
Dim XFF_VALORNATUREZA As Double 'Valor que falta cadastrar para determinada natureza
Dim XFO_NATUREZA As New ADODB.Recordset
Dim XFF_TOTALCADASTRADO As Double 'Total já cadastrado por uma natureza
Dim XFT_CONTROLE As String 'Usado evitar que não seja chamado a tela de cotações por mais de um controle
Dim XFB_ALTEROU As Boolean 'Indica se houve alteração na data base,no valor histórico

Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    
    If FunObrigatorioMSK(MskEmpreendimento, "O Empreendimento é Obrigatório") Then Exit Function
    If FunObrigatorioCBO(CboImovel, "O Imóvel é Obrigatório") Then Exit Function
    If FunObrigatorioCBO(CboTiposPlanos, "O Tipo de Plano é Obrigatório") Then Exit Function
    If FunObrigatorioCBOX(CboNaturezasPlanos, "A Natureza do Plano é Obrigatório") Then Exit Function
    If FunObrigatorioTXT(TxtParcelas, "O número de parcelas é Obrigatório") Then Exit Function
    If FunObrigatorioTXT(TxtValorReal, "O Valor Histórico da parcelas é Obrigatório") Then Exit Function
    
    If TxtValorReal.Text <= 0 Then
      MsgBox "O valor da parcela deve ser maior que 0!", vbCritical, "ATENÇÃO"
      Exit Function
    End If
    

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
    CboTiposPlanos.Enabled = XHab
    CboNaturezasPlanos.Enabled = XHab
    TxtParcelas.Enabled = XHab
    DtpVencimento.Enabled = XHab
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
    Mskconta.Enabled = XHab
    CboPlanoContas.Enabled = XHab
    CboCCorrente.Enabled = XHab
    TabAuxiliar.Enabled = XHab
    CboIdentificador.Enabled = XHab
    
End Sub

Sub subPrepararInsercao()
    
    DtpDtBase.Value = Date
    CboTiposPlanos.Text = ""
    CboNaturezasPlanos.ListIndex = -1
    LblPlano2.Caption = ""
    LblContrato2.Caption = ""
    TxtParcelas.Text = ""
    LblPeriodicidade2.Caption = ""
    DtpVencimento.Value = Date
    CboObs.Text = ""
    TxtJurosFin.Text = ""
    TxtJuros.Text = ""
    TxtMulta.Text = ""
    TxtValorReal.Text = ""
    LblValorMoeda12.Caption = ""
    LblValorMoeda22.Caption = ""
    LblSimbolo1.Caption = ""
    LblSimbolo2.Caption = ""
    CboMoeda1.Text = ""
    CboMoeda2.Text = ""
    CboPcr.BoundText = ""
    MskPcr.PromptInclude = False
    MskPcr.Text = ""
    MskPcr.PromptInclude = True
    Mskconta.Text = ""
    CboPlanoContas.Text = ""
    CboCCorrente.Text = ""
    LblVlNatureza.Caption = ""
    LblVlCadastrado.Caption = ""
    LblVlCadastrar.Caption = ""
    CboIdentificador.Text = ""
    CboImovel.Enabled = False
    TabAuxiliar.Tab = 0
    
End Sub

Private Sub CboEmpreendimento_Change()
    
    If CboEmpreendimento.Text <> "" Then
        subPrepararInsercao
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        CboImovel.Enabled = True
        subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='A' ORDER BY imov_cd_Imovel", Estatico
    Else
     subPrepararInsercao
   End If
   CboImovel.Text = ""
   SubDesabHabCampos (False)
    
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

Private Sub CmdLimparIdentif_Click()
    CboIdentificador.BoundText = ""
End Sub

Private Sub CboImovel_Change()
    
  If CboImovel.Text <> "" Then
    DatImovel.Recordset.bookmark = CboImovel.SelectedItem
    If Not IsNull(DatImovel.Recordset.Fields("cont_dt_base")) Then DtpDtBase.Value = DatImovel.Recordset.Fields("cont_dt_base")
    Select Case DatImovel.Recordset.Fields("cont_tx_TipoJurosMora")
        Case "D"
            LblAD.Caption = "a.d."
        Case "M"
            LblAD.Caption = "a.m."
        Case Else
            LblAD.Caption = "Cheio"
    End Select
    TxtJuros.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosMora, "##0.0000")
    
    If DatImovel.Recordset.Fields("cont_tx_multaprorata") Then LblProRata.Visible = True
    TxtMulta.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosMulta, "##0.0000")
    
    If DatImovel.Recordset.Fields("cont_tx_desagio") = "S" Then
        LblJurosFin.Caption = "% Juros C.E.F.:"
    Else
        LblJurosFin.Caption = "% Deságio:"
    End If
    
    TxtJurosFin.Text = Format(DatImovel.Recordset.Fields!cont_vl_PercJurosFin, "##0.0000")
            
    LblContrato2.Caption = DatImovel.Recordset.Fields!cont_cd_Contrato
    LblPlano2.Caption = Format(FunGeraPlano, "00")
    DtpDtBase.Value = DatImovel.Recordset.Fields!cont_dt_Base
      
    If Not DatConta.Recordset.EOF Then
      DatConta.Recordset.MoveFirst
      If Not IsNull(DatImovel.Recordset.Fields!plco_cd_Conta) Then
        DatConta.Recordset.Find "plco_cd_Conta =" & DatImovel.Recordset.Fields!plco_cd_Conta
        Mskconta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
        CboPlanoContas.BoundText = Mskconta.Text
      Else
        Mskconta.Text = ""
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
    
    If Not IsNull(DatImovel.Recordset.Fields!moed_cd_CodMoeda) Then
      CboMoeda1.BoundText = DatImovel.Recordset.Fields!moed_cd_CodMoeda
    End If
    
    If Not IsNull(DatImovel.Recordset.Fields!moed_cd_CodMoeda) Then
      CboMoeda2.BoundText = DatImovel.Recordset.Fields!moed_cd_CodMoeda
    End If
    
    XFT_PRORATA = DatImovel.Recordset.Fields("cont_tx_CorrecProRata")

    
    SubQOpenRecordset XFO_NATUREZA, "SELECT cont_vl_Sinal,cont_vl_Financiamento," & _
    "cont_vl_Poupanca,cont_vl_Fgts,cont_vl_Outros,cont_vl_Avista FROM Contratos WHERE " & _
    "empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "'" & _
    " AND imov_cd_Imovel='" & CboImovel.BoundText & "' AND cont_cd_Contrato='" & LblContrato2 & "'", Estatico
    
    
    CboTiposPlanos.BoundText = ""
    LblVlCadastrado = ""
    LblVlCadastrar = ""
    LblVlNatureza = ""
    CboNaturezasPlanos.Clear
    
    If XFO_NATUREZA!cont_vl_Sinal <> 0 Then
      CboNaturezasPlanos.AddItem "SINAL"
    End If
    If XFO_NATUREZA!cont_vl_Poupanca <> 0 Then
      CboNaturezasPlanos.AddItem "POUPANÇA"
    End If
    If XFO_NATUREZA!cont_vl_Financiamento <> 0 Then
      CboNaturezasPlanos.AddItem "FINANCIAMENTO"
    End If
    If XFO_NATUREZA!cont_vl_Fgts <> 0 Then
      CboNaturezasPlanos.AddItem "FGTS"
    End If
    If XFO_NATUREZA!cont_vl_Outros <> 0 Then
      CboNaturezasPlanos.AddItem "OUTROS"
    End If
    If XFO_NATUREZA!cont_vl_Avista <> 0 Then
      CboNaturezasPlanos.AddItem "A VISTA"
    End If
    
    SubDesabHabCampos (True)
      
  End If
End Sub

Function FunGeraPlano() As String
    Dim XLO_TITULO As New ADODB.Recordset
       
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapTitulos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "' and imov_cd_imovel = '" & CboImovel.BoundText & "' AND cont_cd_Contrato= '" & LblContrato2.Caption & "'   ORDER BY Titulo desc", Estatico
    If Not XLO_TITULO.EOF Then
        FunGeraPlano = CStr(CInt(Mid$(XLO_TITULO!Titulo, 14, 2)) + 1)
    Else
        FunGeraPlano = "00"
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

Private Sub CboNaturezasPlanos_Click()
  
  Dim XLT_NATUREZAPLANO As String  'Codigo da Natureza
  Dim XLO_PLANO As New ADODB.Recordset
  
  Select Case CboNaturezasPlanos.Text
    Case "SINAL"
      XLT_NATUREZAPLANO = "S"
    Case "POUPANÇA"
      XLT_NATUREZAPLANO = "P"
    Case "FINANCIAMENTO"
      XLT_NATUREZAPLANO = "F"
    Case "FGTS"
      XLT_NATUREZAPLANO = "G"
    Case "OUTROS"
      XLT_NATUREZAPLANO = "O"
    Case "A VISTA"
      XLT_NATUREZAPLANO = "V"
    Case Else
      XLT_NATUREZAPLANO = ""
  End Select
  
  If XLT_NATUREZAPLANO <> "" Then
    'Pega o valor total já cadastrado para a natureza escolhida
    SubQOpenRecordset XLO_PLANO, "SELECT CONVERT(decimal(18,2),SUM(titu_vl_Parcela)) as Total From Titulos " & _
      "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
      "AND imov_cd_Imovel='" & CboImovel.BoundText & "' AND cont_cd_Contrato='" & LblContrato2.Caption & "' " & _
      "AND titu_tx_NaturezaPlano='" & XLT_NATUREZAPLANO & "' " & _
      "AND titu_cd_Residuo='00'", Estatico
      
    If Not IsNull(XLO_PLANO!Total) Then
       XFF_TOTALCADASTRADO = XLO_PLANO!Total
    Else
       XFF_TOTALCADASTRADO = 0
    End If
      
    XLO_PLANO.Close
    Set XLO_PLANO = Nothing
    
    'Calcula quanto ainda falta cadastrar para a natureza escolhida
    Select Case CboNaturezasPlanos.Text
      Case "SINAL"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Sinal, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Sinal - XFF_TOTALCADASTRADO
      Case "POUPANÇA"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Poupanca, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Poupanca - XFF_TOTALCADASTRADO
      Case "FINANCIAMENTO"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Financiamento, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Financiamento - XFF_TOTALCADASTRADO
      Case "FGTS"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Fgts, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Fgts - XFF_TOTALCADASTRADO
      Case "OUTROS"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Outros, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Outros - XFF_TOTALCADASTRADO
      Case "A VISTA"
        LblVlNatureza = Format(XFO_NATUREZA!cont_vl_Avista, "Standard")
        XFF_VALORNATUREZA = XFO_NATUREZA!cont_vl_Avista - XFF_TOTALCADASTRADO
    End Select
    
    If XFF_VALORNATUREZA < 0 Then
      XFF_VALORNATUREZA = 0
    End If
    
    LblVlCadastrar = Format(XFF_VALORNATUREZA, "standard")
    LblVlCadastrado = Format(XFF_TOTALCADASTRADO, "standard")
    
    
    If TxtParcelas.Text <> "" Then
      TxtValorReal.Text = Format(XFF_VALORNATUREZA / CInt(TxtParcelas.Text), "standard")
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
  MskPcr.PromptInclude = False
  MskPcr.Text = ""
  MskPcr.PromptInclude = True
  
End Sub

Private Sub CmdLimparplanoContas_Click()
  CboPlanoContas.BoundText = ""
  Mskconta.Text = ""
End Sub

Private Sub CmdNovo_Click()
  
  Chave = -1
  XInserir = "I"
  subPrepararInsercao
  'SubDesabHabCampos (True)
  MskEmpreendimento.Enabled = True
  CboEmpreendimento.Enabled = True
  CboImovel.Enabled = True
  Call CboImovel_Change
  CmdGravar.Enabled = True
  CmdNovo.Enabled = False
  
End Sub

Private Sub DtpDtBase_Change()
  
  XFT_CONTROLE = ""
  XFB_ALTEROU = True
  
End Sub

Private Sub Mskconta_LostFocus()
  
  If Mskconta.Text <> "" Then
    
    If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
        Mskconta.Text = Format(Mskconta.Text, XGT_MSKCODREDUZIDO)
    End If
    
    CboPlanoContas.BoundText = Mskconta.Text
    If CboPlanoContas.Text = "" Then
      MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
      Mskconta.SetFocus
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
      CboPlanoContas.BoundText = Mskconta.Text
      CboPlanoContas.SetFocus
      Exit Sub
    End If
    Mskconta.Text = CboPlanoContas.BoundText
  End If
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

Private Sub CboTiposPlanos_Change()
    If CboTiposPlanos.Text <> "" Then
        DatTipoPlano.Recordset.bookmark = CboTiposPlanos.SelectedItem
        If Not IsNull(DatTipoPlano.Recordset.Fields("tipl_nr_periodicidade")) Then
            LblPeriodicidade2.Caption = DatTipoPlano.Recordset.Fields("tipl_nr_periodicidade")
        End If
        If DatTipoPlano.Recordset.Fields("tipl_tx_Unidade") = "M" Then
          LblMes.Caption = "Mês(es)"
        Else
          LblMes.Caption = "Dia(s)"
        End If
    End If
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Function FunAlteraDia(XLD_VENCIMENTO As Date, XLI_DIAS As Long) As Date
  
  'Altera o dia, mas se ultrapassar o limite do mês vai para último dia do mês
  
  Dim XLD_NOVOVENCIMENTO As Date
  
  'Altera o dia do vencimento
  XLD_NOVOVENCIMENTO = Format(XLD_VENCIMENTO, "mm/yyyy")
  XLD_NOVOVENCIMENTO = DateAdd("d", XLI_DIAS - 1, XLD_NOVOVENCIMENTO)
  
  'Verifica se o dia do vencimento existe naquele mês
  If Month(XLD_VENCIMENTO) <> Month(XLD_NOVOVENCIMENTO) Then
    XLD_NOVOVENCIMENTO = Format(XLD_VENCIMENTO, "mm/yyyy")
    If Month(XLD_VENCIMENTO) = 2 And Year(XLD_VENCIMENTO) Mod 4 = 0 Then
      XLD_NOVOVENCIMENTO = DateAdd("d", 28, XLD_NOVOVENCIMENTO)
    ElseIf Month(XLD_VENCIMENTO) = 2 And Year(XLD_VENCIMENTO) Mod 4 <> 0 Then
       XLD_NOVOVENCIMENTO = DateAdd("d", 27, XLD_NOVOVENCIMENTO)
    Else
       XLD_NOVOVENCIMENTO = DateAdd("d", 29, XLD_NOVOVENCIMENTO)
    End If
  End If
  FunAlteraDia = XLD_VENCIMENTO

End Function

Private Sub CmdGravar_Click()
    
    Dim XLT_SQL As String, XLT_CODTITULO As String
    Dim XLD_VENCIMENTO As Date
    Dim XLT_PCR As String, XLT_CONTA As String
    Dim XLI_X As Integer, XLI_PARCELA As Integer
    Dim XLT_NATUREZA As String
    Dim XLF_TOTAL As Double
    Dim XLD_CRIACAO As Date
    Dim XLO_TITULO As Object
    Dim XLT_STATUSQUALIFIN As String    'Indica a operação que será executada no Qualifin (compromissar ou alterar)
    Dim XLF_JUROS As Double
    Dim XLF_DIFERENCA As Double

    On Error GoTo RotuloErro
    
    If funChecarCampos Then Exit Sub
    
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
    
    XLF_DIFERENCA = CDbl(LblVlNatureza.Caption) - (XFF_TOTALCADASTRADO + CInt(TxtParcelas.Text) * CDbl(TxtValorReal.Text))
    Select Case CboNaturezasPlanos.Text
      Case "SINAL"
        XLT_NATUREZA = "S"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor do Sinal em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      Case "POUPANÇA"
        XLT_NATUREZA = "P"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor da Poupança em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      Case "FINANCIAMENTO"
        XLT_NATUREZA = "F"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor do Financiamento em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      Case "FGTS"
        XLT_NATUREZA = "G"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor do FGTS em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      Case "OUTROS"
        XLT_NATUREZA = "O"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor de Outros em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      Case "A VISTA"
        XLT_NATUREZA = "V"
        If XLF_DIFERENCA < 0 Then
          If MsgBox("A criação desse plano excederá o valor A Vista em " & Format(-XLF_DIFERENCA, "standard") & "!" & vbCrLf & "Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
          End If
        End If
      End Select
      
     XLD_CRIACAO = Date
        
    Conexao.BeginTrans
    
    XLI_PARCELA = 0
    
    For XLI_X = 1 To CInt(TxtParcelas.Text)
        XLD_VENCIMENTO = DtpVencimento.Value
        XLI_PARCELA = XLI_PARCELA + 1 'Gera nova Parcela
        If DatImovel.Recordset!cont_tx_Desagio = "S" Then
          XLF_JUROS = XLF_JUROS + CDbl(FunNuloVal(TxtJurosFin.Text))
        Else
          XLF_JUROS = CDbl(FunNuloVal(TxtJurosFin.Text))
        End If
        
        If DatTipoPlano.Recordset.Fields("tipl_tx_Unidade") = "M" Then
          'Incrementa o mês
          XLD_VENCIMENTO = DateAdd("m", CInt(LblPeriodicidade2.Caption) * (XLI_PARCELA - 1), XLD_VENCIMENTO)
        Else
          'Incrementa o dia
          XLD_VENCIMENTO = DateAdd("d", CInt(LblPeriodicidade2.Caption) * (XLI_PARCELA - 1), XLD_VENCIMENTO)
        End If
        
        XLT_SQL = "SELECT * FROM Titulos " & _
                 "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                       "empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
                       "imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
                       "cont_cd_Contrato='" & LblContrato2 & "' AND " & _
                       "titu_cd_Plano='" & LblPlano2.Caption & "'" & " AND " & _
                       "titu_cd_Parcela='-1' AND " & _
                       "titu_cd_Residuo='00'"
        SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
        XLO_TITULO.AddNew
        
        XLO_TITULO!empr_cd_empresa = PCodEmpresa
        XLO_TITULO!empd_cd_Empreendimento = MskEmpreendimento.Text
        XLO_TITULO!imov_cd_imovel = CboImovel.BoundText
        XLO_TITULO!cont_cd_Contrato = LblContrato2.Caption
        XLO_TITULO!titu_cd_Plano = LblPlano2.Caption
        XLO_TITULO!titu_cd_Parcela = Format(CStr(XLI_PARCELA), "000")
        XLO_TITULO!titu_cd_residuo = "00"
        XLO_TITULO!obse_cd_Observacao = funNull(CboObs.BoundText)
        XLO_TITULO!moed_cd_Moeda1 = funNull(CboMoeda1.BoundText)
        XLO_TITULO!moed_cd_Moeda2 = funNull(CboMoeda2.BoundText)
        XLO_TITULO!titu_vl_Parcela = TxtValorReal.Text
        XLO_TITULO!titu_vl_ParcelaIndex1 = FunNuloVal(LblValorMoeda12.Caption)
        XLO_TITULO!titu_vl_ParcelaIndex2 = FunNuloVal(LblValorMoeda22.Caption)
        XLO_TITULO!titu_dt_Vencimento = Format(XLD_VENCIMENTO, "dd/MM/yy")
        XLO_TITULO!tipl_cd_TipoPlano = CboTiposPlanos.BoundText
        XLO_TITULO!titu_vl_PercJurosFin = XLF_JUROS
        XLO_TITULO!titu_vl_PercMultaMora = FunNuloVal(TxtMulta.Text)
        XLO_TITULO!titu_vl_PercJurosMora = FunNuloVal(TxtJuros.Text)
        XLO_TITULO!cere_cd_Pcr = funNull(XLT_PCR)
        XLO_TITULO!plco_cd_Conta = funNull(XLT_CONTA)
        XLO_TITULO!coco_cd_codigo = funNull(CboCCorrente.BoundText)
        XLO_TITULO!titu_dt_Base = Format(DtpDtBase.Value, "dd/MM/yy")
        XLO_TITULO!titu_dt_Prorrogacao = Format(XLD_VENCIMENTO, "dd/MM/yy")
        XLO_TITULO!titu_tx_NaturezaPlano = XLT_NATUREZA
        XLO_TITULO!titu_dt_Criacao = Format(XLD_CRIACAO, "dd/MM/yy")
        XLO_TITULO!iden_cd_Projeto = funNull(CboIdentificador.BoundText)
        XLO_TITULO!titu_nr_ContratoBanco = FunNulo(txtContratoBanco.Text)
                           
        XLO_TITULO.Update
        
        Chave = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & LblContrato2.Caption & "." & LblPlano2.Caption & "." & Format(CStr(XLI_PARCELA), "000") & "." & "00"
        
        If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then 'Interligado com Qualifin e exp. compromissos
            subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", DatImovel.Recordset.Fields("focl_cd_forncli"), CboTiposPlanos.Text, CboNaturezasPlanos.Text
        End If
        XLO_TITULO.Close
        
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        ReDim XGM_MATRIZLOG(18, 2) As Variant

        XGM_MATRIZLOG(0, 0) = "Título"
        XGM_MATRIZLOG(1, 0) = "Data Base"
        XGM_MATRIZLOG(2, 0) = "Data Vencimento"
        XGM_MATRIZLOG(3, 0) = "Data Prorrogação"
        XGM_MATRIZLOG(4, 0) = "Percentual Deságio"
        XGM_MATRIZLOG(5, 0) = "Percentual Juros de Mora"
        XGM_MATRIZLOG(6, 0) = "Percentual Multa"
        XGM_MATRIZLOG(7, 0) = "Valor Histórico"
        XGM_MATRIZLOG(8, 0) = "Moeda 1"
        XGM_MATRIZLOG(9, 0) = "Moeda 2"
        XGM_MATRIZLOG(10, 0) = "Parcela Index1"
        XGM_MATRIZLOG(11, 0) = "Parcela Index2"
        XGM_MATRIZLOG(12, 0) = "Observação"
        XGM_MATRIZLOG(13, 0) = "Centro de Custo"
        XGM_MATRIZLOG(14, 0) = "Conta Contábil"
        XGM_MATRIZLOG(15, 0) = "Conta Corrente"
        XGM_MATRIZLOG(16, 0) = "Identificador"
        XGM_MATRIZLOG(17, 0) = "Nº do Contrato no Banco"
        
        XGM_MATRIZLOG(0, 1) = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & LblContrato2.Caption & "." & LblPlano2.Caption & "." & Format(CStr(XLI_PARCELA), "000") & ".00"
        XGM_MATRIZLOG(1, 1) = Format(DtpDtBase.Value, "dd/MM/yy")
        XGM_MATRIZLOG(2, 1) = Format(DtpVencimento.Value, "dd/MM/yy")
        XGM_MATRIZLOG(3, 1) = Format(DtpVencimento.Value, "dd/MM/yy")
        XGM_MATRIZLOG(4, 1) = FunNuloVal(TxtJurosFin.Text)
        XGM_MATRIZLOG(5, 1) = FunNuloVal(TxtJuros.Text)
        XGM_MATRIZLOG(6, 1) = FunNuloVal(TxtMulta.Text)
        XGM_MATRIZLOG(7, 1) = TxtValorReal.Text
        XGM_MATRIZLOG(8, 1) = funNull(CboMoeda1.BoundText)
        XGM_MATRIZLOG(9, 1) = funNull(CboMoeda2.BoundText)
        XGM_MATRIZLOG(10, 1) = FunNuloVal(LblValorMoeda12.Caption)
        XGM_MATRIZLOG(11, 1) = FunNuloVal(LblValorMoeda22.Caption)
        XGM_MATRIZLOG(12, 1) = funNull(CboObs.BoundText)
        XGM_MATRIZLOG(13, 1) = DatPcr.Recordset.Fields("cere_tx_Descricao") & "-" & DatPcr.Recordset.Fields("cere_cd_estruturado")
        XGM_MATRIZLOG(14, 1) = DatConta.Recordset.Fields("desccombo")
        XGM_MATRIZLOG(15, 1) = funNull(CboCCorrente.BoundText)
        XGM_MATRIZLOG(16, 1) = funNull(CboIdentificador.BoundText)
        XGM_MATRIZLOG(17, 1) = FunNulo(txtContratoBanco.Text)
                    
        Call subRegistraLog("FrmFinPlanos", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CADASTRAR, TITULO_CAPI))
        '*******************************************************************
        
    Next
           
    Conexao.CommitTrans
    MskEmpreendimento.Enabled = False
    CboEmpreendimento.Enabled = False
    CboImovel.Enabled = False
    CmdGravar.Enabled = False
    CmdNovo.Enabled = True
    subPrepararInsercao
    SubDesabHabCampos (False)
    
  Exit Sub
    
RotuloErro:
    funTrataErros (ComMensagem)
    Resume Next
    
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFinPlanos
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
      
    Set Formulario = FrmFinPlanos
    
    ' Verifica se tem o Qualifin acoplado
    If XGB_PQUALIFIN = False Then
      FraConta.Enabled = False
    End If

    subConectarControleDadosNV DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    subConectarControleDadosNV DatMoeda1, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatMoeda2, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
    subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico
      
    DtpDtBase.Value = Date
    DtpVencimento.Value = Date
    CmdNovo.Enabled = False
    SubDesabHabCampos (False)
    XFB_ALTEROU = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  subManutencaoJanelasAtivas "R", "FrmFinPlanos"
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
  Else
    CboEmpreendimento.BoundText = ""
  End If
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
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

Private Sub TxtParcelas_Change()
    
  If CboNaturezasPlanos.ListIndex <> -1 And TxtParcelas.Text <> "" And TxtParcelas.Text <> "0" Then
    TxtValorReal.Text = Format(XFF_VALORNATUREZA / CInt(TxtParcelas.Text), "standard")
  End If
  
End Sub

Private Sub TxtParcelas_KeyPress(KeyAscii As Integer)
  
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
   Else
     XFB_ALTEROU = True
   End If
  
End Sub

Private Sub TxtParcelas_LostFocus()
  
  If TxtParcelas.Text = "0" Then
    TxtParcelas.Text = ""
    XFB_ALTEROU = False
    Exit Sub
  End If
  
  If XFB_ALTEROU = True Then
  
    If CboNaturezasPlanos.ListIndex <> -1 And TxtParcelas.Text <> "" Then
      
      'Preenche valor indexado 1
      If CboMoeda1.BoundText <> "" Then
        XFT_CONTROLE = "TxtParcelas"
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
        XFT_CONTROLE = "TxtParcelas"
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
    XFB_ALTEROU = False
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

Private Sub DtpDtBase_LostFocus()
 
  If XFB_ALTEROU = True Then
    If XFT_CONTROLE <> "TxtValorReal" And XFT_CONTROLE <> "TxtParcelas" Then
    
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
    End If
  XFB_ALTEROU = False
  End If
  XFT_CONTROLE = ""

End Sub

Private Sub TxtValorReal_LostFocus()
    
  If XFB_ALTEROU = True Then
    If XFT_CONTROLE <> "DtpDtBase" And XFT_CONTROLE <> "TxtParcelas" Then
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
    End If
    XFB_ALTEROU = False
  End If
  XFT_CONTROLE = ""
    
End Sub
