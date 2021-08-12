VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFornecedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Cadastro de Fornecedores"
   ClientHeight    =   7965
   ClientLeft      =   2670
   ClientTop       =   2130
   ClientWidth     =   11385
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab TabFornecedores 
      Height          =   7395
      Left            =   0
      TabIndex        =   198
      Top             =   120
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dados Gerais"
      TabPicture(0)   =   "FrmFornecedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PanDadosGerais"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Financeiros/Contábeis"
      TabPicture(1)   =   "FrmFornecedores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PanDadosFinanceiros"
      Tab(1).Control(1)=   "PanCadBancoForn"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Descontos/I&mpostos"
      TabPicture(2)   =   "FrmFornecedores.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraDescontos"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Dados &Pessoais"
      TabPicture(3)   =   "FrmFornecedores.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "PanDadosPessoais"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Qualificações"
      TabPicture(4)   =   "FrmFornecedores.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrmQualificacoes"
      Tab(4).ControlCount=   1
      Begin VB.Frame FrmQualificacoes 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   192
         Top             =   360
         Width           =   10150
         Begin VB.Frame Frame3 
            Caption         =   "Observações"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4335
            Left            =   150
            TabIndex        =   229
            Top             =   1680
            Width           =   9855
            Begin VB.TextBox TxtObservacoes 
               Height          =   3885
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   230
               Top             =   360
               Width           =   9615
            End
         End
         Begin VB.CheckBox chkFornecedorPrincipal 
            Caption         =   "Fornecedor Principal?"
            Height          =   255
            Left            =   150
            TabIndex        =   226
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkFornecedorCritico 
            Caption         =   "Fornecedor Crítico?"
            Height          =   255
            Left            =   2880
            TabIndex        =   225
            Top             =   480
            Width           =   1935
         End
         Begin VB.Frame Frame1 
            Height          =   735
            Left            =   150
            TabIndex        =   193
            Top             =   840
            Width           =   9855
            Begin VB.OptionButton OptFornExclusivo 
               Caption         =   "Fornecedor Exclusivo"
               Height          =   375
               Left            =   6600
               TabIndex        =   197
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton OptNaoQualificado 
               Caption         =   "Não Qualificado"
               Height          =   375
               Left            =   4680
               TabIndex        =   196
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton OptQualificadoRestrincao 
               Caption         =   "Qualificado com Restrinção"
               Height          =   375
               Left            =   1920
               TabIndex        =   195
               Top             =   240
               Width           =   2415
            End
            Begin VB.OptionButton OptQualificado 
               Caption         =   "Qualificado"
               Height          =   375
               Left            =   240
               TabIndex        =   194
               Top             =   240
               Width           =   1455
            End
         End
         Begin MSComCtl2.DTPicker DtpPesquisaUltimaPesquisa 
            Height          =   315
            Left            =   8090
            TabIndex        =   227
            Top             =   450
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   106823681
            CurrentDate     =   37825.4625578704
         End
         Begin VB.Label lblDataPesquisaUltimaQualificacao 
            Caption         =   "Data da Pesquisa da Última Qualificação:"
            Height          =   255
            Left            =   5040
            TabIndex        =   228
            Top             =   480
            Width           =   3015
         End
      End
      Begin Threed.SSPanel PanCadBancoForn 
         Height          =   1545
         Left            =   -72960
         TabIndex        =   102
         Top             =   5040
         Visible         =   0   'False
         Width           =   7770
         _Version        =   65536
         _ExtentX        =   13705
         _ExtentY        =   2725
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
         Begin VB.TextBox TxtDag 
            DataSource      =   "DatBancos"
            Height          =   315
            Left            =   7365
            MaxLength       =   2
            TabIndex        =   108
            Top             =   210
            Width           =   300
         End
         Begin VB.TextBox TxtDac 
            DataField       =   "NumConta"
            DataSource      =   "DatBancos"
            Height          =   315
            Left            =   7365
            MaxLength       =   2
            TabIndex        =   114
            Top             =   675
            Width           =   300
         End
         Begin VB.TextBox TxtNomeAgencia 
            DataField       =   "NomeAgencia"
            DataSource      =   "DatBancos"
            Height          =   315
            Left            =   780
            MaxLength       =   30
            TabIndex        =   110
            Top             =   675
            Width           =   3285
         End
         Begin VB.TextBox TxtCodAgencia 
            DataField       =   "CodAgencia"
            DataSource      =   "DatBancos"
            Height          =   315
            Left            =   5640
            MaxLength       =   10
            TabIndex        =   106
            Top             =   210
            Width           =   1230
         End
         Begin VB.TextBox TxtContaCorrente 
            DataField       =   "NumConta"
            DataSource      =   "DatBancos"
            Height          =   315
            Left            =   4575
            MaxLength       =   20
            TabIndex        =   112
            Top             =   675
            Width           =   2295
         End
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
            Left            =   5310
            TabIndex        =   115
            Top             =   1125
            Width           =   1140
         End
         Begin VB.CommandButton CmdSairBanco 
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
            Left            =   6525
            TabIndex        =   116
            Top             =   1125
            Width           =   1140
         End
         Begin MSDataListLib.DataCombo CboBanco 
            Bindings        =   "FrmFornecedores.frx":008C
            Height          =   315
            Left            =   780
            TabIndex        =   104
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "banc_tx_descricao"
            BoundColumn     =   "banc_cd_codigo"
            Text            =   "CboBanco"
         End
         Begin MSAdodcLib.Adodc DatBanco 
            Height          =   330
            Left            =   2175
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
         Begin VB.Label LblDag 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dag:"
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
            Left            =   6930
            TabIndex        =   107
            Top             =   240
            Width           =   375
         End
         Begin VB.Label LblDac 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dac:"
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
            Left            =   6930
            TabIndex        =   113
            Top             =   705
            Width           =   375
         End
         Begin VB.Label LblBanco 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Banco:"
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
            Left            =   60
            TabIndex        =   103
            Top             =   240
            Width           =   705
         End
         Begin VB.Label LblNomeAgencia 
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
            Left            =   165
            TabIndex        =   109
            Top             =   705
            Width           =   600
         End
         Begin VB.Label LblAgencia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cód. Agência:"
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
            Left            =   4380
            TabIndex        =   105
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label LblContaCorrente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C/C:"
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
            Left            =   4155
            TabIndex        =   111
            Top             =   720
            Width           =   375
         End
      End
      Begin Threed.SSPanel PanDadosPessoais 
         Height          =   6105
         Left            =   -74970
         TabIndex        =   216
         Top             =   270
         Width           =   10305
         _Version        =   65536
         _ExtentX        =   18177
         _ExtentY        =   10769
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame FraDadosPessoais 
            Height          =   2130
            Left            =   30
            TabIndex        =   152
            Top             =   30
            Width           =   10245
            Begin VB.TextBox txtDocEstrangeiro 
               Height          =   315
               Left            =   7560
               TabIndex        =   151
               Top             =   1680
               Width           =   2520
            End
            Begin VB.TextBox txtPassaporte 
               Height          =   315
               Left            =   1200
               TabIndex        =   149
               Top             =   1680
               Width           =   4080
            End
            Begin VB.TextBox txtFiliacaoMae 
               Height          =   315
               Left            =   6075
               TabIndex        =   147
               Top             =   1320
               Width           =   4080
            End
            Begin VB.TextBox txtFiliacaoPai 
               Height          =   315
               Left            =   1200
               TabIndex        =   145
               Top             =   1320
               Width           =   4080
            End
            Begin VB.OptionButton optFeminino 
               Caption         =   "Feminino"
               Height          =   255
               Left            =   2160
               TabIndex        =   139
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton optMasculino 
               Caption         =   "Masculino"
               Height          =   255
               Left            =   960
               TabIndex        =   138
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox txtNaturalidade 
               Height          =   315
               Left            =   8400
               TabIndex        =   136
               Top             =   525
               Width           =   1755
            End
            Begin MSComCtl2.DTPicker MskDtExpedicao 
               Height          =   315
               Left            =   8400
               TabIndex        =   130
               Top             =   165
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   106823681
               CurrentDate     =   39057
            End
            Begin VB.TextBox TxtRG 
               Height          =   315
               Left            =   930
               MaxLength       =   15
               TabIndex        =   124
               Top             =   165
               Width           =   1680
            End
            Begin VB.TextBox TxtRenda 
               Height          =   315
               Left            =   5280
               TabIndex        =   141
               Top             =   960
               Width           =   1020
            End
            Begin VB.ComboBox CboEstCivil 
               Height          =   315
               ItemData        =   "FrmFornecedores.frx":00A3
               Left            =   8400
               List            =   "FrmFornecedores.frx":00B9
               Style           =   2  'Dropdown List
               TabIndex        =   143
               Top             =   960
               Width           =   1785
            End
            Begin VB.TextBox TxtNacionalidade 
               Height          =   315
               Left            =   5280
               MaxLength       =   10
               TabIndex        =   134
               Top             =   525
               Width           =   1755
            End
            Begin VB.TextBox TxtOrgEmissor 
               Height          =   315
               Left            =   3810
               MaxLength       =   8
               TabIndex        =   126
               Top             =   165
               Width           =   930
            End
            Begin MSComCtl2.DTPicker MskDtNascimento 
               Height          =   315
               Left            =   5625
               TabIndex        =   128
               Top             =   165
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   106823681
               CurrentDate     =   37180
            End
            Begin MSDataListLib.DataCombo CboProfissao 
               Bindings        =   "FrmFornecedores.frx":0105
               Height          =   330
               Left            =   930
               TabIndex        =   132
               Top             =   525
               Width           =   2610
               _ExtentX        =   4604
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "prof_tx_descricao"
               BoundColumn     =   "prof_cd_profissao"
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
            Begin VB.Label lblDocEstrangeiro 
               Caption         =   "Doc. Ident. Estrangeira:"
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
               Left            =   5400
               TabIndex        =   150
               Top             =   1680
               Width           =   2055
            End
            Begin VB.Label lblPassaporte 
               Caption         =   "Passaporte:"
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
               TabIndex        =   148
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lblE 
               Caption         =   "e"
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
               Left            =   5760
               TabIndex        =   146
               Top             =   1320
               Width           =   255
            End
            Begin VB.Label lblFiliacao 
               Caption         =   "Filiação:"
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
               TabIndex        =   144
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label lblSexo 
               Caption         =   "Sexo:"
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
               TabIndex        =   137
               Top             =   960
               Width           =   495
            End
            Begin VB.Label lblDtExpedicao 
               Caption         =   "Dt. Expedição:"
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
               Left            =   7080
               TabIndex        =   129
               Top             =   210
               Width           =   1335
            End
            Begin VB.Label LblRenda 
               Alignment       =   1  'Right Justify
               Caption         =   "Renda Familiar:"
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
               Left            =   3840
               TabIndex        =   140
               Top             =   960
               Width           =   1365
            End
            Begin VB.Label LblEstCivil 
               Alignment       =   1  'Right Justify
               Caption         =   "Est. Civil:"
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
               Left            =   7440
               TabIndex        =   142
               Top             =   960
               Width           =   900
            End
            Begin VB.Label LblRG 
               Alignment       =   1  'Right Justify
               Caption         =   "RG:"
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
               Left            =   450
               TabIndex        =   123
               Top             =   210
               Width           =   465
            End
            Begin VB.Label LblOrgEmissor 
               Alignment       =   1  'Right Justify
               Caption         =   "Org. Emissor:"
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
               Left            =   2640
               TabIndex        =   125
               Top             =   210
               Width           =   1155
            End
            Begin VB.Label LblNacionalidade 
               Alignment       =   1  'Right Justify
               Caption         =   "Nacionalidade:"
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
               Left            =   3960
               TabIndex        =   133
               Top             =   570
               Width           =   1305
            End
            Begin VB.Label LblProfissao 
               Alignment       =   1  'Right Justify
               Caption         =   "Profissão:"
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
               TabIndex        =   131
               Top             =   600
               Width           =   855
            End
            Begin VB.Label LblNascimento 
               Alignment       =   1  'Right Justify
               Caption         =   "Dt. Nasc:"
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
               Left            =   4740
               TabIndex        =   127
               Top             =   210
               Width           =   870
            End
            Begin VB.Label lblNaturalidade 
               Caption         =   "Naturalidade:"
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
               Left            =   7200
               TabIndex        =   135
               Top             =   570
               Width           =   1215
            End
         End
         Begin VB.Frame FraEndResidencial 
            Caption         =   "Endereço Residencial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   30
            TabIndex        =   153
            Top             =   2160
            Width           =   10245
            Begin VB.TextBox TxtEndResidencial 
               Height          =   315
               Left            =   930
               MaxLength       =   150
               TabIndex        =   155
               Top             =   225
               Width           =   9240
            End
            Begin VB.TextBox TxtCidadeResidencial 
               Height          =   315
               Left            =   4635
               MaxLength       =   20
               TabIndex        =   159
               Top             =   570
               Width           =   2685
            End
            Begin VB.TextBox TxtBairroResidencial 
               Height          =   315
               Left            =   930
               MaxLength       =   20
               TabIndex        =   157
               Top             =   570
               Width           =   2865
            End
            Begin MSMask.MaskEdBox MskCepResidencial 
               Height          =   315
               Left            =   9090
               TabIndex        =   163
               Top             =   570
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox MskEstadoResidencial 
               Height          =   315
               Left            =   8130
               TabIndex        =   161
               Top             =   570
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin VB.Label LblBairroResidencial 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
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
               Left            =   210
               TabIndex        =   156
               Top             =   600
               Width           =   705
            End
            Begin VB.Label LblEndResidencial 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
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
               Left            =   45
               TabIndex        =   154
               Top             =   255
               Width           =   885
            End
            Begin VB.Label LblCidadeResidencial 
               Alignment       =   1  'Right Justify
               Caption         =   "Cidade:"
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
               Left            =   3900
               TabIndex        =   158
               Top             =   600
               Width           =   705
            End
            Begin VB.Label LblEstadoResidencial 
               Alignment       =   1  'Right Justify
               Caption         =   "Estado:"
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
               Left            =   7440
               TabIndex        =   160
               Top             =   600
               Width           =   660
            End
            Begin VB.Label LblCepResidencial 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
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
               Left            =   8580
               TabIndex        =   162
               Top             =   600
               Width           =   465
            End
         End
         Begin VB.Frame FraDadosConjuge 
            Caption         =   "Dados do Conjuge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   30
            TabIndex        =   177
            Top             =   4440
            Width           =   10245
            Begin VB.TextBox TxtNacConjuge 
               Height          =   315
               Left            =   7545
               MaxLength       =   10
               TabIndex        =   185
               Top             =   570
               Width           =   2625
            End
            Begin VB.TextBox TxtOrgEmissorConjuge 
               Height          =   315
               Left            =   6120
               MaxLength       =   8
               TabIndex        =   189
               Top             =   960
               Width           =   1140
            End
            Begin VB.TextBox TxtRGConjuge 
               Height          =   315
               Left            =   930
               MaxLength       =   15
               TabIndex        =   187
               Top             =   930
               Width           =   2895
            End
            Begin VB.TextBox TxtNomeConjuge 
               Height          =   315
               Left            =   930
               MaxLength       =   50
               TabIndex        =   179
               Top             =   225
               Width           =   5430
            End
            Begin MSComCtl2.DTPicker MskDtNascConjuge 
               Height          =   315
               Left            =   8670
               TabIndex        =   181
               Top             =   225
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   106823681
               CurrentDate     =   37180
            End
            Begin MSDataListLib.DataCombo CboProfConjuge 
               Bindings        =   "FrmFornecedores.frx":0120
               Height          =   330
               Left            =   930
               TabIndex        =   183
               Top             =   570
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "prof_tx_descricao"
               BoundColumn     =   "prof_cd_profissao"
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
            Begin MSMask.MaskEdBox MskCPFConjuge 
               Height          =   315
               Left            =   7890
               TabIndex        =   191
               Top             =   930
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   " "
            End
            Begin VB.Label LblRGConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "RG:"
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
               Left            =   210
               TabIndex        =   186
               Top             =   975
               Width           =   705
            End
            Begin VB.Label LblProfConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "Profissão:"
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
               Left            =   30
               TabIndex        =   182
               Top             =   615
               Width           =   885
            End
            Begin VB.Label LblOrgEmissorConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "Orgão Emissor:"
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
               Left            =   4560
               TabIndex        =   188
               Top             =   975
               Width           =   1350
            End
            Begin VB.Label LblCPFConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "CPF:"
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
               Left            =   7380
               TabIndex        =   190
               Top             =   975
               Width           =   480
            End
            Begin VB.Label LblNomeConjuge 
               Alignment       =   1  'Right Justify
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
               Height          =   210
               Left            =   120
               TabIndex        =   178
               Top             =   270
               Width           =   795
            End
            Begin VB.Label LblNacConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "Nacionalidade:"
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
               Left            =   6210
               TabIndex        =   184
               Top             =   615
               Width           =   1305
            End
            Begin VB.Label LblNascConjuge 
               Alignment       =   1  'Right Justify
               Caption         =   "Nascimento:"
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
               Left            =   7530
               TabIndex        =   180
               Top             =   270
               Width           =   1125
            End
         End
         Begin VB.Frame FraEndCorresp 
            Caption         =   "Endereço para Correspondência"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   30
            TabIndex        =   164
            Top             =   3120
            Width           =   10245
            Begin VB.TextBox TxtBairroCorresp 
               Height          =   315
               Left            =   930
               MaxLength       =   20
               TabIndex        =   168
               Top             =   570
               Width           =   2865
            End
            Begin VB.TextBox TxtCidadeCorresp 
               Height          =   315
               Left            =   4635
               MaxLength       =   20
               TabIndex        =   170
               Top             =   600
               Width           =   2685
            End
            Begin VB.TextBox TxtEndCorresp 
               Height          =   315
               Left            =   930
               MaxLength       =   150
               TabIndex        =   166
               Top             =   225
               Width           =   9240
            End
            Begin MSMask.MaskEdBox MskCepCorresp 
               Height          =   315
               Left            =   9120
               TabIndex        =   174
               Top             =   570
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox MskEstadoCorresp 
               Height          =   315
               Left            =   8130
               TabIndex        =   172
               Top             =   570
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin Threed.SSCommand CmdCopiar 
               Height          =   270
               Left            =   5100
               TabIndex        =   175
               Top             =   930
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   476
               _StockProps     =   78
               Caption         =   "Copiar End. Residencial"
               ForeColor       =   16711680
            End
            Begin Threed.SSCommand CmdCopiarComercial 
               Height          =   270
               Left            =   7140
               TabIndex        =   176
               Top             =   930
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   476
               _StockProps     =   78
               Caption         =   "Copiar End. Comercial"
               ForeColor       =   16711680
            End
            Begin VB.Label LblCepCorresp 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
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
               Left            =   8580
               TabIndex        =   173
               Top             =   615
               Width           =   465
            End
            Begin VB.Label LblEstadoCorresp 
               Alignment       =   1  'Right Justify
               Caption         =   "Estado:"
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
               Left            =   7440
               TabIndex        =   171
               Top             =   615
               Width           =   660
            End
            Begin VB.Label LblCidadeCorresp 
               Alignment       =   1  'Right Justify
               Caption         =   "Cidade:"
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
               Left            =   3840
               TabIndex        =   169
               Top             =   615
               Width           =   705
            End
            Begin VB.Label LblEndCorresp 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
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
               Left            =   45
               TabIndex        =   165
               Top             =   255
               Width           =   885
            End
            Begin VB.Label LblBairroCorresp 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
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
               Left            =   210
               TabIndex        =   167
               Top             =   615
               Width           =   705
            End
         End
      End
      Begin Threed.SSPanel PanDadosGerais 
         Height          =   7020
         Left            =   180
         TabIndex        =   207
         Top             =   300
         Width           =   10230
         _Version        =   65536
         _ExtentX        =   18045
         _ExtentY        =   12382
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
         Begin Threed.SSPanel PanRespCoaf 
            Height          =   1830
            Left            =   0
            TabIndex        =   50
            Top             =   5040
            Visible         =   0   'False
            Width           =   10215
            _Version        =   65536
            _ExtentX        =   18018
            _ExtentY        =   3228
            _StockProps     =   15
            Caption         =   "Responsáveis pela Empresa (Sócios)"
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
            Alignment       =   0
            Begin Threed.SSPanel PanDadosRespCoaf 
               Height          =   1425
               Left            =   1560
               TabIndex        =   54
               Top             =   240
               Visible         =   0   'False
               Width           =   6450
               _Version        =   65536
               _ExtentX        =   11377
               _ExtentY        =   2514
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
               Begin VB.ComboBox CboFuncaoRespCoaf 
                  Height          =   315
                  ItemData        =   "FrmFornecedores.frx":013B
                  Left            =   960
                  List            =   "FrmFornecedores.frx":014E
                  TabIndex        =   58
                  Top             =   600
                  Width           =   2415
               End
               Begin VB.TextBox TxtResponsavelCoaf 
                  DataField       =   "NomeAgencia"
                  DataSource      =   "DatBancos"
                  Height          =   315
                  Left            =   960
                  MaxLength       =   30
                  TabIndex        =   56
                  Top             =   120
                  Width           =   5325
               End
               Begin VB.CommandButton CmdConfirmarRespCoaf 
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
                  Left            =   3840
                  TabIndex        =   61
                  Top             =   960
                  Width           =   1140
               End
               Begin VB.CommandButton CmdRetornarRespCoaf 
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
                  Left            =   5040
                  TabIndex        =   62
                  Top             =   960
                  Width           =   1140
               End
               Begin MSAdodcLib.Adodc Adodc1 
                  Height          =   330
                  Left            =   2175
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
               Begin MSMask.MaskEdBox TxtCPFRespCoaf 
                  Height          =   315
                  Left            =   3840
                  TabIndex        =   60
                  Top             =   570
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   14
                  Mask            =   "###.###.###-##"
                  PromptChar      =   "_"
               End
               Begin VB.Label LblResponsavelCoaf 
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
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   600
               End
               Begin VB.Label LblFuncaoRespCoaf 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Função:"
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
                  TabIndex        =   57
                  Top             =   600
                  Width           =   750
               End
               Begin VB.Label LblCPFRespCoaf 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "CPF:"
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
                  Left            =   3360
                  TabIndex        =   59
                  Top             =   600
                  Width           =   375
               End
            End
            Begin VB.CommandButton cmdInserirRespCoaf 
               Caption         =   "&Inserir"
               Height          =   255
               Left            =   8910
               TabIndex        =   51
               Top             =   255
               Width           =   975
            End
            Begin VB.CommandButton cmdExcluirRespCoaf 
               Caption         =   "&Excluir"
               Height          =   255
               Left            =   8910
               TabIndex        =   53
               Top             =   885
               Width           =   975
            End
            Begin VB.CommandButton cmdAlterarRespCoaf 
               Caption         =   "&Alterar"
               Height          =   255
               Left            =   8910
               TabIndex        =   52
               Top             =   570
               Width           =   975
            End
            Begin TrueDBGrid70.TDBGrid TDBGrid2 
               Height          =   1335
               Left            =   120
               TabIndex        =   222
               Top             =   240
               Width           =   8670
               _ExtentX        =   15293
               _ExtentY        =   2355
               _LayoutType     =   4
               _RowHeight      =   12
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Nome"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Função"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "CPF"
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=6535"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6456"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=74256"
               Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(7)=   "Column(0).FetchStyle=1"
               Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(9)=   "Column(1).Width=4604"
               Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=4524"
               Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74000"
               Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(15)=   "Column(1).FetchStyle=1"
               Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(17)=   "Column(2).Width=2884"
               Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2805"
               Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=74256"
               Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
               Splits(0)._ColumnProps(23)=   "Column(2).FetchStyle=1"
               Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
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
               DeadAreaBackColor=   -2147483633
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(37)  =   ":id=32,.wraptext=-1,.locked=-1"
               _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=44,.alignment=2"
               _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=45"
               _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=47"
               _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=16,.parent=43,.locked=-1"
               _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
               _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45,.bold=0,.fontsize=825"
               _StyleDefs(44)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(45)  =   ":id=14,.fontname=MS Sans Serif"
               _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
               _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=62,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(48)  =   ":id=62,.wraptext=-1,.locked=-1"
               _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=44,.alignment=2"
               _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=45"
               _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=47"
               _StyleDefs(52)  =   "Named:id=33:Normal"
               _StyleDefs(53)  =   ":id=33,.parent=0"
               _StyleDefs(54)  =   "Named:id=34:Heading"
               _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(56)  =   ":id=34,.wraptext=-1"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
         End
         Begin VB.Frame FraTipoForn2 
            Caption         =   "Tipo Fornecedor"
            Enabled         =   0   'False
            Height          =   585
            Left            =   4560
            TabIndex        =   3
            Top             =   120
            Visible         =   0   'False
            Width           =   5580
            Begin VB.OptionButton OptPJServ 
               Caption         =   "P.J. Serviço"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   1440
               TabIndex        =   5
               Top             =   240
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton OptPJProd 
               Caption         =   "P.J. Prod."
               ForeColor       =   &H000040C0&
               Height          =   315
               Left            =   2880
               TabIndex        =   6
               Top             =   240
               Width           =   1035
            End
            Begin VB.OptionButton OptPJProdServ 
               Caption         =   "P.J. Prod./Serv."
               Height          =   300
               Left            =   4080
               TabIndex        =   7
               Top             =   240
               Width           =   1440
            End
            Begin VB.OptionButton OptPF 
               Caption         =   "P. Física"
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   240
               TabIndex        =   4
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Frame FraTipo 
            Caption         =   "Tipo Fornecedor"
            Enabled         =   0   'False
            Height          =   600
            Left            =   4545
            TabIndex        =   217
            Top             =   120
            Visible         =   0   'False
            Width           =   4560
            Begin VB.OptionButton OptProdServ 
               Caption         =   "Produto e Serviço"
               Height          =   300
               Left            =   2730
               TabIndex        =   220
               Top             =   225
               Width           =   1650
            End
            Begin VB.OptionButton OptProduto 
               Caption         =   "Produto"
               Height          =   300
               Left            =   1545
               TabIndex        =   219
               Top             =   225
               Width           =   915
            End
            Begin VB.OptionButton OptServico 
               Caption         =   "Serviço"
               Height          =   300
               Left            =   330
               TabIndex        =   218
               Top             =   225
               Width           =   1155
            End
         End
         Begin VB.Frame FraCGCCPF 
            Height          =   600
            Left            =   45
            TabIndex        =   0
            Top             =   120
            Width           =   4230
            Begin MSMask.MaskEdBox TxtCgc 
               Height          =   315
               Left            =   1530
               TabIndex        =   2
               Top             =   120
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   19
               PromptChar      =   "_"
            End
            Begin VB.Label LblCGC 
               Alignment       =   1  'Right Justify
               Caption         =   "CGC/CPF:"
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
               Left            =   390
               TabIndex        =   1
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame FraEndereco 
            Height          =   4170
            Left            =   60
            TabIndex        =   208
            Top             =   720
            Width           =   10140
            Begin VB.CheckBox chkContICMS 
               Caption         =   "Sim"
               Height          =   255
               Left            =   1755
               TabIndex        =   238
               Top             =   3630
               Width           =   690
            End
            Begin VB.ComboBox cboRegimeApuracao 
               Height          =   315
               Left            =   1640
               Style           =   2  'Dropdown List
               TabIndex        =   236
               Top             =   3120
               Width           =   7870
            End
            Begin VB.ComboBox cboNaturezaJuridica 
               Height          =   315
               ItemData        =   "FrmFornecedores.frx":019D
               Left            =   1640
               List            =   "FrmFornecedores.frx":019F
               Style           =   2  'Dropdown List
               TabIndex        =   234
               Top             =   2760
               Width           =   7870
            End
            Begin VB.TextBox txtNumEndereco 
               Height          =   285
               Left            =   7920
               TabIndex        =   232
               Top             =   885
               Width           =   495
            End
            Begin VB.TextBox txtCodigoMunicipioIbge 
               Height          =   315
               Left            =   9120
               MaxLength       =   7
               TabIndex        =   224
               Top             =   1250
               Width           =   950
            End
            Begin VB.CheckBox chkSuperSimples 
               Caption         =   "Super Simples"
               Enabled         =   0   'False
               Height          =   255
               Left            =   8595
               TabIndex        =   44
               Top             =   3840
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtPais 
               Height          =   315
               Left            =   2760
               TabIndex        =   27
               Top             =   1613
               Width           =   2535
            End
            Begin VB.TextBox TxtIMunicipal 
               Height          =   315
               Left            =   8280
               MaxLength       =   15
               TabIndex        =   13
               Top             =   180
               Width           =   1770
            End
            Begin VB.TextBox TxtBairro 
               Height          =   315
               Left            =   1635
               MaxLength       =   20
               TabIndex        =   21
               Top             =   1250
               Width           =   2505
            End
            Begin VB.TextBox TxtInscricaoEstadual 
               Height          =   315
               Left            =   5520
               MaxLength       =   15
               TabIndex        =   11
               Top             =   180
               Width           =   1620
            End
            Begin VB.TextBox TxtCidade 
               Height          =   315
               Left            =   5040
               MaxLength       =   20
               TabIndex        =   23
               Top             =   1250
               Width           =   2055
            End
            Begin VB.TextBox TxtEndereco 
               Height          =   315
               Left            =   1635
               MaxLength       =   100
               TabIndex        =   17
               Top             =   889
               Width           =   5460
            End
            Begin VB.TextBox TxtEMail 
               Height          =   315
               Left            =   7680
               MaxLength       =   50
               TabIndex        =   31
               Top             =   1613
               Width           =   2385
            End
            Begin VB.TextBox TxtRazaoSocial 
               Height          =   315
               Left            =   1635
               LinkTimeout     =   65
               MaxLength       =   65
               TabIndex        =   15
               Top             =   527
               Width           =   8430
            End
            Begin VB.TextBox TxtNomeFantasia 
               Height          =   315
               Left            =   1635
               MaxLength       =   20
               TabIndex        =   9
               Top             =   165
               Width           =   2715
            End
            Begin VB.TextBox TxtNumPorta 
               Height          =   315
               Left            =   9300
               MaxLength       =   5
               TabIndex        =   19
               Top             =   885
               Width           =   765
            End
            Begin MSMask.MaskEdBox MskCep 
               Height          =   315
               Left            =   5760
               TabIndex        =   29
               Top             =   1620
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox MskEstado 
               Height          =   315
               Left            =   1635
               TabIndex        =   25
               Top             =   1613
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskCelular 
               Height          =   315
               Left            =   8520
               TabIndex        =   39
               Top             =   1980
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   15
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFax 
               Height          =   315
               Left            =   6000
               TabIndex        =   37
               Top             =   1980
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   15
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFoneResid 
               Height          =   315
               Left            =   1635
               TabIndex        =   33
               Top             =   1980
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   15
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox MskFoneCom 
               Height          =   315
               Left            =   4080
               TabIndex        =   35
               Top             =   1980
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   15
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskAtividadePrincipal 
               Height          =   330
               Left            =   1635
               TabIndex        =   41
               Top             =   2350
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   9
               Mask            =   "9999-9/99"
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboAtividadePrincipal 
               Bindings        =   "FrmFornecedores.frx":01A1
               Height          =   330
               Left            =   2760
               TabIndex        =   42
               Top             =   2355
               Visible         =   0   'False
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "cnse_tx_descricao"
               BoundColumn     =   "cnse_cd_cnaesecao"
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
            Begin Threed.SSCommand CmdAtividadePrincipal 
               Height          =   315
               Left            =   9600
               TabIndex        =   43
               Top             =   2355
               Visible         =   0   'False
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":01C5
               Picture         =   "FrmFornecedores.frx":0617
            End
            Begin VB.Label lblContICMS 
               Caption         =   "Contribuinte ICMS:"
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
               Left            =   45
               TabIndex        =   237
               Top             =   3630
               Width           =   1620
            End
            Begin VB.Label lblRegimeApuracao 
               Alignment       =   1  'Right Justify
               Caption         =   "Regime Apuração:"
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
               Left            =   45
               TabIndex        =   235
               Top             =   3165
               Width           =   1575
            End
            Begin VB.Label lblNaturezaJuridica 
               Caption         =   "Natureza Jurídica:"
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
               Left            =   40
               TabIndex        =   233
               Top             =   2780
               Width           =   1575
            End
            Begin VB.Label lblNumEndereco 
               Caption         =   "Número:"
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
               Left            =   7200
               TabIndex        =   231
               Top             =   930
               Width           =   735
            End
            Begin VB.Label lblCodigoMunicipioIbge 
               Caption         =   "Cód. Município IBGE:"
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
               Left            =   7200
               TabIndex        =   223
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label LblPais 
               Caption         =   "País:"
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
               Left            =   2280
               TabIndex        =   26
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label lblAtividadePrincipal 
               Alignment       =   1  'Right Justify
               Caption         =   "Ativ. Principal:"
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
               Left            =   150
               TabIndex        =   40
               Top             =   2400
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label LblIMunicipal 
               Alignment       =   1  'Right Justify
               Caption         =   "I. Municipal:"
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
               Left            =   7200
               TabIndex        =   12
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label LblInscricaoEstadual 
               Alignment       =   1  'Right Justify
               Caption         =   "I. Estadual:"
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
               Left            =   4395
               TabIndex        =   10
               Top             =   225
               Width           =   1080
            End
            Begin VB.Label LblEMail 
               Alignment       =   1  'Right Justify
               Caption         =   "E-Mail:"
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
               Left            =   6840
               TabIndex        =   30
               Top             =   1680
               Width           =   690
            End
            Begin VB.Label LblCep 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
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
               Left            =   5280
               TabIndex        =   28
               Top             =   1680
               Width           =   465
            End
            Begin VB.Label LblCelular 
               Alignment       =   1  'Right Justify
               Caption         =   "Celular:"
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
               Left            =   7560
               TabIndex        =   38
               Top             =   2040
               Width           =   825
            End
            Begin VB.Label LblEstado 
               Alignment       =   1  'Right Justify
               Caption         =   "Estado:"
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
               Left            =   945
               TabIndex        =   24
               Top             =   1665
               Width           =   660
            End
            Begin VB.Label LblCidade 
               Alignment       =   1  'Right Justify
               Caption         =   "Cidade:"
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
               Left            =   4200
               TabIndex        =   22
               Top             =   1290
               Width           =   735
            End
            Begin VB.Label LblFax 
               Alignment       =   1  'Right Justify
               Caption         =   "Fax:"
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
               Left            =   5520
               TabIndex        =   36
               Top             =   2040
               Width           =   465
            End
            Begin VB.Label LblNomeFantasia 
               Alignment       =   1  'Right Justify
               Caption         =   "Nome Fantasia:"
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
               TabIndex        =   8
               Top             =   210
               Width           =   1365
            End
            Begin VB.Label LblRazaoSocial 
               Alignment       =   1  'Right Justify
               Caption         =   "Razão Social:"
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
               Left            =   210
               TabIndex        =   14
               Top             =   570
               Width           =   1395
            End
            Begin VB.Label LblEndereco 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
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
               Left            =   30
               TabIndex        =   16
               Top             =   915
               Width           =   1575
            End
            Begin VB.Label LblBairro 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
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
               Left            =   900
               TabIndex        =   20
               Top             =   1290
               Width           =   705
            End
            Begin VB.Label LblFoneCom 
               Alignment       =   1  'Right Justify
               Caption         =   "Comercial:"
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
               Left            =   3120
               TabIndex        =   34
               Top             =   2040
               Width           =   915
            End
            Begin VB.Label LblFoneResid 
               Alignment       =   1  'Right Justify
               Caption         =   "Fone Residencial:"
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
               Left            =   60
               TabIndex        =   32
               Top             =   2040
               Width           =   1545
            End
            Begin VB.Label LblNumPorta 
               Alignment       =   1  'Right Justify
               Caption         =   "Nº Porta:"
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
               Left            =   8505
               TabIndex        =   18
               Top             =   930
               Width           =   795
            End
         End
         Begin Threed.SSFrame FraContas2 
            Height          =   1200
            Left            =   60
            TabIndex        =   209
            Top             =   5760
            Visible         =   0   'False
            Width           =   10125
            _Version        =   65536
            _ExtentX        =   17859
            _ExtentY        =   2117
            _StockProps     =   14
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
            Begin MSDataListLib.DataCombo CboDesprec2 
               Bindings        =   "FrmFornecedores.frx":0729
               Height          =   330
               Left            =   2595
               TabIndex        =   199
               Top             =   135
               Width           =   6930
               _ExtentX        =   12224
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
            Begin MSMask.MaskEdBox MskDespRec2 
               Height          =   315
               Left            =   1635
               TabIndex        =   63
               Top             =   135
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskCusto2 
               Height          =   315
               Left            =   1635
               TabIndex        =   200
               Top             =   480
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboCusto2 
               Bindings        =   "FrmFornecedores.frx":0742
               Height          =   330
               Left            =   2595
               TabIndex        =   201
               Top             =   480
               Width           =   6930
               _ExtentX        =   12224
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
            Begin Threed.SSCommand CmdLimparDespRec2 
               Height          =   315
               Left            =   9600
               TabIndex        =   210
               Top             =   120
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":0759
               Picture         =   "FrmFornecedores.frx":0BAB
            End
            Begin Threed.SSCommand CmdLimparPcr2 
               Height          =   300
               Left            =   9675
               TabIndex        =   211
               Top             =   855
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   529
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":0CBD
               Picture         =   "FrmFornecedores.frx":110F
            End
            Begin Threed.SSCommand CmdLimparCusto2 
               Height          =   300
               Left            =   9675
               TabIndex        =   212
               Top             =   510
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   529
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":1221
               Picture         =   "FrmFornecedores.frx":1673
            End
            Begin MSMask.MaskEdBox MskPcr2 
               Height          =   315
               Left            =   1635
               TabIndex        =   202
               Top             =   825
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr2 
               Bindings        =   "FrmFornecedores.frx":1785
               Height          =   330
               Left            =   2595
               TabIndex        =   203
               Top             =   825
               Width           =   6930
               _ExtentX        =   12224
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
            Begin VB.Label LblDespRec2 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Receita:"
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
               TabIndex        =   215
               Top             =   195
               Width           =   1425
            End
            Begin VB.Label LblCusto2 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Ativo:"
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
               Left            =   45
               TabIndex        =   214
               Top             =   525
               Width           =   1560
            End
            Begin VB.Label LblPCR2 
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
               Height          =   210
               Left            =   75
               TabIndex        =   213
               Top             =   870
               Width           =   1545
            End
         End
         Begin Threed.SSFrame FraContato 
            Height          =   765
            Left            =   60
            TabIndex        =   45
            Top             =   4920
            Width           =   10140
            _Version        =   65536
            _ExtentX        =   17886
            _ExtentY        =   1349
            _StockProps     =   14
            Caption         =   "Contatos"
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
            Begin VB.TextBox TxtContatoFinanceiro 
               Height          =   315
               Left            =   5895
               MaxLength       =   20
               TabIndex        =   49
               Top             =   315
               Width           =   3024
            End
            Begin VB.TextBox TxtContatoComercial 
               Height          =   315
               Left            =   1530
               MaxLength       =   20
               TabIndex        =   47
               Top             =   315
               Width           =   3024
            End
            Begin VB.Label LblContatoFinanceiro 
               Alignment       =   1  'Right Justify
               Caption         =   "Financeiro:"
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
               Left            =   4785
               TabIndex        =   48
               Top             =   345
               Width           =   1035
            End
            Begin VB.Label LblContatoComercial 
               Alignment       =   1  'Right Justify
               Caption         =   "Comercial:"
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
               Left            =   555
               TabIndex        =   46
               Top             =   345
               Width           =   930
            End
         End
      End
      Begin Threed.SSPanel PanDadosFinanceiros 
         Height          =   4755
         Left            =   -74940
         TabIndex        =   206
         Top             =   300
         Width           =   10230
         _Version        =   65536
         _ExtentX        =   18045
         _ExtentY        =   8387
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
         Begin Threed.SSPanel PanIcms 
            Height          =   510
            Left            =   7560
            TabIndex        =   87
            Top             =   1800
            Visible         =   0   'False
            Width           =   2460
            _Version        =   65536
            _ExtentX        =   4339
            _ExtentY        =   900
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
            Begin VB.TextBox TxtIcms 
               Height          =   315
               Left            =   1230
               TabIndex        =   89
               Top             =   90
               Width           =   660
            End
            Begin VB.Label LblIcms 
               Alignment       =   1  'Right Justify
               Caption         =   "ICMS (%):"
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
               Height          =   210
               Left            =   285
               TabIndex        =   88
               Top             =   135
               Width           =   900
            End
         End
         Begin Threed.SSFrame FraContas 
            Height          =   1710
            Left            =   45
            TabIndex        =   64
            Top             =   0
            Width           =   10005
            _Version        =   65536
            _ExtentX        =   17648
            _ExtentY        =   3016
            _StockProps     =   14
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
            Begin MSMask.MaskEdBox MskDespRec 
               Height          =   330
               Left            =   2000
               TabIndex        =   66
               Top             =   180
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskProvAdiant 
               Height          =   330
               Left            =   2000
               TabIndex        =   74
               Top             =   900
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskCusto 
               Height          =   330
               Left            =   2000
               TabIndex        =   70
               Top             =   540
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboProvAdiant 
               Bindings        =   "FrmFornecedores.frx":179A
               Height          =   330
               Left            =   2925
               TabIndex        =   75
               Top             =   900
               Width           =   6375
               _ExtentX        =   11245
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
            Begin MSDataListLib.DataCombo CboCusto 
               Bindings        =   "FrmFornecedores.frx":17B6
               Height          =   330
               Left            =   2925
               TabIndex        =   71
               Top             =   540
               Width           =   6375
               _ExtentX        =   11245
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
            Begin MSDataListLib.DataCombo CboDespRec 
               Bindings        =   "FrmFornecedores.frx":17CD
               Height          =   330
               Left            =   2925
               TabIndex        =   67
               Top             =   180
               Width           =   6375
               _ExtentX        =   11245
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
            Begin Threed.SSCommand CmdLimparDespRec 
               Height          =   315
               Left            =   9435
               TabIndex        =   68
               Top             =   180
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":17E6
               Picture         =   "FrmFornecedores.frx":1C38
            End
            Begin Threed.SSCommand CmdLimparPcr 
               Height          =   315
               Left            =   9435
               TabIndex        =   80
               Top             =   1260
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":1D4A
               Picture         =   "FrmFornecedores.frx":219C
            End
            Begin Threed.SSCommand CmdLimparProvAdiant 
               Height          =   315
               Left            =   9435
               TabIndex        =   76
               Top             =   900
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":22AE
               Picture         =   "FrmFornecedores.frx":2700
            End
            Begin Threed.SSCommand CmdLimparCusto 
               Height          =   315
               Left            =   9435
               TabIndex        =   72
               Top             =   540
               Width           =   390
               _Version        =   65536
               _ExtentX        =   688
               _ExtentY        =   556
               _StockProps     =   78
               ForeColor       =   -2147483630
               MouseIcon       =   "FrmFornecedores.frx":2812
               Picture         =   "FrmFornecedores.frx":2C64
            End
            Begin MSMask.MaskEdBox MskPcr 
               Height          =   330
               Left            =   2000
               TabIndex        =   78
               Top             =   1260
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmFornecedores.frx":2D76
               Height          =   330
               Left            =   2925
               TabIndex        =   79
               Top             =   1260
               Width           =   6375
               _ExtentX        =   11245
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
            Begin VB.Label LblProvAdiant 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Adiantamento:"
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
               Left            =   60
               TabIndex        =   73
               Top             =   915
               Width           =   1890
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
               Height          =   210
               Left            =   375
               TabIndex        =   77
               Top             =   1290
               Width           =   1575
            End
            Begin VB.Label LblCusto 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Passivo:"
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
               Left            =   375
               TabIndex        =   69
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label LblDespRec 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Despesa:"
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
               Left            =   405
               TabIndex        =   65
               Top             =   195
               Width           =   1545
            End
         End
         Begin Threed.SSPanel PanCodAtividade 
            Height          =   510
            Left            =   4800
            TabIndex        =   84
            Top             =   1800
            Width           =   2580
            _Version        =   65536
            _ExtentX        =   4551
            _ExtentY        =   900
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
            Begin VB.TextBox TxtCodAtividade 
               Height          =   315
               Left            =   1455
               MaxLength       =   5
               TabIndex        =   86
               Top             =   90
               Width           =   660
            End
            Begin VB.Label LblCodAtividade 
               Alignment       =   1  'Right Justify
               Caption         =   "Cód. Atividade:"
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
               Height          =   210
               Left            =   90
               TabIndex        =   85
               Top             =   135
               Width           =   1335
            End
         End
         Begin Threed.SSPanel PanTipoPagto 
            Height          =   510
            Left            =   60
            TabIndex        =   81
            Top             =   1800
            Width           =   4620
            _Version        =   65536
            _ExtentX        =   8149
            _ExtentY        =   900
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
            Begin VB.ComboBox CboTipoPagamento 
               Height          =   315
               ItemData        =   "FrmFornecedores.frx":2D8B
               Left            =   2010
               List            =   "FrmFornecedores.frx":2DA1
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   90
               Width           =   2235
            End
            Begin VB.Label LblTipoPagamento 
               Caption         =   "Forma de Pagamento:"
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
               TabIndex        =   82
               Top             =   135
               Width           =   1905
            End
         End
         Begin Threed.SSPanel PanBancoForn 
            Height          =   1830
            Left            =   60
            TabIndex        =   90
            Top             =   2400
            Width           =   9990
            _Version        =   65536
            _ExtentX        =   17621
            _ExtentY        =   3228
            _StockProps     =   15
            Caption         =   "Banco do Fornecedor"
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
            Alignment       =   0
            Begin VB.CommandButton CmdAlteraBanco 
               Caption         =   "&Alterar"
               Height          =   255
               Left            =   8910
               TabIndex        =   93
               Top             =   570
               Width           =   975
            End
            Begin VB.CommandButton CmdRemoverBanco 
               Caption         =   "&Excluir"
               Height          =   255
               Left            =   8910
               TabIndex        =   94
               Top             =   885
               Width           =   975
            End
            Begin VB.CommandButton CmdInsereBanco 
               Caption         =   "&Inserir"
               Height          =   255
               Left            =   8910
               TabIndex        =   92
               Top             =   255
               Width           =   975
            End
            Begin VB.CommandButton CmdPadrao 
               Caption         =   "&Bc. Padrão"
               Height          =   255
               Left            =   8910
               TabIndex        =   95
               Top             =   1200
               Width           =   975
            End
            Begin TrueDBGrid70.TDBGrid TDBGrid3 
               Height          =   1335
               Left            =   165
               TabIndex        =   91
               Top             =   240
               Width           =   8670
               _ExtentX        =   15293
               _ExtentY        =   2355
               _LayoutType     =   4
               _RowHeight      =   12
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Banco"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Nome"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Agência"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Dag"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Nome"
               Columns(4).DataField=   ""
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "Conta Corrente"
               Columns(5).DataField=   ""
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "Dac"
               Columns(6).DataField=   ""
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Padrao"
               Columns(7).DataField=   ""
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "CodBanco"
               Columns(8).DataField=   ""
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(9)._VlistStyle=   0
               Columns(9)._MaxComboItems=   5
               Columns(9).Caption=   "Status"
               Columns(9).DataField=   ""
               Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   10
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).Locked=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0).AllowColMove=   -1  'True
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=10"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=74256"
               Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(7)=   "Column(0).FetchStyle=1"
               Splits(0)._ColumnProps(8)=   "Column(0).AllowFocus=0"
               Splits(0)._ColumnProps(9)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(10)=   "Column(1).Width=3334"
               Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3254"
               Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=74256"
               Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(16)=   "Column(1).FetchStyle=1"
               Splits(0)._ColumnProps(17)=   "Column(1).AllowFocus=0"
               Splits(0)._ColumnProps(18)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(19)=   "Column(2).Width=1667"
               Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=1588"
               Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=74256"
               Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
               Splits(0)._ColumnProps(25)=   "Column(2).FetchStyle=1"
               Splits(0)._ColumnProps(26)=   "Column(2).AllowFocus=0"
               Splits(0)._ColumnProps(27)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(28)=   "Column(3).Width=714"
               Splits(0)._ColumnProps(29)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(30)=   "Column(3)._WidthInPix=635"
               Splits(0)._ColumnProps(31)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(32)=   "Column(3)._ColStyle=66065"
               Splits(0)._ColumnProps(33)=   "Column(3).WrapText=1"
               Splits(0)._ColumnProps(34)=   "Column(3).FetchStyle=1"
               Splits(0)._ColumnProps(35)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(36)=   "Column(4).Width=2884"
               Splits(0)._ColumnProps(37)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(4)._WidthInPix=2805"
               Splits(0)._ColumnProps(39)=   "Column(4)._EditAlways=0"
               Splits(0)._ColumnProps(40)=   "Column(4)._ColStyle=74256"
               Splits(0)._ColumnProps(41)=   "Column(4).WrapText=1"
               Splits(0)._ColumnProps(42)=   "Column(4).FetchStyle=1"
               Splits(0)._ColumnProps(43)=   "Column(4).AllowFocus=0"
               Splits(0)._ColumnProps(44)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(45)=   "Column(5).Width=2593"
               Splits(0)._ColumnProps(46)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(47)=   "Column(5)._WidthInPix=2514"
               Splits(0)._ColumnProps(48)=   "Column(5)._EditAlways=0"
               Splits(0)._ColumnProps(49)=   "Column(5)._ColStyle=74256"
               Splits(0)._ColumnProps(50)=   "Column(5).WrapText=1"
               Splits(0)._ColumnProps(51)=   "Column(5).FetchStyle=1"
               Splits(0)._ColumnProps(52)=   "Column(5).AllowFocus=0"
               Splits(0)._ColumnProps(53)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(54)=   "Column(6).Width=661"
               Splits(0)._ColumnProps(55)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(56)=   "Column(6)._WidthInPix=582"
               Splits(0)._ColumnProps(57)=   "Column(6)._EditAlways=0"
               Splits(0)._ColumnProps(58)=   "Column(6)._ColStyle=66065"
               Splits(0)._ColumnProps(59)=   "Column(6).WrapText=1"
               Splits(0)._ColumnProps(60)=   "Column(6).FetchStyle=1"
               Splits(0)._ColumnProps(61)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(62)=   "Column(7).Width=661"
               Splits(0)._ColumnProps(63)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(64)=   "Column(7)._WidthInPix=582"
               Splits(0)._ColumnProps(65)=   "Column(7)._EditAlways=0"
               Splits(0)._ColumnProps(66)=   "Column(7).AllowSizing=0"
               Splits(0)._ColumnProps(67)=   "Column(7)._ColStyle=74256"
               Splits(0)._ColumnProps(68)=   "Column(7).Visible=0"
               Splits(0)._ColumnProps(69)=   "Column(7).WrapText=1"
               Splits(0)._ColumnProps(70)=   "Column(7).FetchStyle=1"
               Splits(0)._ColumnProps(71)=   "Column(7).AllowFocus=0"
               Splits(0)._ColumnProps(72)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(73)=   "Column(8).Width=2725"
               Splits(0)._ColumnProps(74)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(75)=   "Column(8)._WidthInPix=2646"
               Splits(0)._ColumnProps(76)=   "Column(8)._EditAlways=0"
               Splits(0)._ColumnProps(77)=   "Column(8).AllowSizing=0"
               Splits(0)._ColumnProps(78)=   "Column(8)._ColStyle=65808"
               Splits(0)._ColumnProps(79)=   "Column(8).Visible=0"
               Splits(0)._ColumnProps(80)=   "Column(8).WrapText=1"
               Splits(0)._ColumnProps(81)=   "Column(8).AllowFocus=0"
               Splits(0)._ColumnProps(82)=   "Column(8).Order=9"
               Splits(0)._ColumnProps(83)=   "Column(8)._MinWidth=90520480"
               Splits(0)._ColumnProps(84)=   "Column(9).Width=2725"
               Splits(0)._ColumnProps(85)=   "Column(9).DividerColor=0"
               Splits(0)._ColumnProps(86)=   "Column(9)._WidthInPix=2646"
               Splits(0)._ColumnProps(87)=   "Column(9)._EditAlways=0"
               Splits(0)._ColumnProps(88)=   "Column(9).AllowSizing=0"
               Splits(0)._ColumnProps(89)=   "Column(9)._ColStyle=65808"
               Splits(0)._ColumnProps(90)=   "Column(9).Visible=0"
               Splits(0)._ColumnProps(91)=   "Column(9).WrapText=1"
               Splits(0)._ColumnProps(92)=   "Column(9).AllowFocus=0"
               Splits(0)._ColumnProps(93)=   "Column(9).Order=10"
               Splits(0)._ColumnProps(94)=   "Column(9)._MinWidth=90519328"
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
               DeadAreaBackColor=   -2147483633
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
               _StyleDefs(37)  =   ":id=28,.wraptext=-1,.locked=-1"
               _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
               _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
               _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
               _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(42)  =   ":id=32,.wraptext=-1,.locked=-1"
               _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
               _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
               _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
               _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(47)  =   ":id=58,.wraptext=-1,.locked=-1"
               _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
               _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
               _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=24,.parent=43,.alignment=2"
               _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=21,.parent=44,.alignment=2"
               _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=22,.parent=45"
               _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=23,.parent=47"
               _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=62,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(56)  =   ":id=62,.wraptext=-1,.locked=-1"
               _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=44,.alignment=2"
               _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=45"
               _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=47"
               _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(61)  =   ":id=66,.wraptext=-1,.locked=-1"
               _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44,.alignment=2"
               _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
               _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
               _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=2"
               _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
               _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
               _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
               _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=70,.parent=43,.alignment=0,.valignment=2"
               _StyleDefs(70)  =   ":id=70,.wraptext=-1,.locked=-1"
               _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=44,.alignment=2"
               _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=45"
               _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=47"
               _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=16,.parent=43"
               _StyleDefs(75)  =   "Splits(0).Columns(8).HeadingStyle:id=13,.parent=44"
               _StyleDefs(76)  =   "Splits(0).Columns(8).FooterStyle:id=14,.parent=45"
               _StyleDefs(77)  =   "Splits(0).Columns(8).EditorStyle:id=15,.parent=47"
               _StyleDefs(78)  =   "Splits(0).Columns(9).Style:id=20,.parent=43"
               _StyleDefs(79)  =   "Splits(0).Columns(9).HeadingStyle:id=17,.parent=44"
               _StyleDefs(80)  =   "Splits(0).Columns(9).FooterStyle:id=18,.parent=45"
               _StyleDefs(81)  =   "Splits(0).Columns(9).EditorStyle:id=19,.parent=47"
               _StyleDefs(82)  =   "Named:id=33:Normal"
               _StyleDefs(83)  =   ":id=33,.parent=0"
               _StyleDefs(84)  =   "Named:id=34:Heading"
               _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(86)  =   ":id=34,.wraptext=-1"
               _StyleDefs(87)  =   "Named:id=35:Footing"
               _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(89)  =   "Named:id=36:Selected"
               _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(91)  =   "Named:id=37:Caption"
               _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(93)  =   "Named:id=38:HighlightRow"
               _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
               _StyleDefs(95)  =   "Named:id=39:EvenRow"
               _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(97)  =   "Named:id=40:OddRow"
               _StyleDefs(98)  =   ":id=40,.parent=33"
               _StyleDefs(99)  =   "Named:id=41:RecordSelector"
               _StyleDefs(100) =   ":id=41,.parent=34"
               _StyleDefs(101) =   "Named:id=42:FilterBar"
               _StyleDefs(102) =   ":id=42,.parent=33"
            End
            Begin VB.Label LblExclusao 
               BackStyle       =   0  'Transparent
               Caption         =   "Banco Excluído"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   165
               TabIndex        =   96
               Top             =   1590
               Width           =   1485
            End
            Begin VB.Label LblPadrao 
               BackStyle       =   0  'Transparent
               Caption         =   "Banco Padrão"
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   1740
               TabIndex        =   97
               Top             =   1590
               Width           =   1425
            End
         End
         Begin MSMask.MaskEdBox TxtNit 
            Height          =   315
            Left            =   2880
            TabIndex        =   101
            Top             =   4320
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   19
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtPis 
            Height          =   315
            Left            =   600
            TabIndex        =   99
            Top             =   4320
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   19
            PromptChar      =   "_"
         End
         Begin VB.Label LblNit 
            Caption         =   "NIT:"
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
            Left            =   2400
            TabIndex        =   100
            Top             =   4365
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LblPis 
            Caption         =   "PIS:"
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
            TabIndex        =   98
            Top             =   4350
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame FraDescontos 
         Height          =   4545
         Left            =   -74865
         TabIndex        =   117
         Top             =   435
         Width           =   10035
         Begin TrueDBGrid70.TDBGrid TDBGrid1 
            Height          =   3780
            Left            =   105
            TabIndex        =   122
            Top             =   675
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   6668
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Descontos/Impostos"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Valor (%)"
            Columns(1).DataField=   ""
            Columns(1).NumberFormat=   "Standard"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Codigo"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=10134"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=10054"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=74256"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=2170"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2090"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=66322"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=74000"
            Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
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
            HeadLines       =   1,2
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
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
            _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1"
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
            _StyleDefs(37)  =   ":id=28,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=1,.valignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=1"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
            _StyleDefs(46)  =   ":id=58,.locked=-1"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
            _StyleDefs(50)  =   "Named:id=33:Normal"
            _StyleDefs(51)  =   ":id=33,.parent=0"
            _StyleDefs(52)  =   "Named:id=34:Heading"
            _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(54)  =   ":id=34,.wraptext=-1"
            _StyleDefs(55)  =   "Named:id=35:Footing"
            _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=36:Selected"
            _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=37:Caption"
            _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(61)  =   "Named:id=38:HighlightRow"
            _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(63)  =   "Named:id=39:EvenRow"
            _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(65)  =   "Named:id=40:OddRow"
            _StyleDefs(66)  =   ":id=40,.parent=33"
            _StyleDefs(67)  =   "Named:id=41:RecordSelector"
            _StyleDefs(68)  =   ":id=41,.parent=34"
            _StyleDefs(69)  =   "Named:id=42:FilterBar"
            _StyleDefs(70)  =   ":id=42,.parent=33"
         End
         Begin VB.CommandButton CmdExcluir 
            Caption         =   "E&xcluir"
            Height          =   330
            Left            =   8940
            TabIndex        =   121
            Top             =   690
            Width           =   960
         End
         Begin VB.CommandButton CmdInsere 
            Caption         =   "I&nserir"
            Height          =   330
            Left            =   8940
            TabIndex        =   120
            Top             =   270
            Width           =   960
         End
         Begin MSAdodcLib.Adodc DatDescontosImpostos 
            Height          =   330
            Left            =   3510
            Top             =   270
            Visible         =   0   'False
            Width           =   2925
            _ExtentX        =   5159
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
            Caption         =   "DatDescontosImpostos"
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
         Begin MSDataListLib.DataCombo CboDescontosImpostos 
            Bindings        =   "FrmFornecedores.frx":2DE9
            Height          =   315
            Left            =   3240
            TabIndex        =   119
            Top             =   270
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "desc_tx_descricao"
            BoundColumn     =   "desc_cd_desconto"
            Text            =   ""
         End
         Begin VB.Label LblSelMat 
            Alignment       =   1  'Right Justify
            Caption         =   "Selecione o Imposto a ser incluído:"
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
            Left            =   165
            TabIndex        =   118
            Top             =   315
            Width           =   3030
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   221
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   3360
      TabIndex        =   204
      Top             =   7560
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2413
      _ExtentY        =   593
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
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   6360
      TabIndex        =   205
      Top             =   7560
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
   Begin MSAdodcLib.Adodc DatCusto 
      Height          =   330
      Left            =   30
      Top             =   3495
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "DatCM"
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
   Begin MSAdodcLib.Adodc DatProvAdiant 
      Height          =   330
      Left            =   30
      Top             =   3090
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "DatAdiantamento"
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
   Begin MSAdodcLib.Adodc DatDespRec 
      Height          =   330
      Left            =   15
      Top             =   3900
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "DatProvisionamento"
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
      Left            =   30
      Negotiate       =   -1  'True
      Top             =   4290
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
   Begin MSAdodcLib.Adodc DatProfissao 
      Height          =   330
      Left            =   15
      Negotiate       =   -1  'True
      Top             =   4635
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "DatProfissao"
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
   Begin MSAdodcLib.Adodc DatAtividadePrincipal 
      Height          =   330
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "DatAtividadePrincipal"
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
Attribute VB_Name = "FrmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ResCGCCPF As ADODB.Recordset  'Necessário para realizar a consistência do CGCCPF, tipo e classe
Dim ChecaCGCCPF As Boolean 'Esta variável irá checar se o CGC e CPF já foi preenchido
Dim CodigoDocumento As Long 'Armazena a chave

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor1 As New XArray
Dim Vetor3 As New XArray
Dim Vetor2 As New XArray

Dim EntrouImp As Boolean 'Controla a entrada na orelha de impostos

Dim XBanco As String 'I-Inserir A-Alterar
Dim XCoaf As String 'I-Inserir A-Alterar Responsáveis COAF
Dim XBancoPadrao As String

Dim ResFormulario As Object

Function ChecarCampos() As Boolean
    ChecarCampos = True
    
    If Pessoa = "F" Then 'Física
        If FunObrigatorioTXT(TxtNomeFantasia, "O Nome é Obrigatório") Then Exit Function
        If FunObrigatorioTXT(TxtRazaoSocial, "O Nome Completo é Obrigatório") Then Exit Function
    ElseIf Pessoa = "J" Or Pessoa = "3" Or Pessoa = "4" Then 'Jurídica ou Produtos
        If FunObrigatorioTXT(TxtNomeFantasia, "O Nome Fantasia é Obrigatório") Then Exit Function
        If FunObrigatorioTXT(TxtRazaoSocial, "A Razão Social é Obrigatória") Then Exit Function
    End If
    
    If TxtNumPorta <> "" Then
        If Not IsNumeric(TxtNumPorta.Text) Then
            MsgBox "O número da porta deve ter formato numérico!", vbExclamation, "AVISO"
            Exit Function
        End If
    End If
    
    If CboAtividadePrincipal.Text <> "" Then 'Só pode ser gravado item de subgategoria 19/05/09 - Patrícia
            DatAtividadePrincipal.Recordset.bookmark = CboAtividadePrincipal.SelectedItem
            If Len(DatAtividadePrincipal.Recordset.Fields("cnse_tx_codigo")) <> 10 Then
                MsgBox "A Atividade Principal cadastrada não é uma Sub-Classe, logo não pode ser atribuída ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
                CboAtividadePrincipal.BoundText = MskAtividadePrincipal.Text
                Exit Function
            End If
            MskAtividadePrincipal.Text = Right(Replace(DatAtividadePrincipal.Recordset.Fields("cnse_tx_codigo"), "-", ""), 8)
    End If
    ChecarCampos = False
End Function

Sub DesabHabCampos(XHab As Boolean)
'Desabilita/Habilita os campos em caso de inserção
    TabFornecedores.TabEnabled(1) = XHab
    TabFornecedores.TabEnabled(2) = XHab
    TabFornecedores.TabEnabled(4) = XHab
    
    TxtNomeFantasia.Enabled = XHab
    TxtRazaoSocial.Enabled = XHab
    TxtEndereco.Enabled = XHab
    TxtNumPorta.Enabled = XHab
    TxtBairro.Enabled = XHab
    TxtCidade.Enabled = XHab
    txtCodigoMunicipioIbge.Enabled = XHab
    txtNumEndereco.Enabled = XHab
    
    MskFax.Enabled = XHab
    MskFoneResid.Enabled = XHab
    MskFoneCom.Enabled = XHab
    MskCelular.Enabled = XHab
    MskEstado.Enabled = XHab
    MskCep.Enabled = XHab
        
    TxtEmail.Enabled = XHab
    TxtInscricaoEstadual.Enabled = XHab
    TxtIMunicipal.Enabled = XHab
    TxtContatoComercial.Enabled = XHab
    TxtContatoFinanceiro.Enabled = XHab
    
    LblNomeFantasia.Enabled = XHab
    LblRazaoSocial.Enabled = XHab
    LblEndereco.Enabled = XHab
    LblNumPorta.Enabled = XHab
    LblBairro.Enabled = XHab
    LblCidade.Enabled = XHab
    LblEstado.Enabled = XHab
    LblCep.Enabled = XHab
    lblCodigoMunicipioIbge.Enabled = XHab
    LblEMail.Enabled = XHab
    LblFoneResid.Enabled = XHab
    LblFoneCom.Enabled = XHab
    LblFax.Enabled = XHab
    LblCelular.Enabled = XHab
    LblFax.Enabled = XHab
    LblInscricaoEstadual.Enabled = XHab
    LblIMunicipal.Enabled = XHab
    LblContatoComercial.Enabled = XHab
    LblContatoFinanceiro.Enabled = XHab
    lblNumEndereco.Enabled = XHab
    lblNaturezaJuridica.Enabled = XHab
    lblRegimeApuracao.Enabled = XHab
    lblContICMS.Enabled = XHab
    
    chkContICMS.Enabled = XHab
    cboRegimeApuracao.Enabled = XHab
    cboNaturezaJuridica.Enabled = XHab
    
    FraEndereco.Enabled = XHab
    FraContato.Enabled = XHab
    
    CmdGravar.Enabled = XHab
    
    If FornClie = "C" Then
        LblPCR2.Enabled = XHab
        LblCusto2.Enabled = XHab
        LblDespRec2.Enabled = XHab
        CboDesprec2.Enabled = XHab
        CboCusto2.Enabled = XHab
        CboPcr2.Enabled = XHab
        MskDespRec2.Enabled = XHab
        MskCusto2.Enabled = XHab
        CmdLimparDespRec2.Enabled = XHab
        CmdLimparCusto2.Enabled = XHab
        CmdLimparPcr2.Enabled = XHab
    End If
    
End Sub

Sub DesabilitaBotoes(orelha As String)
    If orelha = "I" Then 'Desabilita os botões da orelha de impostos
        CmdInsere.Enabled = False
        'CmdAltera.Enabled = False
        CmdExcluir.Enabled = False
    ElseIf orelha = "B" Then 'Desabilita os botões do painel de Cad. Banco fornecedor
        CmdInsereBanco.Enabled = False
        CmdAlteraBanco.Enabled = False
        CmdRemoverBanco.Enabled = False
    End If
End Sub


Sub HabilitaBotoes(orelha As String)
    If orelha = "I" Then   'Habilita os botoes da orelha de impostos
        CmdInsere.Enabled = True
        'CmdAltera.Enabled = True
        CmdExcluir.Enabled = True
    ElseIf orelha = "B" Then 'Habilita botões do painel de Cad. Banco Fornecedor
        CmdInsereBanco.Enabled = True
        CmdAlteraBanco.Enabled = True
        CmdRemoverBanco.Enabled = True
    End If
End Sub



Sub HabilitaCampos(classe As String, Tipo As String)
'Habilita os campos a depender do tipo e da classe
    lblNumEndereco.Visible = True
    txtNumEndereco.Visible = True
    
    lblNaturezaJuridica.Visible = True
    cboNaturezaJuridica.Visible = True
    
    lblRegimeApuracao.Visible = True
    cboRegimeApuracao.Visible = True
    
    lblContICMS.Visible = True
    chkContICMS.Visible = True
    
    If classe = "F" Then 'Caso seja fornecedor
        TabFornecedores.TabVisible(3) = True
        TabFornecedores.TabsPerRow = 5
        FraCGCCPF.Top = 150
        FraEndereco.Top = 800
        If Tipo = "F" Then 'Pessoa Física
            FrmFornecedores.Caption = "Cadastro de Fornecedores (Serviços - Pessoa Física)"
            FraEndereco.Top = 1300
            'FraEndereco.Left = 90
            LblInscricaoEstadual.Visible = False
            TxtInscricaoEstadual.Visible = False
            LblIMunicipal.Visible = True
            TxtIMunicipal.Visible = True
            FraContato.Visible = False
            FraCGCCPF.Top = 500
            LblPis.Visible = True
            LblNit.Visible = True
            TxtNit.Visible = True
            TxtPis.Visible = True
            TxtPis.Mask = Format$("###.###.###-##")
            TxtNit.Mask = Format$("###.###.###-##")
            'FraCGCCPF.Left = 90
            If XGT_NomeAplic = "QUALIPAT" Then ' Se o sistema for o QualiPat
              TabFornecedores.TabVisible(1) = False
              TabFornecedores.TabVisible(2) = False
            End If
        Else ' Pessoa Jurídica
            FrmFornecedores.Caption = "Cadastro de Fornecedores"
            If XInserir = "A" Then
                LblContatoComercial.Enabled = True
                TxtContatoComercial.Enabled = True
                LblContatoFinanceiro.Enabled = True
                TxtContatoFinanceiro.Enabled = True
                LblInscricaoEstadual.Enabled = True
                TxtInscricaoEstadual.Enabled = True
                LblIMunicipal.Enabled = True
                TxtIMunicipal.Enabled = True
                FraContato.Enabled = True
            End If
            FraTipo.Visible = True
            FraTipo.Enabled = True
            OptServico.TabIndex = 1
            OptProduto.TabIndex = 2
            OptProdServ.TabIndex = 3
            If XGT_NomeAplic = "QUALIPAT" Then ' Se o sistema for o QualiPat
              'FraTipo.Visible = False
              'FraTipoForn2.Visible = False
              TabFornecedores.TabVisible(1) = False
              TabFornecedores.TabVisible(2) = False
            End If
        End If
        'Se a tela for chamada por Notas Fiscais
        If XFormForn = "FrmNotasFiscais" Then
            FraTipo.Visible = False
            FraTipoForn2.Visible = True
            FraTipoForn2.Enabled = True
            FraTipoForn2.Top = FraCGCCPF.Top
        End If
    ElseIf classe = "C" Then
        TabFornecedores.TabVisible(1) = False
        TabFornecedores.TabVisible(2) = False
        TabFornecedores.TabVisible(4) = False
        
        If XGT_NomeAplic = "CAPI" Then
            TabFornecedores.TabVisible(3) = True
            TabFornecedores.TabsPerRow = 2
            FraContas2.Visible = False
            FraEndereco.Top = 1300
            FraCGCCPF.Top = 500
            FraTipoForn2.Top = FraCGCCPF.Top
            If Tipo = "F" Then LblEndereco.Caption = "End. Comercial:"
            'Responsáveis Legais do COAF
            If Pessoa = "J" And FornClie = "C" Then
                PanRespCoaf.Visible = True
                Sql = "SELECT focl_cd_FornCli, recf_tx_Responsavel, recf_tx_Funcao, recf_tx_Cpf " & _
                    " FROM ResponsavelCoaf" & _
                    " WHERE focl_cd_forncli =" & CodigoDocumento
                subCarregaVetor Sql, Array("recf_tx_Responsavel", "recf_tx_Funcao", "recf_tx_Cpf", ""), Vetor2, TDBGrid2
            If Vetor2(0, 0) = Empty Then
                DesabilitaBotoes ("I")
                cmdInserirRespCoaf.Enabled = True
            End If
            End If
        Else
            TabFornecedores.TabVisible(3) = False
            TabFornecedores.TabsPerRow = 2
            FraContas2.Visible = True
        End If
    
        If Tipo = "F" Then 'Pessoa Física
            FrmFornecedores.Caption = "Cadastro de Clientes (Pessoa Física)"
            LblInscricaoEstadual.Enabled = False
            TxtInscricaoEstadual.Enabled = False
            LblIMunicipal.Enabled = True
            TxtIMunicipal.Enabled = True
            'Adicionado no dia 14/08/2018 - Samuel Jesus. Solicitação feita por Daniela para atender ao E-Social
            LblPis.Enabled = True
            TxtPis.Enabled = True
        Else 'Pessoa Jurídica
            FrmFornecedores.Caption = "Cadastro de Clientes (Pessoa Jurídica)"
            LblInscricaoEstadual.Enabled = True
            TxtInscricaoEstadual.Enabled = True
            LblIMunicipal.Enabled = True
            TxtIMunicipal.Enabled = True
        End If
        TabFornecedores.TabCaption(0) = "Dados Cadastrais"
        FraContato.Visible = False
        FraEndereco.Visible = True
        PanCodAtividade.Visible = False
        
        If XFormForn = "FrmNotasFiscais" Or XFormForn = "FrmManContratos" Then
            FraTipoForn2.Caption = "Tipo Cliente"
            FraTipoForn2.Visible = True
            FraTipoForn2.Enabled = True
            OptPJProd.Visible = False
            OptPJProdServ.Visible = False
            OptPJServ.Caption = "Pessoa Jurídica"
            OptPJServ.Left = 2595
            OptPJServ.Width = 2500
            OptPF.Caption = "Pessoa Física"
            OptPF.Left = 690
            OptPF.Width = 2000
            FraTipoForn2.Top = FraCGCCPF.Top
        End If
    End If
End Sub

Sub PreencheCampos()
    Dim XLT_QUALIFICACAO    As String
    Dim XLT_SQL As String
    Dim XLO_FORNCLI As Object
    
    SubQOpenRecordset XLO_FORNCLI, "SELECT * FROM ConsGENFornCli WHERE empr_cd_empresa = " & PCodEmpresa & " AND focl_cd_forncli =" & Chave, "Estatico"
    '*************************************************************************************
    'XLT_SQL = "SELECT *, " & _
            "(SELECT plco_cd_ContaRecDesp FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaRecDesp, " & _
            "(SELECT plco_cd_ContaCusto FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaCusto, " & _
            "(SELECT plco_cd_ContaProvAdiant FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaProvAdiant, " & _
            "(SELECT cere_cd_Pcr FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS cere_cd_Pcr, " & _
            "(SELECT PlanoContas.plco_cd_CodReduzido FROM PlanoContas RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedRecDesp, " & _
            "(SELECT  PlanoContas.plco_tx_DescContabil FROM PlanoContas RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contarecdesp, " & _
            "(SELECT  PlanoContas_1.plco_cd_CodReduzido FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedCusto, " & _
            "(SELECT  PlanoContas_1.plco_tx_DescContabil FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contacusto, " & _
            "(SELECT  PlanoContas_2.plco_cd_CodReduzido FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedProvAdiant, " & _
            "(SELECT  PlanoContas_2.plco_tx_DescContabil FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contaprovadiant, " & _
            "(SELECT  dbo.CentroResultados.cere_tx_Descricao FROM CentroResultados RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS cere_tx_Descricao, " & _
            "(SELECT  dbo.CentroResultados.cere_cd_Estruturado FROM dbo.CentroResultados RIGHT OUTER JOIN " & _
            "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ")  AS cere_cd_Estruturado" & _
            " FROM fornclientes WHERE focl_cd_forncli =" & Chave
    '*************************************************************************************
    'SubQOpenRecordset XLO_FORNCLI, XLT_SQL, "Estatico"
    
    If Not (XLO_FORNCLI.BOF And XLO_FORNCLI.EOF) Then
        If Not IsNull(XLO_FORNCLI!focl_tx_CGCCPF) And (XLO_FORNCLI!focl_tx_CGCCPF <> "") Then
            TxtCgc.Mask = ""
            If Pessoa = "J" Then
                TxtCgc.Text = Format$(XLO_FORNCLI!focl_tx_CGCCPF, "##.###.###/####-##")
                TxtInscricaoEstadual.Text = FunNulo(XLO_FORNCLI!focl_tx_iestadual)
                TxtIMunicipal.Text = FunNulo(XLO_FORNCLI!focl_tx_imunicipal)
            Else
                TxtCgc.Text = Format$(XLO_FORNCLI!focl_tx_CGCCPF, "##.###.###-##")
            End If
        End If
        TxtNomeFantasia.Text = FunNulo(XLO_FORNCLI!focl_tx_fantasia)
        TxtRazaoSocial.Text = FunNulo(XLO_FORNCLI!focl_tx_RazaoSocial)
        TxtEndereco.Text = FunNulo(XLO_FORNCLI!focl_tx_endereco)
        txtCodigoMunicipioIbge.Text = FunNulo(XLO_FORNCLI!focl_cd_MunicipioIbge)
        TxtNumPorta.Text = FunNulo(XLO_FORNCLI!focl_nr_porta)
        TxtBairro.Text = FunNulo(XLO_FORNCLI!focl_tx_Bairro)
        TxtCidade.Text = FunNulo(XLO_FORNCLI!focl_tx_cidade)
        TxtEmail.Text = FunNulo(XLO_FORNCLI!focl_tx_email)
        TxtIMunicipal.Text = FunNulo(XLO_FORNCLI!focl_tx_imunicipal)
        txtNumEndereco.Text = FunNulo(XLO_FORNCLI!focl_tx_EnderecoNumero)
        cboNaturezaJuridica.ListIndex = FunNuloVal(XLO_FORNCLI!focl_nr_NaturezaJuridica)
        cboRegimeApuracao.ListIndex = FunNuloVal(XLO_FORNCLI!focl_nr_RegimeApuracao)
        If XLO_FORNCLI!focl_bl_ContribuinteICMS = 0 Then
            chkContICMS.Value = vbUnchecked
        Else
            chkContICMS.Value = vbChecked
        End If
            
        If XGT_NomeAplic = "CAPI" Then
            If Not IsNull(XLO_FORNCLI!focl_tx_pais) Then
                txtPais.Text = FunNulo(XLO_FORNCLI!focl_tx_pais)
            End If
        End If
        TxtObservacoes.Text = FunNulo(XLO_FORNCLI!focl_tx_Observacoes)
        If (XLO_FORNCLI!focl_nr_pis) <> "" Then
            TxtPis.Text = Format$(XLO_FORNCLI!focl_nr_pis, "###.###.###-##")
        End If
        
        If (XLO_FORNCLI!focl_nr_nit) <> "" Then
            TxtNit.Text = FunNulo(Format$(XLO_FORNCLI!focl_nr_nit, "###.###.###-##"))
        End If
       'Qualificação
       If XLO_FORNCLI("Qualificacoes").Value = "Qualificado" Then
          OptQualificado.Value = True
       ElseIf XLO_FORNCLI("Qualificacoes").Value = "Não Qualificado" Then
         OptNaoQualificado.Value = True
       ElseIf XLO_FORNCLI("Qualificacoes").Value = "Qualificado com restrinção" Then
         OptQualificadoRestrincao.Value = True
         ElseIf XLO_FORNCLI("Qualificacoes").Value = "Fornecedor Exclusivo" Then
       OptFornExclusivo.Value = True
       End If
   
        TxtContatoComercial.Text = FunNulo(XLO_FORNCLI!focl_tx_contatoCom)
        TxtContatoFinanceiro.Text = FunNulo(XLO_FORNCLI!focl_tx_ContatoFin)
        If (Not IsNull(XLO_FORNCLI!focl_tx_cep)) And XLO_FORNCLI!focl_tx_cep <> "" Then
            MskCep.PromptInclude = False
            MskCep.Text = XLO_FORNCLI!focl_tx_cep
            MskCep.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_tx_estado)) And XLO_FORNCLI!focl_tx_estado <> "" Then
            MskEstado.PromptInclude = False
            MskEstado.Text = XLO_FORNCLI!focl_tx_estado
            MskEstado.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_tx_fone)) And XLO_FORNCLI!focl_tx_fone <> "" Then
            MskFoneResid.PromptInclude = False
            MskFoneResid.Text = XLO_FORNCLI!focl_tx_fone
            MskFoneResid.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_tx_fone1)) And XLO_FORNCLI!focl_tx_fone1 <> "" Then
            MskFoneCom.PromptInclude = False
            MskFoneCom.Text = XLO_FORNCLI!focl_tx_fone1
            MskFoneCom.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_tx_Celular)) And XLO_FORNCLI!focl_tx_Celular <> "" Then
            MskCelular.PromptInclude = False
            MskCelular.Text = XLO_FORNCLI!focl_tx_Celular
            MskCelular.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_tx_fax)) And XLO_FORNCLI!focl_tx_fax <> "" Then
            MskFax.PromptInclude = False
            MskFax.Text = XLO_FORNCLI!focl_tx_fax
            MskFax.PromptInclude = True
        End If
        
        If (Not IsNull(XLO_FORNCLI!focl_vl_Icms)) And XLO_FORNCLI!focl_vl_Icms <> "" Then
            TxtIcms.Text = Format(XLO_FORNCLI!focl_vl_Icms, "standard")
        End If
        
        If FornClie = "F" Then 'Fornecedor
            If (XLO_FORNCLI!plco_cd_ContaRecDesp <> "") And Not IsNull(XLO_FORNCLI!plco_cd_ContaRecDesp) Then
                    DatDespRec.Recordset.MoveFirst
                    DatDespRec.Recordset.Find "plco_cd_Conta =" & XLO_FORNCLI!plco_cd_ContaRecDesp
                    If Not DatDespRec.Recordset.EOF Then
                        MskDespRec.Text = DatDespRec.Recordset.Fields("plco_cd_codreduzido")
                    Else
                        MsgBox "A Conta Despesa definida nesse cadastro, não é de natureza 'Despesa' e não será mostrada. Defina uma nova Conta.", vbCritical + vbOKOnly, "ATENÇÃO"
                    End If
                    CboDespRec.BoundText = MskDespRec.Text
            End If
            
            If (XLO_FORNCLI!plco_cd_contaprovadiant <> "") And Not IsNull(XLO_FORNCLI!plco_cd_contaprovadiant) Then
                DatProvAdiant.Recordset.MoveFirst
                DatProvAdiant.Recordset.Find "plco_cd_Conta =" & XLO_FORNCLI!plco_cd_contaprovadiant
                MskProvAdiant.Text = DatProvAdiant.Recordset.Fields("plco_cd_Codreduzido")
                CboProvAdiant.BoundText = MskProvAdiant.Text
            End If
            
            If (XLO_FORNCLI!plco_cd_ContaCusto <> "") And Not IsNull(XLO_FORNCLI!plco_cd_ContaCusto) Then
                DatCusto.Recordset.MoveFirst
                DatCusto.Recordset.Find "plco_cd_Conta =" & XLO_FORNCLI!plco_cd_ContaCusto
                If Not DatCusto.Recordset.EOF Then
                    MskCusto.Text = DatCusto.Recordset.Fields("plco_cd_Codreduzido")
                Else
                    MsgBox "A Conta Passivo definida nesse cadastro, não é de natureza 'Passivo' e não será mostrada. Defina uma nova Conta.", vbCritical + vbOKOnly, "ATENÇÃO"
                End If
                CboCusto.BoundText = MskCusto.Text
            End If
            
            If (XLO_FORNCLI!cere_cd_Pcr <> "") And Not IsNull(XLO_FORNCLI!cere_cd_Pcr) Then
                DatPcr.Recordset.MoveFirst
                DatPcr.Recordset.Find "cere_cd_pcr =" & XLO_FORNCLI!cere_cd_Pcr
                
                'se o recordset vier vazio é pq o centro de resultados está inativo.
                If Not DatPcr.Recordset.EOF Then
                    'so atribui ao cbo se o status for ativo
                    MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
                    CboPcr.BoundText = MskPcr.Text
                End If
            End If
            
            CboTipoPagamento.ListIndex = FunNuloVal(XLO_FORNCLI!focl_tx_TipoPag)
            If Not IsNull(XLO_FORNCLI!focl_nr_codatividade) Then TxtCodAtividade = XLO_FORNCLI!focl_nr_codatividade
            If Pessoa <> "F" Then
                Select Case XLO_FORNCLI!focl_tx_Tipo
                    Case "J"
                        OptServico.Value = True
                    Case "3"
                        OptProduto.Value = True
                    Case "4"
                        OptProdServ.Value = True
                End Select
                
                'Henrique Lima - 16/04/09 - Carrega o valor do campo Super Simples
                If FunNuloVal(XLO_FORNCLI!focl_tx_SuperSimples) = "S" Then
                    chkSuperSimples.Value = Checked
                Else
                    chkSuperSimples.Value = Unchecked
                End If
                
                'Cássio Medeiros - 16/04/13
                If FunNuloVal(XLO_FORNCLI!focl_tx_fornecedorPrincipal) = "S" Then
                    chkFornecedorPrincipal.Value = Checked
                Else
                    chkFornecedorPrincipal.Value = Unchecked
                End If
                
                If FunNulo(XLO_FORNCLI!focl_dt_pesquisaUltimaQualificacao) <> "" Then
                    DtpPesquisaUltimaPesquisa.Value = XLO_FORNCLI!focl_dt_pesquisaUltimaQualificacao
                End If
        
                'Everaldo Sena - 30/05/12
                If FunNuloVal(XLO_FORNCLI!focl_tx_fornecedorcritico) = "S" Then
                    chkFornecedorCritico.Value = Checked
                Else
                    chkFornecedorCritico.Value = Unchecked
                End If
            End If
        ElseIf FornClie = "C" Then 'Cliente
            If (XLO_FORNCLI!plco_cd_ContaRecDesp <> "") And Not IsNull(XLO_FORNCLI!plco_cd_ContaRecDesp) Then
                DatDespRec.Recordset.MoveFirst
                DatDespRec.Recordset.Find "plco_cd_Conta =" & XLO_FORNCLI!plco_cd_ContaRecDesp
                If Not DatDespRec.Recordset.EOF Then
                    MskDespRec2.Text = DatDespRec.Recordset.Fields("plco_cd_Codreduzido")
                Else
                    MsgBox "A Conta Receita definida nesse cadastro, não é de natureza 'Receita' e não será mostrada. Defina uma nova Conta.", vbCritical + vbOKOnly, "ATENÇÃO"
                End If
                CboDesprec2.BoundText = MskDespRec2.Text
            End If
        
            If (XLO_FORNCLI!plco_cd_ContaCusto <> "") And Not IsNull(XLO_FORNCLI!plco_cd_ContaCusto) Then
                DatCusto.Recordset.MoveFirst
                DatCusto.Recordset.Find "plco_cd_Conta =" & XLO_FORNCLI!plco_cd_ContaCusto
                If Not DatCusto.Recordset.EOF Then
                    MskCusto2.Text = DatCusto.Recordset.Fields("plco_cd_Codreduzido")
                Else
                    MsgBox "A Conta Ativo definida nesse cadastro, não é de natureza 'Ativo' e não será mostrada. Defina uma nova Conta.", vbCritical + vbOKOnly, "ATENÇÃO"
                End If
                CboCusto2.BoundText = MskCusto2.Text
            End If
            
            If (XLO_FORNCLI!cere_cd_Pcr <> "") And Not IsNull(XLO_FORNCLI!cere_cd_Pcr) Then
                DatPcr.Recordset.MoveFirst
                DatPcr.Recordset.Find "cere_cd_pcr =" & XLO_FORNCLI!cere_cd_Pcr
                
                If Not (DatPcr.Recordset.BOF Or DatPcr.Recordset.EOF) Then
                    MskPcr2.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
                Else
                    If FornClie = "C" Then
                        MsgBox "O Centro de Custo definido para esse cliente foi desativado. Defina um novo Centro de Custo para o cliente.", vbCritical + vbOKOnly, "AVISO"
                    Else
                        MsgBox "O Centro de Custo definido para esse fornecedor foi desativado. Defina um novo Centro de Custo para o fornecedor.", vbCritical + vbOKOnly, "AVISO"
                    End If
                End If
                CboPcr2.BoundText = MskPcr2.Text
            End If
            
            If XGT_NomeAplic = "CAPI" Or XGB_PCAPI Then 'MÓDULO CAPI
            
                Dim XLO_CLIENTECAPI As ADODB.Recordset
                
                If XLO_FORNCLI!focl_tx_Tipo = "J" Then
                                        
                    If Not IsNull(XLO_FORNCLI!cnsc_cd_cnaeSubclasse) Then
                        CboAtividadePrincipal.BoundText = XLO_FORNCLI!cnsc_cd_cnaeSubclasse
                   
                    
                        SubQOpenRecordset XLO_CLIENTECAPI, "SELECT cnsc_tx_descricao, cnsc_tx_codigo FROM CnaeSubClasse WHERE cnsc_cd_cnaeSubClasse = " & XLO_FORNCLI!cnsc_cd_cnaeSubclasse & "", Dinamico
                        If Not XLO_CLIENTECAPI.EOF Then
                            MskAtividadePrincipal.Mask = ""
                            MskAtividadePrincipal.Text = Replace(XLO_CLIENTECAPI!cnsc_tx_codigo, "-", "")
                            'XLO_CLIENTECAPI!cnsc_tx_codigo
                            MskAtividadePrincipal.Mask = "9999-9/99"
                        End If
                        XLO_CLIENTECAPI.Close
                    End If
                End If

                SubQOpenRecordset XLO_CLIENTECAPI, "SELECT * FROM Clientes WHERE focl_cd_forncli = " & XLO_FORNCLI!focl_cd_FornCli & "", Dinamico
                If Not XLO_CLIENTECAPI.EOF Then
                    TxtRG.Text = XLO_CLIENTECAPI!clie_nr_Identidade
                    TxtOrgEmissor.Text = FunNulo(XLO_CLIENTECAPI!clie_tx_OrgEmissor)
                    MskDtNascimento.Value = Format(XLO_CLIENTECAPI!clie_dt_Nascimento, "DD/MM/YYYY")
                    
                    Select Case FunNuloVal(XLO_CLIENTECAPI!clie_nr_EstCivil) 'Linha de código alterada em 05/02/09 - Patrícia. Anterior: CboEstCivil.ListIndex = FunNuloVal(XLO_CLIENTECAPI!clie_nr_EstCivil)
                    Case 1
                        CboEstCivil.ListIndex = 0
                    Case 2
                        CboEstCivil.ListIndex = 1
                    Case 3
                        CboEstCivil.ListIndex = 2
                    Case 4
                        CboEstCivil.ListIndex = 3
                    Case 5
                        CboEstCivil.ListIndex = 4
                    Case 6
                        CboEstCivil.ListIndex = 5
                    End Select
                    
                    If Not IsNull(XLO_CLIENTECAPI!prof_cd_ProfTitular) Then
                        CboProfissao.BoundText = XLO_CLIENTECAPI!prof_cd_ProfTitular
                    End If
                    TxtNacionalidade.Text = XLO_CLIENTECAPI!clie_tx_Nacionalidade
                    If XLO_CLIENTECAPI!clie_vl_RendaFamiliar <> 0 Then
                        TxtRenda.Text = Format(XLO_CLIENTECAPI!clie_vl_RendaFamiliar, "standard")
                    End If
                    TxtEndResidencial.Text = XLO_CLIENTECAPI!clie_tx_EndResidencial
                    TxtBairroResidencial.Text = XLO_CLIENTECAPI!clie_tx_BairroResidencial
                    TxtCidadeResidencial.Text = XLO_CLIENTECAPI!clie_tx_MunResidencial
                    MskEstadoResidencial.Text = XLO_CLIENTECAPI!clie_tx_EstResidencial
                    MskCepResidencial.PromptInclude = False
                    MskCepResidencial.Text = XLO_CLIENTECAPI!clie_nr_CepResidencial
                    MskCepResidencial.PromptInclude = True
                    TxtEndCorresp.Text = XLO_CLIENTECAPI!clie_tx_endcorresp
                    TxtBairroCorresp.Text = XLO_CLIENTECAPI!clie_tx_bairrocorresp
                    TxtCidadeCorresp.Text = XLO_CLIENTECAPI!clie_tx_muncorresp
                    MskEstadoCorresp.Text = XLO_CLIENTECAPI!clie_tx_estcorresp
                    MskCepCorresp.PromptInclude = False
                    MskCepCorresp.Text = XLO_CLIENTECAPI!clie_nr_cepcorresp
                    MskCepCorresp.PromptInclude = True
                    TxtNomeConjuge.Text = XLO_CLIENTECAPI!clie_tx_Conjuge
                    MskDtNascConjuge.Value = Format(XLO_CLIENTECAPI!clie_dt_NascConjuge, "DD/MM/YYYY")
                    If Not IsNull(XLO_CLIENTECAPI!prof_cd_ProfConjuge) Then
                        CboProfConjuge.BoundText = XLO_CLIENTECAPI!prof_cd_ProfConjuge
                    End If
                    TxtNacConjuge.Text = XLO_CLIENTECAPI!clie_tx_NacConjuge
                    TxtRGConjuge.Text = XLO_CLIENTECAPI!clie_nr_IdConjuge
                    TxtOrgEmissorConjuge.Text = XLO_CLIENTECAPI!clie_tx_OrgEmissorConjuge
                    MskCPFConjuge.Text = XLO_CLIENTECAPI!clie_nr_CPFConjuge
                      
                    'EVERALDO - 06/12/2006 - Alterações COAF - INÍCIO
                    'Segundo Art 4º Parág. II inciso "a"
                    'No cadastro de clientes pessoa física incluir
                    '   - sexo (1 para masculino e 0 para feminino)
                    '   - filiação
                    '   - naturalidade
                    If Not IsNull(XLO_CLIENTECAPI!clie_lg_Sexo) Then
                        If XLO_CLIENTECAPI!clie_lg_Sexo = True Then
                            Me.optMasculino.Value = True
                            Me.optFeminino.Value = False
                        ElseIf XLO_CLIENTECAPI!clie_lg_Sexo = False Then
                            Me.optMasculino.Value = False
                            Me.optFeminino.Value = True
                        End If
                    Else
                            Me.optMasculino.Value = False
                            Me.optFeminino.Value = False
                    End If

                    If Not IsNull(XLO_CLIENTECAPI!clie_tx_FiliacaoPai) Then
                        Me.txtFiliacaoPai.Text = XLO_CLIENTECAPI!clie_tx_FiliacaoPai
                    End If
                    If Not IsNull(XLO_CLIENTECAPI!clie_tx_FiliacaoMae) Then
                        Me.txtFiliacaoMae.Text = XLO_CLIENTECAPI!clie_tx_FiliacaoMae
                    End If
                    If Not IsNull(XLO_CLIENTECAPI!clie_tx_Naturalidade) Then
                        Me.txtNaturalidade.Text = XLO_CLIENTECAPI!clie_tx_Naturalidade
                    End If
                    'Segundo Art 4º Item II inciso "d"
                    'No cadastro de clientes pessoa física incluir
                    '   - data de expedição do doc. de identificação e,
                    '   - dados do passaporte e carteira civil caso pessoa fisica estrangeira
                    If Not IsNull(XLO_CLIENTECAPI!clie_dt_ExpedicaoDocumento) Then
                        MskDtExpedicao.Value = XLO_CLIENTECAPI!clie_dt_ExpedicaoDocumento
                    End If
                    If Not IsNull(XLO_CLIENTECAPI!clie_nr_Passaporte) Then
                        Me.txtPassaporte.Text = XLO_CLIENTECAPI!clie_nr_Passaporte
                    End If
                    If Not IsNull(XLO_CLIENTECAPI!clie_nr_IdentidadeEstrangeira) Then
                        Me.txtDocEstrangeiro.Text = XLO_CLIENTECAPI!clie_nr_IdentidadeEstrangeira
                    End If
                    
                    If Not IsNull(XLO_CLIENTECAPI!clie_tx_paisresidencial) Then
                        txtPais.Text = XLO_CLIENTECAPI!clie_tx_paisresidencial
                    End If
                    'EVERALDO - 06/12/2006 - Alterações COAF - FIM
                    
                    XLO_CLIENTECAPI.Close
                End If
            End If
        End If
    End If
    XLO_FORNCLI.Close
    
    XLT_SQL = "SELECT BancoFornecedor.*, Bancos.banc_tx_descricao FROM BancoFornecedor LEFT JOIN Bancos ON Bancos.banc_cd_codigo = Bancofornecedor.banc_cd_codigo WHERE focl_cd_forncli = " & Chave
    subCarregaVetor XLT_SQL, Array("banc_cd_codigo", "banc_tx_descricao", "bafo_cd_Agencia", "bafo_nr_dag", "bafo_nm_Agencia", "bafo_nr_conta", "bafo_nr_dac", "bafo_tx_padrao", "bafo_cd_codigo", "bafo_tx_padrao"), Vetor3, TDBGrid3
End Sub

Sub PrepararInsercao()
    
    TxtCgc.Mask = ""
    TxtCgc.Text = ""
    If Pessoa = "F" Then
        TxtCgc.Mask = "###.###.###-##"
    Else
        TxtCgc.Mask = "##.###.###/####-##"
    End If
        
    TxtRazaoSocial.Text = ""
    TxtInscricaoEstadual.Text = ""
    TxtIMunicipal.Text = ""
    TxtContatoComercial.Text = ""
    TxtContatoFinanceiro.Text = ""
    TxtNomeFantasia.Text = ""
    TxtEndereco.Text = ""
    txtCodigoMunicipioIbge.Text = ""
    TxtNumPorta.Text = ""
    TxtBairro.Text = ""
    TxtCidade.Text = ""
    TxtEmail.Text = ""
    'Adicionado no dia 14/08/2018 - Samuel Jesus. Solicitação feita por Daniela para atender ao E-Social
    TxtPis.Mask = "###.###.###-##"
    txtNumEndereco.Text = ""
    
    MskEstado.Mask = ""
    MskEstado.Text = ""
    MskCep.Mask = ""
    MskCep.Text = ""
    MskCep.Mask = "##.###-###"
    
    MskFoneResid.Mask = ""
    MskFoneResid.Text = ""
    '*MskFoneResid.Mask = "(####)####-####"
    
    MskFax.Mask = ""
    MskFax.Text = ""
    '*MskFax.Mask = "(####)####-####"
        
    MskFoneCom.Mask = ""
    MskFoneCom.Text = ""
    '*MskFoneCom.Mask = "(####)####-####"
    
    MskCelular.Mask = ""
    MskCelular.Text = ""
    '*MskCelular.Mask = "(####)####-####"
    
    If FornClie = "F" Then 'Fornecedor
        MskDespRec.Mask = ""
        MskDespRec.Text = ""
        MskProvAdiant.Mask = ""
        MskProvAdiant.Text = ""
        MskCusto.Mask = ""
        MskCusto.Text = ""
        CboTipoPagamento.ListIndex = 0
        TxtIcms.Text = ""
        TxtCodAtividade.Text = ""
    ElseIf FornClie = "C" Then 'Cliente
        MskDespRec2.Mask = ""
        MskDespRec2.Text = ""
        MskCusto2.Mask = ""
        MskCusto2.Text = ""
        
        If XGT_NomeAplic = "CAPI" Then
            TxtRG.Text = ""
            TxtOrgEmissor.Text = ""
            CboProfissao.Text = ""
            CboEstCivil.ListIndex = 0
            TxtNacionalidade.Text = ""
            TxtRenda.Text = ""
            TxtEndResidencial.Text = ""
            TxtBairroResidencial.Text = ""
            TxtCidadeResidencial.Text = ""
            MskEstadoResidencial.Text = ""
            'MskCepResidencial.Text = ""
            TxtEndCorresp.Text = ""
            TxtBairroCorresp.Text = ""
            TxtCidadeCorresp.Text = ""
            MskEstadoCorresp.Text = ""
            'MskCepCorresp.Text = ""
            TxtNomeConjuge.Text = ""
            CboProfConjuge.Text = ""
            TxtNacConjuge.Text = ""
            TxtRGConjuge.Text = ""
            TxtOrgEmissorConjuge.Text = ""
            MskCPFConjuge.Text = "   .   .   -  "
        End If
    End If
End Sub

Sub SubGravaBancoFornecedor()
    Dim XLI_X As Integer
    
    TDBGrid3.MoveFirst
    For XLI_X = 0 To Vetor3.UpperBound(1)
        If TDBGrid3.Columns("CodBanco") <> "" Then 'Código do banco <> ""
            If TDBGrid3.Columns("Status") = "E" Then 'Excluiu um banco que já existia (DELETE)
                Sql = "DELETE from BancoFornecedor WHERE bafo_cd_codigo = " & TDBGrid3.Columns("CodBanco") & ""
                Conexao.Execute Sql
            Else 'Não excluiu um banco que já existia, mas pode ter alterado (UPDATE)
                Sql = "UPDATE Bancofornecedor " & _
                        "SET banc_cd_codigo=" & Vetor3(XLI_X, 0) & "," & _
                        "bafo_cd_agencia='" & Vetor3(XLI_X, 2) & "'," & _
                        "bafo_nr_dag='" & Vetor3(XLI_X, 3) & "', " & _
                        "bafo_nm_agencia='" & Vetor3(XLI_X, 4) & "'," & _
                        "bafo_nr_conta='" & Vetor3(XLI_X, 5) & "'," & _
                        "bafo_nr_dac='" & Vetor3(XLI_X, 6) & "', " & _
                        "bafo_tx_padrao='" & Vetor3(XLI_X, 7) & "' " & _
                        "WHERE bafo_cd_codigo = " & TDBGrid3.Columns("CodBanco") & ""
                Conexao.Execute Sql
            End If
        Else
            If TDBGrid3.Columns("Status") <> "E" Then 'Inseriu um novo código e não o excluiu (INSERT)
                Sql = "INSERT INTO Bancofornecedor(banc_cd_codigo,bafo_cd_agencia, bafo_nr_dag, bafo_nm_agencia,bafo_nr_conta, bafo_nr_dac, bafo_tx_padrao, focl_cd_forncli)" & _
                    " VALUES (" & CInt(Vetor3(XLI_X, 0)) & ",'" & Vetor3(XLI_X, 2) & "','" & Vetor3(XLI_X, 3) & "','" & Vetor3(XLI_X, 4) & "','" & Vetor3(XLI_X, 5) & "','" & Vetor3(XLI_X, 6) & "','" & Vetor3(XLI_X, 7) & "'," & CodigoDocumento & ")"
                Conexao.Execute Sql
            End If
        End If
        TDBGrid3.MoveNext
    Next
End Sub

Sub SubGravaResponsavelCoaf()

    If XInserir = "A" Then
        Conexao.Execute "DELETE FROM ResponsavelCoaf WHERE focl_cd_forncli=" & CodigoDocumento
    End If

    If Vetor2(0, 0) <> Empty Then
    '    subGravaVetor "ResponsavelCoaf", "focl_cd_forncli,recf_tx_Responsavel,recf_tx_Funcao,recf_tx_Cpf", Array("N", "T", "T", "T"), 4, Array(0, 1, 2, 3), "'" & CodigoDocumento & "', '" & TxtResponsavelCoaf.Text & "', '" & CboFuncaoRespCoaf.Text & "', '" & TxtCPFRespCoaf.Text & "'", Vetor2, TDBGrid2
        subGravaVetor "ResponsavelCoaf", "recf_tx_Responsavel,recf_tx_Funcao,recf_tx_Cpf, focl_cd_forncli", Array("T", "T", "T"), 3, Array(0, 1, 2), "'" & CodigoDocumento & "'", Vetor2, TDBGrid2
      '   subGravaVetor "DesImpForn", "desi_vl_percentual,desc_cd_desconto,focl_cd_forncli", Array("N", "N"), 2, Array(1, 2), "'" & CodigoDocumento & "'", Vetor1, TDBGrid1
    End If
   
End Sub

Private Sub CboAtividadePrincipal_LostFocus()
    If CboAtividadePrincipal.Text <> "" Then
        DatAtividadePrincipal.Recordset.bookmark = CboAtividadePrincipal.SelectedItem
        If Len(DatAtividadePrincipal.Recordset.Fields("cnse_tx_codigo")) <> 10 Then
            MsgBox "Este item do CNAE não é uma Sub-Classe, logo não pode ser atribuído ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboAtividadePrincipal.BoundText = MskAtividadePrincipal.Text
            Exit Sub
        End If
        MskAtividadePrincipal.Text = Right(Replace(DatAtividadePrincipal.Recordset.Fields("cnse_tx_codigo"), "-", ""), 8)
        'Right(DatAtividadePrincipal.Recordset.Fields("cnse_tx_codigo"), 9)
        
    End If
End Sub

Private Sub CboCusto_LostFocus()
    If CboCusto.Text <> "" Then
        DatCusto.Recordset.bookmark = CboCusto.SelectedItem
        If DatCusto.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboCusto.BoundText = MskCusto.Text
            'CboCusto.SetFocus
            Exit Sub
        End If
        MskCusto.Text = CboCusto.BoundText
    End If

End Sub


Private Sub CboCusto2_Change()
    If CboCusto2.Text <> "" Then
        DatCusto.Recordset.bookmark = CboCusto2.SelectedItem
        If DatCusto.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboCusto2.BoundText = MskCusto2.Text
            'CboCusto2.SetFocus
            Exit Sub
        End If
        MskCusto2.Text = CboCusto2.BoundText
    End If
End Sub

Private Sub CboDespRec_LostFocus()
    If CboDespRec.Text <> "" Then
        DatDespRec.Recordset.bookmark = CboDespRec.SelectedItem
        If DatDespRec.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboDespRec.BoundText = MskDespRec.Text
            'CboDespRec.SetFocus
            Exit Sub
        End If
        MskDespRec.Text = CboDespRec.BoundText
    End If

End Sub


Private Sub cboDespRec2_Change()
    If CboDesprec2.Text <> "" Then
        DatDespRec.Recordset.bookmark = CboDesprec2.SelectedItem
        If DatDespRec.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboDesprec2.BoundText = MskDespRec2.Text
            'CboDesprec2.SetFocus
            Exit Sub
        End If
        MskDespRec2.Text = CboDesprec2.BoundText
    End If
End Sub

Private Sub CboPcr_LostFocus()
    If CboPcr.Text <> "" Then
        DatPcr.Recordset.bookmark = CboPcr.SelectedItem
        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            Exit Sub
        End If
        MskPcr.Text = CboPcr.BoundText
    End If

End Sub


Private Sub CboProvAdiant_LostFocus()
    If CboProvAdiant.Text <> "" Then
        DatProvAdiant.Recordset.bookmark = CboProvAdiant.SelectedItem
        If DatProvAdiant.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído ao Fornecedor.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboProvAdiant.BoundText = MskProvAdiant.Text
            'CboProvAdiant.SetFocus
            Exit Sub
        End If
        MskProvAdiant.Text = CboProvAdiant.BoundText
    End If

End Sub


Private Sub chkFornecedorPrincipal_Click()
    If chkFornecedorPrincipal.Value = 1 Then
        chkFornecedorCritico.Enabled = True
    Else
        chkFornecedorCritico.Enabled = False
        If chkFornecedorCritico.Value = 1 Then
            chkFornecedorCritico.Value = Unchecked
        End If
    End If
End Sub

Private Sub CmdAlteraBanco_Click()
    On Error GoTo TrataErro
    
    XBanco = "A"
    
    DesabilitaBotoes ("B")
    TDBGrid3.Enabled = False
    
    PanCadBancoForn.Visible = True
    PanCadBancoForn.Enabled = True
     
    CboBanco.BoundText = TDBGrid3.Columns(0).Text
    TxtCodAgencia.Text = TDBGrid3.Columns(2).Text
    TxtDag.Text = TDBGrid3.Columns(3).Text
    TxtNomeAgencia.Text = TDBGrid3.Columns(4).Text
    TxtContaCorrente.Text = TDBGrid3.Columns(5).Text
    TxtDac.Text = TDBGrid3.Columns(6).Text
    
Exit Sub

TrataErro:

  funTrataErros (ComMensagem)


End Sub

Private Sub cmdAlterarRespCoaf_Click()
     
    On Error GoTo TrataErro
    
    XCoaf = "A"
    
    DesabilitaBotoes ("B")
    TDBGrid2.Enabled = False
    
    TxtResponsavelCoaf = TDBGrid2.Columns(0).Text
    CboFuncaoRespCoaf.Text = TDBGrid2.Columns(1).Text
    TxtCPFRespCoaf.Text = TDBGrid2.Columns(2).Text
    
    PanDadosRespCoaf.Visible = True
    PanDadosRespCoaf.Enabled = True
    
Exit Sub

TrataErro:

funTrataErros (ComMensagem)

End Sub

Private Sub CmdAtividadePrincipal_Click()
    CboAtividadePrincipal.BoundText = 0
    CboAtividadePrincipal.Text = ""
    MskAtividadePrincipal.Text = ""
End Sub

Private Sub CmdConfirmarRespCoaf_Click()

    If XCoaf = "I" Then
          subInsereRegistroVetor Array(TxtResponsavelCoaf.Text, CboFuncaoRespCoaf.Text, TxtCPFRespCoaf.Text, ""), Vetor2, TDBGrid2
    ElseIf XCoaf = "A" Then
        TDBGrid2.Columns(0).Text = TxtResponsavelCoaf.Text
        TDBGrid2.Columns(1).Text = CboFuncaoRespCoaf.Text
        TDBGrid2.Columns(2).Text = TxtCPFRespCoaf.Text
        'TDBGrid2.Columns(3).Text = CodigoDocumento
        TDBGrid2.Update
    End If
    
    PanDadosRespCoaf.Enabled = False
    PanDadosRespCoaf.Visible = False
    
    If Vetor2(0, 0) <> Empty Then
        HabilitaBotoes ("B")
    Else
        CmdInsereBanco.Enabled = True
    End If
    
    TDBGrid2.Enabled = True
    
End Sub

Private Sub CmdCopiar_Click()
    TxtEndCorresp.Text = TxtEndResidencial.Text
    TxtBairroCorresp.Text = TxtBairroResidencial.Text
    TxtCidadeCorresp.Text = TxtCidadeResidencial.Text
    MskEstadoCorresp.Text = MskEstadoResidencial.Text
    MskCepCorresp.Text = MskCepResidencial.Text
End Sub

Private Sub CmdCopiarComercial_Click()
    TxtEndCorresp.Text = TxtEndereco.Text + ", " + TxtNumPorta.Text
    TxtBairroCorresp.Text = TxtBairro.Text
    TxtCidadeCorresp.Text = TxtCidade.Text
    MskEstadoCorresp.Text = MskEstado.Text
    MskCepCorresp.Text = MskCep.Text
End Sub


Private Sub cmdExcluirRespCoaf_Click()

    DesabilitaBotoes ("B")
   
    TDBGrid2.Enabled = False

    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor2, TDBGrid2
    End If
        
    TDBGrid2.Enabled = True
    
    If Vetor2(0, 0) <> Empty Then
        HabilitaBotoes ("B")
    Else
        cmdInserirRespCoaf.Enabled = True
    End If
 
    Exit Sub
    
RotuloErro:
    'Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next

End Sub

Private Sub cmdInserirRespCoaf_Click()
    XCoaf = "I"
    PanDadosRespCoaf.Visible = True
    PanDadosRespCoaf.Enabled = True
    TxtResponsavelCoaf.Text = ""
    TxtResponsavelCoaf.SetFocus
    CboFuncaoRespCoaf.Text = ""
    'Adicionado no dia 14/08/2018 - Samuel Jesus. Solicitação feita por Daniela para atender ao E-Social
    'Comentado dia 12/08/2021
'    If FunNulo(TxtCPFRespCoaf.Text) <> "___.___.___-__" Then
'        ResFormulario!focl_nr_pis = FunNulo(TxtPis.Text)
'    Else
'        ResFormulario!focl_nr_pis = ""
'    End If

    
    If TxtCPFRespCoaf <> "___.___.___-__" Then
        TxtCPFRespCoaf.Text = "___.___.___-__"
    End If
    
End Sub

Private Sub CmdLimparCusto_Click()
    CboCusto.BoundText = 0
    CboCusto.Text = ""
    MskCusto.Text = ""
End Sub

Private Sub CmdLimparCusto2_Click()
    CboCusto2.BoundText = 0
    CboCusto2.Text = ""
    MskCusto2.Text = ""
End Sub


Private Sub CmdLimparDespRec_Click()
    CboDespRec.BoundText = 0
    CboDespRec.Text = ""
    MskDespRec.Text = ""
End Sub


Private Sub CmdLimparDespRec2_Click()
    CboDesprec2.BoundText = 0
    CboDesprec2.Text = ""
    MskDespRec2.Text = ""
End Sub


Private Sub CmdLimparPcr_Click()
    CboPcr.BoundText = 0
    CboPcr.Text = ""
    MskPcr.PromptInclude = False
    MskPcr.Text = ""
    MskPcr.PromptInclude = True
End Sub

Private Sub CmdLimparPcr2_Click()
    CboPcr2.BoundText = 0
    CboPcr2.Text = ""
    MskPcr2.PromptInclude = False
    MskPcr2.Text = ""
    MskPcr2.PromptInclude = True
End Sub


Private Sub CmdLimparProvAdiant_Click()
    CboProvAdiant.BoundText = 0
    CboProvAdiant.Text = ""
    MskProvAdiant.Text = ""
End Sub


Private Sub CmdRetornarRespCoaf_Click()
    
    PanDadosRespCoaf.Visible = False
    PanDadosRespCoaf.Enabled = False
       
    TDBGrid2.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
       HabilitaBotoes ("B")
    Else
       cmdInserirRespCoaf.Enabled = True
    End If
End Sub

Private Sub CmdSairBanco_Click()
    PanCadBancoForn.Visible = False
    PanCadBancoForn.Enabled = False
       
    TDBGrid3.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
       HabilitaBotoes ("B")
    Else
       CmdInsereBanco.Enabled = True
    End If
End Sub
Private Sub CmdConfirmar_Click()
    Dim Xvalor
    Dim Bcpadrao As String

   ' On Error GoTo TrataErro

    'Função para verificar preenchimento dos campos obrigatórios
    'If ChecarCampos Then Exit Sub
    
    If XBanco = "I" Then
        Bcpadrao = "N"
        subInsereRegistroVetor Array(CboBanco.BoundText, CboBanco.Text, TxtCodAgencia.Text, TxtDag.Text, TxtNomeAgencia.Text, TxtContaCorrente.Text, TxtDac.Text, Bcpadrao, "", ""), Vetor3, TDBGrid3
    ElseIf XBanco = "A" Then
        TDBGrid3.Columns(0).Text = CboBanco.BoundText
        TDBGrid3.Columns(1).Text = CboBanco.Text
        TDBGrid3.Columns(2).Text = TxtCodAgencia.Text
        TDBGrid3.Columns(3).Text = TxtDag.Text
        TDBGrid3.Columns(4).Text = TxtNomeAgencia.Text
        TDBGrid3.Columns(5).Text = TxtContaCorrente.Text
        TDBGrid3.Columns(6).Text = TxtDac.Text
        TDBGrid3.Update
    End If
    
    PanCadBancoForn.Enabled = False
    PanCadBancoForn.Visible = False
    If Vetor3(0, 0) <> Empty Then
        HabilitaBotoes ("B")
    Else
        CmdInsereBanco.Enabled = True
    End If
    
    TDBGrid3.Enabled = True
End Sub
Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub CmdExcluir_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor1, TDBGrid1
        If Vetor1.Count(1) = 1 Then
            If IsEmpty(Vetor1(0, 0)) Then
                CmdExcluir.Enabled = False
                TDBGrid1.Columns(1).Locked = True
            End If
        End If
    End If
   
End Sub

Private Sub CmdGravar_Click()
    Dim Sql As String
    Dim XLT_QUALIFICACAO As String
    Dim TipoPag As Byte
    Dim ResFornCli As Object, ResPadrao As Object
    Dim linx As Integer, XCodEmpresa As Integer
    Dim PosElem As Integer 'Linha do elemento no vetor
    Dim XDespRec As String, XCusto As String
    Dim XProvAdiant As String, XPcr As String
    Dim ExisteReg As Boolean
    Dim XLO_CAPI As Object      'Tabela de clientes só cadastrada através do cliente
    Dim XLI_X As Integer
    Dim XLI_I As Integer
    XLI_I = 0

    'On Error GoTo TrataErro
    
    If ChecarCampos Then Exit Sub
         
         
     'Verificando se CPF ou CNPJ foi preenchido - Daniel Coelho - 14/03/2012
    If TxtCgc.Text = "___.___.___-__" Or TxtCgc.Text = "__.___.___/____-__" Then
        If Pessoa = "F" Then
            MsgBox "O preenchimento do CPF é obgrigatório!", vbCritical, "ATENÇÃO"
        Else
            MsgBox "O preenchimento do CNPJ é obgrigatório!", vbCritical, "ATENÇÃO"
        End If
       Exit Sub
    End If
    
    If MskCep.Text = "__.___-___" Then
        MsgBox "O preenchimento do CEP é obgrigatório!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    If FornClie = "F" Then
        If FunObrigatorioTXT(TxtEndereco, "O preenchimento do endereço é obrigatório") Then Exit Sub
        If FunObrigatorioTXT(TxtBairro, "O preenchimento do bairro é obrigatório") Then Exit Sub
        If FunObrigatorioTXT(TxtCidade, "O preenchimento do cidade é obrigatório") Then Exit Sub
        If FunObrigatorioTXT(txtPais, "O preenchimento do país é obrigatório") Then Exit Sub
    End If
         
'MsgBox "Marca 1"
         
    If EntrouImp = False Then
        For XLI_X = 0 To Vetor1.UpperBound(1)
            If Vetor1(0, 0) <> Empty Then
                If Vetor1(XLI_X, 1) = "" Then
                    MsgBox "O campo Valor é obrigatório!", vbCritical, "ATENÇÃO"
                    Exit Sub
                End If
            End If
        Next
    End If
'MsgBox "Marca 2"
    If CboTipoPagamento.ListIndex > -1 Then
        TipoPag = CByte(CboTipoPagamento.ListIndex)
    Else
        TipoPag = 0
    End If
'MsgBox "Marca 3"
    
    'If (TipoPag = "2" Or TipoPag = "4") And XInserir = "I" Then
    If (TipoPag = "2") And XInserir = "I" Then
        If Vetor3(0, 0) <> Empty And XBancoPadrao = "N" Then
            MsgBox "É necessário indicar um Banco Padrão!", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf Vetor3(0, 0) = Empty Then
            MsgBox "É necessário cadastrar um Banco Padrão!", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If
'MsgBox "Marca 4"
    
    If XOrigem <> "PanCessao" Then
      Conexao.BeginTrans
    End If

'MsgBox "Marca 5"

'Qualificações
    If OptQualificado.Value Then
       XLT_QUALIFICACAO = "Q"
  ElseIf OptQualificadoRestrincao Then
    XLT_QUALIFICACAO = "QR"
  ElseIf OptNaoQualificado Then
    XLT_QUALIFICACAO = "NQ"
 ElseIf OptFornExclusivo Then
    XLT_QUALIFICACAO = "FE"
  
  End If
 
'MsgBox "Marca 6"
    If XInserir = "I" Then
        'Abre o banco de Dados
        SubQOpenRecordset ResFormulario, "SELECT * FROM FornClientes WHERE focl_cd_forncli= -1", Dinamico
        ResFormulario.AddNew
        'Os campos que só serão gravados uma vez podem ser colocados aqui
        'Ex.: Data de Cadastro
        ResFormulario!focl_tx_classe = FornClie
        'ResFormulario!focl_tx_Tipo = Pessoa
        
    Else    'Alteração
        SubQOpenRecordset ResFormulario, "SELECT * FROM fornClientes WHERE focl_cd_forncli = " & CodigoDocumento, Dinamico
    End If
'MsgBox "Marca 7"
    If FornClie = "F" Then 'Fornecedor
        ResFormulario!focl_tx_TipoPag = TipoPag
    End If
'MsgBox "Marca 8"
    ResFormulario!focl_tx_Tipo = Pessoa
    ResFormulario!focl_tx_fantasia = FunNulo(TxtNomeFantasia.Text)
    ResFormulario!focl_tx_RazaoSocial = FunNulo(TxtRazaoSocial.Text)
    ResFormulario!focl_tx_endereco = FunNulo(TxtEndereco.Text)
    ResFormulario!focl_cd_MunicipioIbge = FunNulo(txtCodigoMunicipioIbge.Text)
    ResFormulario!focl_nr_porta = FunNulo(TxtNumPorta.Text)
    ResFormulario!focl_tx_Bairro = FunNulo(TxtBairro.Text)
    ResFormulario!focl_tx_cidade = FunNulo(TxtCidade.Text)
    ResFormulario!focl_tx_email = FunNulo(TxtEmail.Text)
    ResFormulario!focl_tx_estado = FunNulo(MskEstado.Text)
    ResFormulario!focl_tx_cep = FunNulo(MskCep.Text)
    ResFormulario!focl_tx_fone = FunNulo(MskFoneResid.Text)
    ResFormulario!focl_tx_fone1 = FunNulo(MskFoneCom.Text)
    ResFormulario!focl_tx_fax = FunNulo(MskFax.Text)
    ResFormulario!focl_tx_Celular = FunNulo(MskCelular.Text)
    ResFormulario!focl_tx_CGCCPF = TxtCgc.Text
    ResFormulario!focl_tx_iestadual = FunNulo(TxtInscricaoEstadual.Text)
    ResFormulario!focl_tx_imunicipal = FunNulo(TxtIMunicipal.Text)
    ResFormulario!focl_tx_contatoCom = FunNulo(TxtContatoComercial.Text)
    ResFormulario!focl_tx_ContatoFin = FunNulo(TxtContatoFinanceiro.Text)
    ResFormulario!focl_vl_Icms = FunNuloVal(TxtIcms.Text)
    ResFormulario!focl_nr_codatividade = FunNulo(TxtCodAtividade.Text)
    ResFormulario!focl_tx_Qualificacoes = FunNulo(XLT_QUALIFICACAO)
    ResFormulario!focl_tx_Observacoes = FunNulo(TxtObservacoes.Text)
    ResFormulario!focl_tx_pais = FunNulo(txtPais.Text)
    ResFormulario!focl_tx_EnderecoNumero = funNull(txtNumEndereco.Text)
    ResFormulario!focl_nr_NaturezaJuridica = FunNulo(cboNaturezaJuridica.ListIndex)
    ResFormulario!focl_nr_RegimeApuracao = FunNulo(cboRegimeApuracao.ListIndex)
    If (chkContICMS.Value = 1) Then
        ResFormulario!focl_bl_ContribuinteICMS = 1
    Else
        ResFormulario!focl_bl_ContribuinteICMS = 0
    End If
    'Henrique Lima - 16/04/09 - Salva o valor do Super Simples no banco
    If (chkSuperSimples.Value) Then
        ResFormulario!focl_tx_SuperSimples = "S"
    Else
        ResFormulario!focl_tx_SuperSimples = "N"
    End If
    
    'Cássio Medeiros - 16/04/13
    If (chkFornecedorPrincipal.Value) Then
        ResFormulario!focl_tx_fornecedorPrincipal = "S"
    Else
        ResFormulario!focl_tx_fornecedorPrincipal = "N"
    End If
    
    If (DtpPesquisaUltimaPesquisa.Value <> "") Then
        ResFormulario!focl_dt_pesquisaUltimaQualificacao = Format$(DtpPesquisaUltimaPesquisa.Value, "dd/mm/yyyy")
    Else
        ResFormulario!focl_dt_pesquisaUltimaQualificacao = Null
    End If
    
    'Everaldo Sena - 30/05/12
    If (chkFornecedorCritico.Value) Then
        ResFormulario!focl_tx_fornecedorcritico = "S"
    Else
        ResFormulario!focl_tx_fornecedorcritico = "N"
    End If
        
    If XGT_NomeAplic = "CAPI" Or XGB_PCAPI Then
        If Pessoa = "J" Then
            If Not IsNull(CboAtividadePrincipal.SelectedItem) Then
                DatAtividadePrincipal.Recordset.bookmark = CboAtividadePrincipal.SelectedItem
                ResFormulario!cnsc_cd_cnaeSubclasse = FunNulo(DatAtividadePrincipal.Recordset.Fields("cnse_cd_cnaesecao"))
            Else
                ResFormulario!cnsc_cd_cnaeSubclasse = Null
            End If
        End If
    End If
    If FunNulo(TxtPis.Text) <> "___.___.___-__" Then
        ResFormulario!focl_nr_pis = FunNulo(TxtPis.Text)
    Else
        ResFormulario!focl_nr_pis = ""
    End If
    
    
    If FunNulo(TxtNit.Text) <> "___.___.___-__" Then
        ResFormulario!focl_nr_nit = FunNulo(TxtNit.Text)
    Else
        ResFormulario!focl_nr_nit = ""
    End If
    ResFormulario.Update
'MsgBox "Marca 9"
    'Colocar esta condição se precisar setar o registro criado
    'por alguma razão. Ex.: Pegar a chave quando for auto numeração
    If XInserir = "I" Then
        CodigoDocumento = ResFormulario!focl_cd_FornCli
        Chave = ResFormulario!focl_cd_FornCli
    End If
'MsgBox "Marca 10"
     'O código abaixo não funcionou para SQL Server. Ele não aceita uma alteração depois do Update
    'Atualiza o CGC com o código gerado, caso o usuário não tenha definido o CGC
'    TxtCgc.PromptInclude = False
'    If TxtCgc.Text = "" Then
'        TxtCgc.Text = Chave
'        TxtCgc.PromptInclude = True
'        ResFormulario!focl_tx_cgccpf = TxtCgc.Text
'        ResFormulario.Update
'    End If
'    TxtCgc.PromptInclude = True
    
    ResFormulario.Close
'MsgBox "Marca 11"
    '********
    'Atualização do CGC funciona tanto para Sql Server quanto para Access
    TxtCgc.PromptInclude = False
    If TxtCgc.Text = "" Then
        TxtCgc.Text = Chave
        TxtCgc.PromptInclude = True
        Conexao.Execute "UPDATE FornClientes " & _
            "SET focl_tx_cgccpf = '" & TxtCgc.Text & "' WHERE focl_cd_forncli = " & CodigoDocumento
    End If
    TxtCgc.PromptInclude = True
    '********
'MsgBox "Marca 12"
    'Gravar dados do banco
    If Vetor3(0, 0) <> Empty Then
        SubGravaBancoFornecedor
    End If
    
    If XGT_NomeAplic = "CAPI" Then
       If Pessoa = "J" Then
        'If Vetor2(0, 0) <> Empty Then
        SubGravaResponsavelCoaf
        'End If
        End If
    End If
'MsgBox "Marca 13"
'    If XInserir = "A" Then
'        Conexao.Execute "DELETE from BancoFornecedor WHERE focl_cd_forncli = " & CodigoDocumento
'    End If
'    If Vetor3(0, 0) <> Empty Then
'        subGravaVetor "Bancofornecedor", "banc_cd_codigo,bafo_cd_agencia,bafo_nm_agencia,bafo_nr_conta, bafo_tx_padrao, focl_cd_forncli", Array("T", "T", "T", "T", "T", "N"), 5, Array(0, 2, 3, 4, 5), "" & CodigoDocumento & "", Vetor3, TDBGrid3
'    End If
    
    'Gravar Dados da Tabela de PlCCustoForn
    If (MskDespRec.Text <> "") Or (MskDespRec2.Text <> "") Then
        If FornClie = "F" Then 'Caso seja fornecedor
            DatDespRec.Recordset.bookmark = CboDespRec.SelectedItem
        Else
            DatDespRec.Recordset.bookmark = CboDesprec2.SelectedItem
        End If
        XDespRec = DatDespRec.Recordset.Fields("plco_cd_conta")
    End If
'MsgBox "Marca 14"
    If (MskCusto.Text <> "") Or (MskCusto2.Text <> "") Then
        If FornClie = "F" Then 'Caso seja fornecedor
            DatCusto.Recordset.bookmark = CboCusto.SelectedItem
        Else
            DatCusto.Recordset.bookmark = CboCusto2.SelectedItem
        End If
        XCusto = DatCusto.Recordset.Fields("plco_cd_conta")
    End If
'MsgBox "Marca 15"
    If (MskProvAdiant.Text <> "") Then
        If FornClie = "F" Then 'Fornecedor
            DatProvAdiant.Recordset.bookmark = CboProvAdiant.SelectedItem
            XProvAdiant = DatProvAdiant.Recordset.Fields("plco_cd_conta")
        ElseIf FornClie = "C" Then 'Cliente
            XProvAdiant = ""
        End If
    End If
'MsgBox "Marca 16"
    If (CboPcr.Text <> "") Or (CboPcr2.Text <> "") Then
        If FornClie = "F" Then 'Caso seja fornecedor
            DatPcr.Recordset.bookmark = CboPcr.SelectedItem
        Else
            DatPcr.Recordset.bookmark = CboPcr2.SelectedItem
        End If
        XPcr = DatPcr.Recordset.Fields("cere_cd_pcr")
    End If
'MsgBox "Marca 17"
    If XInserir = "I" Then
        Sql = "INSERT INTO PlCCustoForncli(focl_cd_forncli,plco_cd_contarecdesp, plco_cd_contacusto, plco_cd_contaProvAdiant, cere_cd_pcr,empr_cd_empresa)" & _
                          " VALUES (" & CodigoDocumento & "," & FunNuloBanco(XDespRec) & "," & FunNuloBanco(XCusto) & "," & FunNuloBanco(XProvAdiant) & "," & FunNuloBanco(XPcr) & "," & PCodEmpresa & ")"
        Conexao.Execute Sql
    Else
        SubQOpenRecordset ResFormulario, "SELECT * FROM PlCCustoFornCli WHERE focl_cd_forncli = " & CodigoDocumento & " and empr_cd_empresa = " & PCodEmpresa, Dinamico
'MsgBox "Marca 18"
        If ResFormulario.EOF Then
            ResFormulario.AddNew
            ResFormulario!empr_cd_Empresa = PCodEmpresa
            ResFormulario!focl_cd_FornCli = CodigoDocumento
        End If
'MsgBox "Marca 19"
        ResFormulario!plco_cd_ContaRecDesp = funNull(XDespRec)
        ResFormulario!plco_cd_ContaCusto = funNull(XCusto)
        ResFormulario!plco_cd_contaprovadiant = funNull(XProvAdiant)
        ResFormulario!cere_cd_Pcr = funNull(XPcr)
        
        ResFormulario.Update
        ResFormulario.Close
    End If
'MsgBox "Marca 20"

    If XInserir = "I" Then
        'Prepara o novo fornecedor para ser acessado por todas as empresas
        'Copia tabela de relacionamento PlCCustoFornCli atualizando os códigos para a nova empresa
        SubQOpenRecordset ResPadrao, "SELECT empr_cd_empresa FROM Empresas where empr_cd_empresa <>" & PCodEmpresa, Estatico
'MsgBox "Marca 21"
        While Not ResPadrao.EOF
            XCodEmpresa = ResPadrao!empr_cd_Empresa
            Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa) " & _
                  "VALUES (" & CodigoDocumento & "," & XCodEmpresa & ")"
            
            Conexao.Execute Sql
            ResPadrao.MoveNext
        Wend
        ResPadrao.Close
'MsgBox "Marca 22"
'        If XGT_NomeAplic = "CAPI" Then
'            Sql = "INSERT INTO Clientes(focl_cd_forncli, clie_nr_identidade, clie_tx_orgemissor, clie_dt_nascimento, clie_nr_estcivil, prof_cd_proftitular, clie_tx_nacionalidade, " & _
'                        "clie_vl_rendafamiliar, clie_tx_endresidencial, clie_tx_bairroresidencial, clie_tx_munresidencial, clie_tx_estresidencial, " & _
'                        "clie_nr_cepresidencial, clie_tx_endcorresp, clie_tx_bairrocorresp, clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp, " & _
'                        "clie_tx_conjuge, clie_dt_nascconjuge, prof_cd_ProfConjuge, clie_tx_nacconjuge, clie_nr_idconjuge, clie_tx_orgemissorconjuge, clie_nr_cpfconjuge) " & _
'                " VALUES (" & CodigoDocumento & ",'" & FunNulo(TxtRG.Text) & "','" & FunNulo(TxtOrgEmissor.Text) & "'," & FunNuloData(MskDtNascimento.Value, NomeSgbd) & ",'" & CboEstCivil.ListIndex & "'," & FunNuloBancoVal(CboProfissao.BoundText) & "," & _
'                        "'" & TxtNacionalidade.Text & "'," & FunNuloVal(FunTrataFloat(TxtRenda.Text)) & ",'" & TxtEndResidencial.Text & "','" & TxtBairroResidencial.Text & "'," & _
'                        "'" & TxtCidadeResidencial.Text & "','" & MskEstadoResidencial.Text & "','" & MskCepResidencial.Text & "','" & TxtEndCorresp.Text & "','" & TxtBairroCorresp.Text & "'," & _
'                        "'" & TxtCidadeCorresp.Text & "','" & MskEstadoCorresp.Text & "', '" & MskCepCorresp.Text & "','" & TxtNomeConjuge.Text & "'," & FunNuloData(MskDtNascConjuge.Value, NomeSgbd) & "," & _
'                        "" & FunNuloBancoVal(CboProfConjuge.BoundText) & ",'" & TxtNacConjuge.Text & "','" & TxtRGConjuge.Text & "','" & TxtOrgEmissorConjuge.Text & "','" & MskCPFConjuge.Text & "')"
'
'            Conexao.Execute Sql
'        End If
'    ElseIf XInserir = "A" And XGT_NomeAplic = "CAPI" Then
'        Conexao.Execute "UPDATE Clientes " & _
'                "SET clie_nr_identidade = '" & TxtRG.Text & "',clie_tx_orgemissor='" & TxtOrgEmissor.Text & "'," & _
'                "clie_dt_nascimento=" & FunNuloData(MskDtNascimento.Value, NomeSgbd) & ",clie_nr_estcivil='" & CboEstCivil.ListIndex & "'," & _
'                "prof_cd_proftitular=" & FunNuloBancoVal(CboProfissao.BoundText) & ",clie_tx_nacionalidade='" & TxtNacionalidade.Text & "'," & _
'                "clie_vl_rendafamiliar=" & FunNuloBancoVal(FunTrataFloat(TxtRenda.Text)) & ",clie_tx_endresidencial='" & TxtEndResidencial.Text & "'," & _
'                "clie_tx_bairroresidencial='" & TxtBairroResidencial.Text & "',clie_tx_munresidencial='" & TxtCidadeResidencial.Text & "'," & _
'                "clie_tx_estresidencial='" & MskEstadoResidencial.Text & "',clie_nr_cepresidencial='" & MskCepResidencial.Text & "'," & _
'                "clie_tx_endcorresp='" & TxtEndCorresp.Text & "',clie_tx_bairrocorresp='" & TxtBairroCorresp.Text & "'," & _
'                "clie_tx_muncorresp='" & TxtCidadeCorresp.Text & "',clie_tx_estcorresp='" & MskEstadoCorresp.Text & "'," & _
'                "clie_nr_cepcorresp='" & MskCepCorresp.Text & "',clie_tx_conjuge='" & TxtNomeConjuge.Text & "'," & _
'                "clie_dt_nascconjuge=" & FunNuloData(MskDtNascConjuge.Value, NomeSgbd) & ",prof_cd_profconjuge=" & FunNuloBancoVal(CboProfConjuge.BoundText) & "," & _
'                "clie_tx_nacconjuge='" & TxtNacConjuge.Text & "',clie_nr_idconjuge='" & TxtRGConjuge.Text & "'," & _
'                "clie_tx_orgemissorconjuge='" & TxtOrgEmissorConjuge.Text & "',clie_nr_cpfconjuge='" & MskCPFConjuge.Text & "' " & _
'                "WHERE focl_cd_forncli = " & CodigoDocumento
    End If
        
    '*******TABELA DE CLIENTES********
    If XGT_NomeAplic = "CAPI" Then
        SubQOpenRecordset XLO_CAPI, "SELECT * FROM Clientes WHERE focl_cd_forncli = " & CodigoDocumento, Dinamico
'MsgBox "Marca 23"
        If XLO_CAPI.EOF Then
            XLO_CAPI.AddNew
        End If
'MsgBox "Marca 24"
        XLO_CAPI!focl_cd_FornCli = CodigoDocumento
        XLO_CAPI!clie_nr_Identidade = FunNulo(TxtRG.Text)
        XLO_CAPI!clie_tx_OrgEmissor = FunNulo(TxtOrgEmissor.Text)
        XLO_CAPI!clie_dt_Nascimento = MskDtNascimento.Value
    'Inserido +1 para resolver o erro com os saltos do combo que nao grvava a opção selecionada
    'Daniel Coelho - 06/07/2011
        'If CboEstCivil.ListIndex > 0 Then
            XLO_CAPI!clie_nr_EstCivil = CboEstCivil.ListIndex + 1
        'End If


   
'MsgBox "Marca 25"
        XLO_CAPI!prof_cd_ProfTitular = funNull(CboProfissao.BoundText)
        XLO_CAPI!clie_tx_Nacionalidade = FunNulo(TxtNacionalidade.Text)
        
        'EVERALDO - 06/12/2006 - Alterações COAF - INÍCIO
        'Segundo Art 4º Parág. II inciso "a"
        'No cadastro de clientes pessoa física incluir
        '   - sexo (1 para masculino e 0 para feminino)
        '   - filiação
        '   - naturalidade
        If Me.optMasculino.Value = True Then
            XLO_CAPI!clie_lg_Sexo = 1
        ElseIf Me.optFeminino.Value = True Then
            XLO_CAPI!clie_lg_Sexo = 0
        End If
        XLO_CAPI!clie_tx_FiliacaoPai = FunNulo(Me.txtFiliacaoPai.Text)
        XLO_CAPI!clie_tx_FiliacaoMae = FunNulo(Me.txtFiliacaoMae.Text)
        XLO_CAPI!clie_tx_Naturalidade = FunNulo(Me.txtNaturalidade.Text)
        'Segundo Art 4º Item II inciso "d"
        'No cadastro de clientes pessoa física incluir
        '   - data de expedição do doc. de identificação e,
        '   - dados do passaporte e carteira civil caso pessoa fisica estrangeira
        XLO_CAPI!clie_dt_ExpedicaoDocumento = MskDtExpedicao.Value
        XLO_CAPI!clie_nr_Passaporte = FunNulo(Me.txtPassaporte.Text)
        XLO_CAPI!clie_nr_IdentidadeEstrangeira = FunNulo(Me.txtDocEstrangeiro.Text)
        'EVERALDO - 06/12/2006 - Alterações COAF - FIM
        
        XLO_CAPI!clie_vl_RendaFamiliar = funNull(FunTrataFloat(TxtRenda.Text))
        XLO_CAPI!clie_tx_EndResidencial = FunNulo(TxtEndResidencial.Text)
        XLO_CAPI!clie_tx_BairroResidencial = FunNulo(TxtBairroResidencial.Text)
        XLO_CAPI!clie_tx_MunResidencial = FunNulo(TxtCidadeResidencial.Text)
        XLO_CAPI!clie_tx_EstResidencial = FunNulo(MskEstadoResidencial.Text)
        XLO_CAPI!clie_nr_CepResidencial = FunNulo(MskCepResidencial.Text)
        XLO_CAPI!clie_tx_endcorresp = FunNulo(TxtEndCorresp.Text)
        XLO_CAPI!clie_tx_bairrocorresp = FunNulo(TxtBairroCorresp.Text)
        XLO_CAPI!clie_tx_muncorresp = FunNulo(TxtCidadeCorresp.Text)
        XLO_CAPI!clie_tx_estcorresp = FunNulo(MskEstadoCorresp.Text)
        XLO_CAPI!clie_nr_cepcorresp = FunNulo(MskCepCorresp.Text)
        XLO_CAPI!clie_tx_Conjuge = FunNulo(TxtNomeConjuge.Text)
        XLO_CAPI!clie_dt_NascConjuge = MskDtNascConjuge.Value
        XLO_CAPI!prof_cd_ProfConjuge = funNull(CboProfConjuge.BoundText)
        XLO_CAPI!clie_tx_NacConjuge = FunNulo(TxtNacConjuge.Text)
        XLO_CAPI!clie_nr_IdConjuge = FunNulo(TxtRGConjuge.Text)
        XLO_CAPI!clie_tx_OrgEmissorConjuge = FunNulo(TxtOrgEmissorConjuge.Text)
        XLO_CAPI!clie_nr_CPFConjuge = FunNulo(MskCPFConjuge.Text)
        XLO_CAPI!clie_tx_paisresidencial = FunNulo(txtPais.Text)
        
        XLO_CAPI.Update
        XLO_CAPI.Close ' inserido em 10/10/08 - Patrícia
    End If
'MsgBox "Marca 26"

    If EntrouImp = False Then 'Não gravará os dados de impostos caso o usuário não entre na orelha destinada ao mesmo
        If XInserir = "A" Then
            Conexao.Execute "DELETE FROM DesImpForn WHERE focl_cd_forncli=" & CodigoDocumento
        End If
        If Vetor1(0, 0) <> Empty Then
            subGravaVetor "DesImpForn", "desi_vl_percentual,desc_cd_desconto,focl_cd_forncli", Array("N", "N"), 2, Array(1, 2), "'" & CodigoDocumento & "'", Vetor1, TDBGrid1
        End If
    End If
'MsgBox "Marca 27"
    If XOrigem <> "PanCessao" Then
      Conexao.CommitTrans
    End If
'MsgBox "Marca 28"
    'Cadastro do Fornecedor chamado da tela de Cadastro de Notas Fiscais
    XGV_FornNotaFiscal = CodigoDocumento
    
    'Cadastro de Ramos de Atividades
'     Conexao.BeginTrans
'
'     Conexao.Execute "DELETE FROM FornCliente_RamosAtividades WHERE focl_cd_forncli = " & CodigoDocumento
'
'     For XLI_I = 0 To lsbRamosAtividades.ListCount - 1
'
'        'Cadastra os ramos de atividades do fornecedor
'        If lsbRamosAtividades.Selected(XLI_I) = True Then
'            Conexao.Execute "INSERT INTO FornCliente_RamosAtividades (focl_cd_forncli, raat_cd_ramoatividade) VALUES (" & CodigoDocumento & "," & CInt(RTrim(LTrim(Split(lsbRamosAtividades.List(XLI_I), "-")(0)))) & ")"
'        End If
'
'     Next XLI_I
'
'     Conexao.CommitTrans
        
    Unload Formulario
    
    Exit Sub
    
TrataErro:
    If Err.Number = NumberTravaRegistro And Err.Description = DescriptionTravaRegistro Then
        If MsgBox("O campo está sendo alterado por outro usuário. Deseja tentar novamente?", vbCritical + vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
            Resume
        Else
            Unload Me
            Exit Sub
        End If
    ElseIf Err.Number = NumberRegistroExcluido And Err.Description = DescriptionRegistroExcluido Then
        MsgBox "O campo foi excluído por outro usuário!", vbCritical, "ATENÇÃO"
        Unload Me
        Exit Sub
    Else
        funTrataErros (ComMensagem)
    End If
    
End Sub
Private Sub CmdInsere_Click()
    If CboDescontosImpostos.Text <> "" Then
        DatDescontosImpostos.Recordset.bookmark = CboDescontosImpostos.SelectedItem
        
        If funExisteChave(1, Array("2"), Array(DatDescontosImpostos.Recordset.Fields("desc_cd_desconto")), Vetor1, TDBGrid1) Then
            MsgBox "Esse desconto já foi cadastrado", vbCritical, "ATENÇÃO !"
            Exit Sub
        End If
        subInsereRegistroVetor Array(CboDescontosImpostos.Text, FunNulo(DatDescontosImpostos.Recordset.Fields("desc_vl_taxa")), CboDescontosImpostos.BoundText, ""), Vetor1, TDBGrid1
        If Not CmdExcluir.Enabled Then
            CmdExcluir.Enabled = True
            TDBGrid1.Columns(1).Locked = False
        End If
    End If
End Sub


Private Sub CmdInsereBanco_Click()
    Dim Xvalor As Double
        
    XBanco = "I"
    'Preparar Insercao
    PanCadBancoForn.Visible = True
    PanCadBancoForn.Enabled = True
    'PanCadBancoForn.Left = PanBancoForn.Left
    'PanCadBancoForn.Top = PanBancoForn.Top
    CboBanco.Text = ""
    CboBanco.SetFocus
    TxtCodAgencia.Text = ""
    TxtNomeAgencia.Text = ""
    TxtContaCorrente.Text = ""
    TxtDac.Text = ""
End Sub


Private Sub CmdPadrao_Click()
    Dim pos As Integer
    Dim pospadrao As Integer
    If Vetor3(0, 0) <> Empty Then
        pos = TDBGrid3.bookmark
        TDBGrid3.MoveFirst
        Do While Not TDBGrid3.EOF
            If TDBGrid3.Columns("Padrao").Value = "S" Then
                pospadrao = TDBGrid3.bookmark
                If pospadrao = pos Then
                    MsgBox "Este banco já é Padrão!", vbCritical, "ATENÇÃO"
                    Exit Sub
                Else
                    If MsgBox("Já existe um Banco Padrão. Deseja alterar?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma alteração
                        TDBGrid3.Columns("Padrao").Text = "N"
                        TDBGrid3.Update
                        TDBGrid3.bookmark = pos
                        TDBGrid3.Columns("Padrao").Text = "S"
                        TDBGrid3.Update
                        XBancoPadrao = "S"
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            End If
            TDBGrid3.MoveNext
        Loop
        TDBGrid3.bookmark = pos
        TDBGrid3.Columns("Padrao").Text = "S"
        TDBGrid3.Update
        XBancoPadrao = "S"
    End If
End Sub

Private Sub CmdRemoverBanco_Click()
    Dim XLO_BANCO As Object

    'Houve_Erro = False
    
    DesabilitaBotoes ("B")
   
    TDBGrid3.Enabled = False
   
    If TDBGrid3.Columns("CodBanco") <> "" Then
        SubQOpenRecordset XLO_BANCO, "SELECT dupl_cd_duplicata FROM Duplicatas WHERE Duplicatas.bafo_cd_codigo = " & TDBGrid3.Columns("CodBanco") & "", Estatico
        If Not XLO_BANCO.EOF Then 'Existem Duplicatas cadastradas com este código do banco
            MsgBox "O Banco possui movimentos cadastrados. Não é possível excluí-lo.", vbCritical + vbOKOnly, "ATENÇÃO"
            XLO_BANCO.Close
            Exit Sub
        End If
        XLO_BANCO.Close
    End If

    If TDBGrid3.Columns("Status") <> "E" Then
        If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
            'RemoveRegistroVetor Vetor3, TDBGrid3
            TDBGrid3.Columns("Status") = "E"
        End If
    Else
        If MsgBox("Cancela Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
            TDBGrid3.Columns("Status").Text = ""
        End If
    End If
        
    TDBGrid3.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
        HabilitaBotoes ("B")
    Else
        CmdInsereBanco.Enabled = True
    End If
 
    Exit Sub
    
RotuloErro:
    'Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next

End Sub

Private Sub Command1_Click()
    Dim ResFormulario As Object
    Dim resformulario2 As Object
    Dim resformulario3 As Object
    Dim Sql As String
    
    Sql = "SELECT focl_cd_FornCli FROM FornClientes"
    SubQOpenRecordset ResFormulario, Sql, Dinamico
    
    SubQOpenRecordset resformulario3, "SELECT empr_cd_empresa FROM Empresas", Dinamico
    While Not ResFormulario.EOF
        resformulario3.MoveFirst
        While Not resformulario3.EOF
            Sql = "SELECT focl_cd_FornCli, empr_cd_Empresa" & _
                " FROM  dbo.PlCCustoFornCli " & _
                " Where (empr_cd_Empresa = " & resformulario3!empr_cd_Empresa & ") and focl_cd_forncli = " & ResFormulario!focl_cd_FornCli & ""
        
            SubQOpenRecordset resformulario2, Sql, Dinamico
            
            If resformulario2.EOF Then
                resformulario2.AddNew
                resformulario2!empr_cd_Empresa = resformulario3!empr_cd_Empresa
                resformulario2!focl_cd_FornCli = ResFormulario!focl_cd_FornCli
        
                resformulario2.Update
                resformulario2.Close
            End If
            resformulario3.MoveNext
        Wend
        
        ResFormulario.MoveNext
    Wend
    ResFormulario.Close 'inserido em 10/10/08 - Patrícia
    resformulario3.Close 'inserido em 10/10/08 - Patrícia
    
    MsgBox "Operação Realizada com Sucesso!", vbCritical + vbOKOnly, "ATENÇÃO"
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFornecedores
    Formulario.Top = 1000
    Formulario.Left = 0
'    If TxtCgc.Enabled Then
'        TxtCgc.SetFocus
'    End If
    If XFormForn = "FrmManContratos" Then
        OptPF.Value = True
    End If
End Sub

Private Sub Form_Load()
    
    Set Formulario = FrmFornecedores

    subConectarControleDadosNV DatBanco, "SELECT banc_cd_codigo, banc_tx_descricao FROM bancos ORDER BY banc_tx_descricao", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatProfissao, "SELECT * FROM Profissoes ORDER BY prof_tx_descricao", Estatico
    If XGT_NomeAplic = "CAPI" Or XGB_PCAPI Then
        subConectarControleDadosNV DatAtividadePrincipal, "SELECT (CASE LEN(cnse_tx_codigo) WHEN 10 THEN cnse_cd_cnaesecao ELSE 0 END) AS cnse_cd_cnaesecao, cnse_tx_codigo, cnse_tx_descricao FROM ConsCAPCnae ORDER BY cnse_tx_descricao", Estatico 'Alterado em 27/06/08 - Patrícia. estava ordenado por cnse_tx_codigo
    End If
    CboTipoPagamento.ListIndex = 0
    MskPcr.Mask = pPCRmascara
    MskPcr2.Mask = pPCRmascara
    
    'INSERT COMBOBOX NATUREZA JURÍDICA, NÃO CRIAMOS TABELA POIS NÃO É NECESSÁRIO
    cboNaturezaJuridica.AddItem "0. Selecione", 0
    cboNaturezaJuridica.AddItem "1. Órgão Público Federal", 1
    cboNaturezaJuridica.AddItem "2. Órgão Público Estadual", 2
    cboNaturezaJuridica.AddItem "3. Órgão Público Municipal", 3
    cboNaturezaJuridica.AddItem "4. Empresa Pública Federal", 4
    cboNaturezaJuridica.AddItem "5. Empresa Pública Estadual", 5
    cboNaturezaJuridica.AddItem "6. Empresa Pública Municipal", 6
    cboNaturezaJuridica.AddItem "7. Empresa Privada", 7
    cboNaturezaJuridica.AddItem "8. Sociedade Cooperativa", 8
    cboNaturezaJuridica.AddItem "9. Empresa Privada - Fab. Máq. Veíc e autopeças", 9
    
    'INSERT COMBOBOX REGIME APURAÇÃO, NÃO CRIAMOS TABELA POIS NÃO É NECESSÁRIO
    cboRegimeApuracao.AddItem "0. Selecione", 0
    cboRegimeApuracao.AddItem "1. Normal", 1
    cboRegimeApuracao.AddItem "2. Microempresa", 2
    cboRegimeApuracao.AddItem "3. Empresa de pequeno porte", 3
    cboRegimeApuracao.AddItem "4. Outros", 4
    cboRegimeApuracao.AddItem "5. Imune do IRPJ-Art. 4º &1º da IN 1.234/12", 5
    cboRegimeApuracao.AddItem "6. Isenta do IRPJ-Art. 4º &1º da IN 1.2345/12", 6
    
    
    EntrouImp = True 'Indica que o usuário entrou na orelha de impostos
    XBancoPadrao = "N" 'Indica se já foi selecionado um banco padrão
    
    'Modifica os Labels da tela de fornecedor e cliente
    If FornClie = "F" Then 'Caso seja fornecedor
        'chkFornecedorCritico.Visible = True
        chkFornecedorCritico.Enabled = False
        'subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp='D' or plco_tx_recdesp='G' or plco_tx_recdesp='N') ORDER BY plco_tx_desccontabil", Estatico
        'subConectarControleDadosNV DatCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp='D' or plco_tx_recdesp='G' or plco_tx_recdesp='N') ORDER BY plco_tx_desccontabil", Estatico
        If pCrAtivoPassivo = "S" Then
            subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (natu_cd_Natureza=4 or natu_cd_Natureza=1 or natu_cd_Natureza=2) ORDER BY plco_tx_desccontabil", Estatico
        Else
            subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and natu_cd_Natureza=4 ORDER BY plco_tx_desccontabil", Estatico
        End If
        
        subConectarControleDadosNV DatCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and natu_cd_Natureza=2 ORDER BY plco_tx_desccontabil", Estatico
        subConectarControleDadosNV DatProvAdiant, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp='D' or plco_tx_recdesp='G' or plco_tx_recdesp='N') ORDER BY plco_tx_desccontabil", Estatico
        LblDespRec.Caption = "Conta Despesa:"
        LblProvAdiant.Caption = "Conta Adiantamento:"
        
        If Pessoa = "F" Then
            LblPis.Visible = True
            LblNit.Visible = True
            TxtNit.Visible = True
            TxtPis.Visible = True
            TxtPis.Mask = Format$("###.###.###-##")
            TxtNit.Mask = Format$("###.###.###-##")
        'Henrique Lima - 15/04/09 - Habilita o check para fornecedores que são pessoas jurídicas
        ElseIf Pessoa = "J" Or Pessoa = "3" Or Pessoa = "4" Then
            chkSuperSimples.Visible = True
            chkSuperSimples.Enabled = True
        End If
        
    ElseIf FornClie = "C" Then 'Cliente
        chkFornecedorCritico.Visible = False
        'subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp='R' or plco_tx_recdesp='G' or plco_tx_recdesp='N') ORDER BY plco_tx_desccontabil", Estatico
        'subConectarControleDadosNV DatCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp='R' or plco_tx_recdesp='G' or plco_tx_recdesp='N') ORDER BY plco_tx_desccontabil", Estatico
    
        If pCrAtivoPassivo = "S" Then
            subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (natu_cd_Natureza=3 or natu_cd_Natureza=1 or natu_cd_Natureza=2)  ORDER BY plco_tx_desccontabil", Estatico
        Else
            subConectarControleDadosNV DatDespRec, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and natu_cd_Natureza=3 ORDER BY plco_tx_desccontabil", Estatico
        End If
        subConectarControleDadosNV DatCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and natu_cd_Natureza=1 ORDER BY plco_tx_desccontabil", Estatico
        LblDespRec.Caption = "Conta Receita:"
        LblProvAdiant.Caption = "Conta Provisionamento:"
    
        If Pessoa = "J" Then 'Pessoa Jurídica
            MskAtividadePrincipal.Visible = True
            CboAtividadePrincipal.Visible = True
            lblAtividadePrincipal.Visible = True
            CmdAtividadePrincipal.Visible = True
        End If
    End If
    
    
    If Pessoa = "F" Then 'Pessoa Física
        LblCGC.Caption = "CPF:"
        TxtCgc.Width = 1600
        FraCGCCPF.Width = 3600
        LblNomeFantasia.Caption = "Nome:"
        LblRazaoSocial.Caption = "Nome Completo:"
    Else 'Pessoa Jurídica
        LblCGC.Caption = "CNPJ:"
        LblRazaoSocial.Caption = "Razão Social:"
        LblFoneResid.Caption = "Fone Comercial:"
    
    End If
    
    If XInserir = "I" Then
        CodigoDocumento = Chave
        PrepararInsercao
        DesabHabCampos False
        ChecaCGCCPF = True
        'If Pessoa = "F" Or FornClie = "C" Then
            HabilitaCampos FornClie, Pessoa
        'End If
    ElseIf XInserir = "A" Or XInserir = "C" Then
        CodigoDocumento = Chave
        FraEndereco.Enabled = True
        FraContato.Enabled = True
        
        HabilitaCampos FornClie, Pessoa
        PreencheCampos
        TxtCgc.Enabled = False
        LblCGC.Enabled = False
        If XInserir = "C" Then
            PanDadosGerais.Enabled = False
            PanDadosFinanceiros.Enabled = False
            FraDescontos.Enabled = False
            PanDadosPessoais.Enabled = True
            CmdGravar.Visible = False
        End If
    End If
    
'    Dim XLO_RsRamosAtividades As New ADODB.Recordset
'    Dim XLO_RsRamosAtividadesFornecedor As New ADODB.Recordset
'    Dim XLI_I As Integer
'
'    SubQOpenRecordset XLO_RsRamosAtividades, "SELECT * FROM RamosAtividades ORDER BY raat_tx_descricao", "Estatico"
'    SubQOpenRecordset XLO_RsRamosAtividadesFornecedor, "SELECT * FROM FornCliente_RamosAtividades WHERE focl_cd_FornCli = " & Chave & "", "Estatico"
'
'    'Preenche o listbox com os ramos de atividades
'    If Not XLO_RsRamosAtividades.EOF Then
'
'        XLO_RsRamosAtividades.MoveFirst
'        While Not XLO_RsRamosAtividades.EOF
'
'          lsbRamosAtividades.AddItem (XLO_RsRamosAtividades.Fields("raat_cd_ramoatividade") & " - " & XLO_RsRamosAtividades.Fields("raat_tx_descricao"))
'          XLO_RsRamosAtividades.MoveNext
'        Wend
'    End If
'
'    If XInserir = "A" And Not XLO_RsRamosAtividadesFornecedor.EOF Then
'
'        If Not XLO_RsRamosAtividadesFornecedor.EOF Then
'
'            XLO_RsRamosAtividadesFornecedor.MoveFirst
'
'            While Not XLO_RsRamosAtividadesFornecedor.EOF
'
'                For XLI_I = 0 To lsbRamosAtividades.ListCount - 1
'
'                    'Seleciona os ramos de atividades do fornecedor
'                    If XLO_RsRamosAtividadesFornecedor!raat_cd_ramoatividade = CInt(RTrim(LTrim(Split(lsbRamosAtividades.List(XLI_I), "-")(0)))) Then
'                         lsbRamosAtividades.Selected(XLI_I) = True
'                    End If
'
'                Next XLI_I
'
'              XLO_RsRamosAtividadesFornecedor.MoveNext
'
'            Wend
'
'        End If
'
'    End If
'
'    XLO_RsRamosAtividades.Close
'    XLO_RsRamosAtividadesFornecedor.Close
    
End Sub



Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub lsbRamosAtividades_Click()

End Sub

Private Sub MskAtividadePrincipal_GotFocus()
    Call subSelecionaMSK(MskAtividadePrincipal)
End Sub

Private Sub MskAtividadePrincipal_LostFocus()
    Dim XLI_CODCNAE As Integer
    Dim ResAtividade As ADODB.Recordset ' Declaração em 27/06/08 - Patrícia
    
    'Criação de linhas de código em 27/06/08 - Patrícia
    If MskAtividadePrincipal.Text <> "" Then
        SubQOpenRecordset ResAtividade, "select cnse_cd_cnaesecao  from ConsCAPCnae where  SUBSTRING(cnse_tx_codigo, 2,len(cnse_tx_codigo)-1)= '" & StrConv(MskAtividadePrincipal.FormattedText, vbUpperCase) & "'", Dinamico 'Inserido em 27/06/08 - Patrícia
        If Not ResAtividade.EOF Then
            XLI_CODCNAE = ResAtividade!cnse_cd_cnaesecao
            CboAtividadePrincipal.BoundText = XLI_CODCNAE
            ResAtividade.Close
            Set ResAtividade = Nothing
        Else
            MsgBox "Não existe item do CNAE com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            ResAtividade.Close
            Set ResAtividade = Nothing
            Exit Sub
        End If
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

Private Sub CboPcr2_Change()
    If CboPcr2.Text <> "" Then
        DatPcr.Recordset.bookmark = CboPcr2.SelectedItem
        If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr2.BoundText = MskPcr2.Text
            CboPcr2.SetFocus
            Exit Sub
        End If
        MskPcr2.Text = CboPcr2.BoundText
    End If
End Sub

Private Sub MskCep_KeyPress(KeyAscii As Integer)
     If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 13
    End If
End Sub

Private Sub MskCusto_GotFocus()
    Call subSelecionaMSK(MskCusto)

End Sub

Private Sub MskCusto2_LostFocus()
     If MskCusto2.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskCusto2.Text = Format(MskCusto2.Text, XGT_MSKCODREDUZIDO)
        End If
        CboCusto2.BoundText = MskCusto2.Text
        If CboCusto2.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskCusto2.Visible = True Then MskCusto2.SetFocus
        End If
    End If
End Sub

Private Sub MskDespRec_GotFocus()
    Call subSelecionaMSK(MskDespRec)

End Sub

Private Sub MskProvAdiant_GotFocus()
    Call subSelecionaMSK(MskProvAdiant)

End Sub

Private Sub MskProvAdiant_LostFocus()
    If MskProvAdiant.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskProvAdiant.Text = Format(MskProvAdiant.Text, XGT_MSKCODREDUZIDO)
        End If
        CboProvAdiant.BoundText = MskProvAdiant.Text
        If CboProvAdiant.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskProvAdiant.Visible = True Then MskProvAdiant.SetFocus
        End If
    End If
End Sub

Private Sub MskCusto_LostFocus()
     If MskCusto.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskCusto.Text = Format(MskCusto.Text, XGT_MSKCODREDUZIDO)
        End If
        CboCusto.BoundText = MskCusto.Text
        If CboCusto.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskCusto.Visible = True Then MskCusto.SetFocus
        End If
    End If
    
End Sub

Private Sub MskDespRec_LostFocus()
     If MskDespRec.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            'MskDespRec.Text = Format(MskDespRec.Text, XGT_MSKCODREDUZIDO)
        End If
        CboDespRec.BoundText = MskDespRec.Text
        If CboDespRec.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskDespRec.Visible = True Then MskDespRec.SetFocus
        End If
    End If
End Sub

Private Sub MskDespRec2_LostFocus()
     If MskDespRec2.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskDespRec2.Text = Format(MskDespRec2.Text, XGT_MSKCODREDUZIDO)
        End If
        CboDesprec2.BoundText = MskDespRec2.Text
        If CboDesprec2.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskDespRec2.Visible = True Then MskDespRec2.SetFocus
        End If
    End If
End Sub



Private Sub OptPF_Click()

    If (OptPF.Value = True) Then
        Pessoa = "F"
        If FornClie = "F" Then
            TabFornecedores.TabEnabled(2) = True
            PanIcms.Visible = False
        End If
        
        LblCGC.Caption = "CPF:"
        TxtCgc.Width = 1600
        FraCGCCPF.Width = 3600
        LblNomeFantasia.Caption = "Nome:"
        LblRazaoSocial.Caption = "Nome Completo:"
        
        XInserir = "I"
        PrepararInsercao
        DesabHabCampos False
        ChecaCGCCPF = True
        HabilitaCampos FornClie, Pessoa
        
        LblCGC.Enabled = True
        TxtCgc.Enabled = True
        
        'TxtCgc.SetFocus
        
    End If
End Sub

Private Sub OptPJProd_Click()
    If (OptPJProd.Value = True) Then
        TabFornecedores.TabEnabled(2) = False
        Pessoa = "3"
        PanIcms.Visible = True
        LblCGC.Caption = "CNPJ:"
        LblRazaoSocial.Caption = "Razão Social:"
        LblFoneResid.Caption = "Fone Comercial:"
        FrmFornecedores.Caption = "Cadastro de Fornecedores (Pessoa Jurídica - Produto)"
        FraContato.Visible = True
        
        XInserir = "I"
        PrepararInsercao
        DesabHabCampos False
        ChecaCGCCPF = True
        HabilitaCampos FornClie, Pessoa
        
        LblCGC.Enabled = True
        TxtCgc.Enabled = True
        'TxtCgc.SetFocus
        
    End If
End Sub

Private Sub OptPJProdServ_Click()
    If OptPJProdServ.Value = True Then
        Pessoa = "4"
        PanIcms.Visible = True
        LblCGC.Caption = "CNPJ:"
        LblRazaoSocial.Caption = "Razão Social:"
        LblFoneResid.Caption = "Fone Comercial:"
        FrmFornecedores.Caption = "Cadastro de Fornecedores (Pessoa Jurídica - Produto e Serviço)"
        FraContato.Visible = True
           
        XInserir = "I"
        PrepararInsercao
        DesabHabCampos False
        ChecaCGCCPF = True
        HabilitaCampos FornClie, Pessoa
        
        LblCGC.Enabled = True
        TxtCgc.Enabled = True
        'TxtCgc.SetFocus
        
    End If
End Sub

Private Sub OptPJServ_Click()
    If OptPJServ.Value = True Then
    
        Pessoa = "J"
        If FornClie = "F" Then
            TabFornecedores.TabEnabled(2) = True
            FraContato.Visible = True
            PanIcms.Visible = False
            FrmFornecedores.Caption = "Cadastro de Fornecedores (Pessoa Jurídica - Serviço)"
        Else
            FrmFornecedores.Caption = "Cadastro de Clientes (Pessoa Jurídica - Serviço)"
        End If
        LblCGC.Caption = "CNPJ:"
        LblRazaoSocial.Caption = "Razão Social:"
        LblFoneResid.Caption = "Fone Comercial:"
        
        XInserir = "I"
        PrepararInsercao
        DesabHabCampos False
        ChecaCGCCPF = True
        HabilitaCampos FornClie, Pessoa
        
        LblCGC.Enabled = True
        TxtCgc.Enabled = True
        'TxtCgc.SetFocus
        
    End If
End Sub

Private Sub OptProdServ_Click()
    Pessoa = "4"
    If FornClie = "F" Then
        TabFornecedores.TabEnabled(2) = True
    End If
    PanIcms.Visible = True
End Sub

Private Sub OptProduto_Click()
    TabFornecedores.TabEnabled(2) = False
    Pessoa = "3"
    PanIcms.Visible = True
End Sub

Private Sub OptServico_Click()
    TabFornecedores.TabEnabled(2) = True
    Pessoa = "J"
    PanIcms.Visible = False
End Sub

Private Sub TabFornecedores_Click(PreviousTab As Integer)
    Dim Sql As String
    
    If TabFornecedores.Tab = 2 Then
      If EntrouImp Then
            Sql = "SELECT DesImpForn.*,DescontosImpostos.desc_tx_descricao " & _
            " FROM DesImpForn LEFT JOIN DescontosImpostos ON DesImpForn.desc_cd_desconto = DescontosImpostos.desc_cd_desconto" & _
            " WHERE DesImpForn.focl_cd_forncli =" & CodigoDocumento
            subCarregaVetor Sql, Array("desc_tx_descricao", "desi_vl_percentual", "desc_cd_desconto"), Vetor1, TDBGrid1
            EntrouImp = False
            If Vetor1(0, 0) = Empty Then
                DesabilitaBotoes ("I")
                CmdInsere.Enabled = True
            End If
      End If
      
      If TabFornecedores.Tab = 4 And XInserir = "I" Then
        OptQualificado.Value = False
        OptNaoQualificado.Value = False
        OptQualificadoRestrincao.Value = False
        OptFornExclusivo.Value = False
      End If
      
      If FornClie = "F" Then
          subConectarControleDadosNV DatDescontosImpostos, "SELECT * FROM DescontosImpostos where desc_tx_classe = 'D' ORDER BY desc_tx_descricao", "Estatico"
      End If
      CboDescontosImpostos.BoundText = ""
    End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TDBGrid1_LostFocus()
    TDBGrid1.Update
End Sub


Private Sub TDBGrid2_LostFocus()
    TDBGrid2.Update
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid3_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid50.StyleDisp)
Private Sub TDBGrid3_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    If TDBGrid3.Columns("padrao").CellText(bookmark) = "S" Then
        CellStyle.ForeColor = &H8000&
    End If
    If TDBGrid3.Columns("Status").CellText(bookmark) = "E" Then 'Excluído
        CellStyle.ForeColor = &HFF&
    End If
End Sub



Private Sub TxtCgc_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 13
    End If
End Sub

Private Sub TxtCgc_LostFocus()
    Dim Sql As String
    
   If TxtCgc.Text = "___.___.___-__" Or TxtCgc.Text = "__.___.___/____-__" Then
        Exit Sub
    End If
    
    If XInserir = "I" Then
        TxtCgc.PromptInclude = False
        If TxtCgc.Text <> "" Then
            'verifica CPF e CGC
            If Pessoa = "F" Then  'Física
                If Not FunConfereCPF(TxtCgc) Then
                    MsgBox "CPF Invalido!", vbInformation + vbOKOnly, "Atenção!"
                    TxtCgc.PromptInclude = True
                    TxtCgc.SetFocus
                    TxtCgc.SelStart = 0
                    TxtCgc.SelLength = 14

                    Exit Sub
                End If
            ElseIf Pessoa = "J" Or Pessoa = "3" Or Pessoa = "4" Then 'Jurídica ou Produtos
                
                    If Not FunConfereCGC(TxtCgc) Then
                        MsgBox "CGC Invalido!", vbInformation + vbOKOnly, "Atenção!"
                        TxtCgc.PromptInclude = True
                        TxtCgc.SetFocus
                        TxtCgc.SelStart = 0
                        TxtCgc.SelLength = 18
                        Exit Sub
                    End If
                
            End If
        End If
        TxtCgc.PromptInclude = True

        'Habilita os Campos
        DesabHabCampos True
        HabilitaCampos FornClie, Pessoa

        'SubQOpenRecordset ResCGCCPF, "SELECT * FROM FornClientes WHERE focl_tx_cgccpf = '" & TxtCgc.Text & "' AND focl_tx_classe = '" & FornClie & "' AND focl_tx_tipo='" & Pessoa & "'", "Estatico"
        SubQOpenRecordset ResCGCCPF, "SELECT * FROM FornClientes WHERE focl_tx_cgccpf = '" & TxtCgc.Text & "' AND focl_tx_classe = '" & FornClie & "'", "Estatico"

        If Not (ResCGCCPF.BOF And ResCGCCPF.EOF) Then
            ResCGCCPF.MoveFirst
            XInserir = "A"
            Chave = ResCGCCPF!focl_cd_FornCli
            CodigoDocumento = Chave
            ResCGCCPF.Close
            Set ResCGCCPF = Nothing

            PreencheCampos
            TxtCgc.Enabled = False
            LblCGC.Enabled = False
        Else
            Sql = "SELECT BancoFornecedor.*, Bancos.banc_cd_codigo, Bancos.banc_tx_descricao FROM BancoFornecedor LEFT JOIN Bancos ON Bancos.banc_cd_codigo = Bancofornecedor.banc_cd_codigo WHERE focl_cd_forncli = " & Chave
            subCarregaVetor Sql, Array("banc_cd_codigo", "banc_tx_descricao", "bafo_cd_Agencia", "bafo_nr_dag", "bafo_nm_Agencia", "bafo_nr_conta", "bafo_tx_padrao", "bafo_cd_codigo", "bafo_tx_padrao", "bafo_nr_dac"), Vetor3, TDBGrid3
        End If
    End If
    If FraTipoForn2.Enabled = False Or FraTipo.Enabled = False Then
        TxtNomeFantasia.SetFocus
    Else
        TxtCgc.PromptInclude = False
        If TxtCgc.Text <> "" Then
            TxtNomeFantasia.SetFocus
        End If
        TxtCgc.PromptInclude = True
    End If
End Sub

Private Sub TxtContatoFinanceiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        TabFornecedores.Tab = 1
    End If
    
End Sub

Private Sub TxtIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtIcms_LostFocus()
    If IsNumeric(TxtIcms.Text) Then
        TxtIcms.Text = Format$(TxtIcms.Text, "##,##0.00")
    Else
        TxtIcms.Text = ""
    End If
End Sub

Private Sub TxtNomeFantasia_LostFocus()
    If TxtRazaoSocial.Text = "" Then TxtRazaoSocial.Text = TxtNomeFantasia.Text
End Sub

Private Sub TxtRenda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtRenda_LostFocus()
    If IsNumeric(TxtRenda.Text) Then
        TxtRenda.Text = Format$(TxtRenda.Text, "##,##0.00")
    Else
        TxtRenda.Text = ""
    End If
End Sub
