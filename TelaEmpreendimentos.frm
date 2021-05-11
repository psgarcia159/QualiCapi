VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaManEmpreendimentos 
   Caption         =   "Tabela de Empreendimentos"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   1260
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanEmpreendimento 
      Height          =   5115
      Left            =   360
      TabIndex        =   52
      Top             =   360
      Visible         =   0   'False
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
      _ExtentY        =   9022
      _StockProps     =   15
      Caption         =   " Cadastro de Empreendimento"
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
         Left            =   6960
         TabIndex        =   56
         Top             =   4440
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
         Left            =   7980
         TabIndex        =   57
         Top             =   4440
         Width           =   975
      End
      Begin VB.Frame FraDados 
         BackColor       =   &H00E0E0E0&
         Height          =   4110
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   8850
         Begin VB.ComboBox cboAtividadeSped 
            Height          =   315
            ItemData        =   "TelaEmpreendimentos.frx":0000
            Left            =   1440
            List            =   "TelaEmpreendimentos.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   3600
            Width           =   6975
         End
         Begin VB.TextBox txtCodigoMunicipioIbge 
            Height          =   315
            Left            =   4680
            TabIndex        =   16
            Top             =   1380
            Width           =   1095
         End
         Begin VB.TextBox txtBloco 
            Height          =   315
            Left            =   8400
            MaxLength       =   1
            TabIndex        =   7
            Top             =   540
            Width           =   345
         End
         Begin VB.TextBox txtPais 
            Height          =   315
            Left            =   6420
            TabIndex        =   19
            Top             =   1380
            Width           =   2340
         End
         Begin VB.TextBox TxtCga 
            Height          =   315
            Left            =   3360
            TabIndex        =   27
            Top             =   2100
            Width           =   1695
         End
         Begin VB.TextBox TxtTotUnidades 
            Height          =   315
            Left            =   6480
            MaxLength       =   3
            TabIndex        =   29
            Top             =   2100
            Width           =   495
         End
         Begin VB.TextBox TxtFax 
            Height          =   315
            Left            =   1470
            MaxLength       =   15
            TabIndex        =   25
            Top             =   2100
            Width           =   1335
         End
         Begin VB.TextBox TxtFone 
            Height          =   315
            Left            =   4950
            MaxLength       =   15
            TabIndex        =   23
            Top             =   1740
            Width           =   2055
         End
         Begin VB.TextBox TxtNome 
            Height          =   315
            Left            =   2775
            MaxLength       =   50
            TabIndex        =   4
            Top             =   180
            Width           =   5985
         End
         Begin VB.TextBox TxtEndereco 
            Height          =   315
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   6
            Top             =   577
            Width           =   6225
         End
         Begin VB.TextBox TxtBairro 
            Height          =   315
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   9
            Top             =   974
            Width           =   2490
         End
         Begin VB.TextBox TxtCidade 
            Height          =   315
            Left            =   4740
            MaxLength       =   20
            TabIndex        =   11
            Top             =   974
            Width           =   2820
         End
         Begin VB.TextBox TxtResponsavel 
            Height          =   315
            Left            =   1455
            MaxLength       =   30
            TabIndex        =   21
            Top             =   1740
            Width           =   2745
         End
         Begin MSMask.MaskEdBox MskCep 
            Height          =   315
            Left            =   1470
            TabIndex        =   15
            Top             =   1380
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##.###-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskEstado 
            Height          =   315
            Left            =   8310
            TabIndex        =   13
            Top             =   975
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCodigo 
            Height          =   315
            Left            =   1470
            TabIndex        =   2
            Top             =   180
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskConta 
            Height          =   330
            Left            =   1470
            TabIndex        =   33
            Top             =   2460
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
            Left            =   1470
            TabIndex        =   37
            Top             =   2850
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr 
            Bindings        =   "TelaEmpreendimentos.frx":0039
            Height          =   330
            Left            =   2310
            TabIndex        =   38
            Top             =   2850
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
         Begin MSDataListLib.DataCombo CboPlanoContas 
            Bindings        =   "TelaEmpreendimentos.frx":004E
            Height          =   330
            Left            =   2310
            TabIndex        =   34
            Top             =   2460
            Width           =   6075
            _ExtentX        =   10716
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
            Left            =   8415
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2850
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "TelaEmpreendimentos.frx":0065
            Picture         =   "TelaEmpreendimentos.frx":04B7
         End
         Begin Threed.SSCommand CmdLimparplanoContas 
            Height          =   330
            Left            =   8415
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2460
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   -2147483633
            MouseIcon       =   "TelaEmpreendimentos.frx":05C9
            Picture         =   "TelaEmpreendimentos.frx":0A1B
         End
         Begin MSDataListLib.DataCombo CboCCorrente 
            Bindings        =   "TelaEmpreendimentos.frx":0B2D
            Height          =   330
            Left            =   1455
            TabIndex        =   41
            Top             =   3240
            Width           =   6945
            _ExtentX        =   12250
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
            Left            =   8415
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   3240
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            MouseIcon       =   "TelaEmpreendimentos.frx":0B4C
            Picture         =   "TelaEmpreendimentos.frx":0F9E
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Atividade SPED:"
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
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   3600
            Width           =   1125
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   2640
            TabIndex        =   58
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bloco:"
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
            Left            =   7800
            TabIndex        =   17
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblPais 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   5880
            TabIndex        =   18
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblCGA 
            BackColor       =   &H00E0E0E0&
            Caption         =   "CGA:"
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
            Left            =   2880
            TabIndex        =   26
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label LblPcr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   240
            TabIndex        =   36
            Top             =   2910
            Width           =   1200
         End
         Begin VB.Label LblConta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   32
            Top             =   2520
            Width           =   1320
         End
         Begin VB.Label LblUnVendidas2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   8460
            TabIndex        =   31
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LblBairro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   675
            TabIndex        =   8
            Top             =   1005
            Width           =   705
         End
         Begin VB.Label LblEndereco 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   495
            TabIndex        =   5
            Top             =   585
            Width           =   885
         End
         Begin VB.Label LblNome 
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
            Height          =   210
            Left            =   2025
            TabIndex        =   3
            Top             =   225
            Width           =   720
         End
         Begin VB.Label LblCidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   3975
            TabIndex        =   10
            Top             =   1005
            Width           =   735
         End
         Begin VB.Label LblEstado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   7620
            TabIndex        =   12
            Top             =   1005
            Width           =   660
         End
         Begin VB.Label LblCep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   915
            TabIndex        =   14
            Top             =   1425
            Width           =   465
         End
         Begin VB.Label LblCodigo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Código:"
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
            Height          =   225
            Left            =   690
            TabIndex        =   1
            Top             =   210
            Width           =   690
         End
         Begin VB.Label LblUnVendidas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Un. Vendidas:"
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
            TabIndex        =   30
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label LblTotUnidades 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Total Unidades:"
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
            Left            =   5130
            TabIndex        =   28
            Top             =   2160
            Width           =   1365
         End
         Begin VB.Label LblResponsavel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável:"
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
            Left            =   240
            TabIndex        =   20
            Top             =   1785
            Width           =   1185
         End
         Begin VB.Label LblFax 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   975
            TabIndex        =   24
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label LblFone 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fone:"
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
            Left            =   4260
            TabIndex        =   22
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label LblCCorrente 
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
            Height          =   255
            Left            =   60
            TabIndex        =   40
            Top             =   3300
            Width           =   1365
         End
      End
      Begin MSAdodcLib.Adodc DatPcr 
         Height          =   330
         Left            =   240
         Top             =   4680
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
         Left            =   2400
         Top             =   4680
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
         Left            =   4560
         Top             =   4680
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6360
      Top             =   3285
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaEmpreendimentos.frx":10B0
      Height          =   3690
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "empd_cd_empreendimento"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nome"
      Columns(1).DataField=   "empd_tx_nome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Total Unidades"
      Columns(2).DataField=   "empd_nr_unidades"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Unidades Vendidas"
      Columns(3).DataField=   "empd_nr_unvendidas"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Unidades Disponíveis"
      Columns(4).DataField=   "UnDisponiveis"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Responsável"
      Columns(5).DataField=   "empd_tx_responsavel"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Telefone"
      Columns(6).DataField=   "empd_nr_fone"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fax"
      Columns(7).DataField=   "empd_nr_fax"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Endereço"
      Columns(8).DataField=   "empd_tx_endereco"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Bairro"
      Columns(9).DataField=   "empd_tx_bairro"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Cidade"
      Columns(10).DataField=   "empd_tx_cidade"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Estado"
      Columns(11).DataField=   "empd_tx_estado"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "CEP"
      Columns(12).DataField=   "empd_nr_cep"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=6191"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=6112"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(1)._MinWidth=4"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1535"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=530"
      Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=1640"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1561"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=1958"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1879"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=530"
      Splits(0)._ColumnProps(36)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(38)=   "Column(5).Width=3122"
      Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=3043"
      Splits(0)._ColumnProps(41)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=528"
      Splits(0)._ColumnProps(43)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(44)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(45)=   "Column(6).Width=2328"
      Splits(0)._ColumnProps(46)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(6)._WidthInPix=2249"
      Splits(0)._ColumnProps(48)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(6)._ColStyle=528"
      Splits(0)._ColumnProps(50)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(51)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(52)=   "Column(7).Width=2487"
      Splits(0)._ColumnProps(53)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(7)._WidthInPix=2408"
      Splits(0)._ColumnProps(55)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._ColStyle=528"
      Splits(0)._ColumnProps(57)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(58)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(59)=   "Column(8).Width=5689"
      Splits(0)._ColumnProps(60)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(8)._WidthInPix=5609"
      Splits(0)._ColumnProps(62)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(8)._ColStyle=528"
      Splits(0)._ColumnProps(64)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(65)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(66)=   "Column(8)._MinWidth=89391764"
      Splits(0)._ColumnProps(67)=   "Column(9).Width=3493"
      Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=3413"
      Splits(0)._ColumnProps(70)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(9)._ColStyle=528"
      Splits(0)._ColumnProps(72)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(73)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(74)=   "Column(10).Width=2884"
      Splits(0)._ColumnProps(75)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(10)._WidthInPix=2805"
      Splits(0)._ColumnProps(77)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(78)=   "Column(10)._ColStyle=528"
      Splits(0)._ColumnProps(79)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(81)=   "Column(11).Width=1244"
      Splits(0)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(11)._WidthInPix=1164"
      Splits(0)._ColumnProps(84)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=528"
      Splits(0)._ColumnProps(86)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(87)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(88)=   "Column(12).Width=1773"
      Splits(0)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(12)._WidthInPix=1693"
      Splits(0)._ColumnProps(91)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(92)=   "Column(12)._ColStyle=528"
      Splits(0)._ColumnProps(93)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(94)=   "Column(12).Order=13"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   2
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7,.wraptext=-1"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=74,.parent=13,.alignment=0"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=17"
      _StyleDefs(90)  =   "Named:id=33:Normal"
      _StyleDefs(91)  =   ":id=33,.parent=0"
      _StyleDefs(92)  =   "Named:id=34:Heading"
      _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   ":id=34,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=35:Footing"
      _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=36:Selected"
      _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=37:Caption"
      _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(101) =   "Named:id=38:HighlightRow"
      _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=39:EvenRow"
      _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(105) =   "Named:id=40:OddRow"
      _StyleDefs(106) =   ":id=40,.parent=33"
      _StyleDefs(107) =   "Named:id=41:RecordSelector"
      _StyleDefs(108) =   ":id=41,.parent=34"
      _StyleDefs(109) =   "Named:id=42:FilterBar"
      _StyleDefs(110) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   240
      TabIndex        =   51
      Top             =   4125
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
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
      BevelOuter      =   0
      Begin VB.CommandButton CmdImoveis 
         BackColor       =   &H00000000&
         Caption         =   "&Imóveis"
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
         Left            =   3300
         TabIndex        =   46
         Top             =   0
         Width           =   975
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
         Left            =   75
         TabIndex        =   43
         Top             =   0
         Width           =   975
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
         Left            =   7680
         TabIndex        =   50
         Top             =   0
         Width           =   975
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
         Left            =   6540
         TabIndex        =   49
         Top             =   0
         Width           =   975
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
         Left            =   5460
         TabIndex        =   48
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdRemoverFiltro 
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
         Left            =   4380
         TabIndex        =   47
         Top             =   0
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
         Left            =   2220
         TabIndex        =   45
         Top             =   0
         Width           =   975
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
         Left            =   1140
         TabIndex        =   44
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaEmpreendimentos.frx":10C5
      Height          =   315
      Left            =   2760
      TabIndex        =   54
      Top             =   3750
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empd_cd_empreendimento"
      BoundColumn     =   "empd_cd_empreendimento"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Código:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   55
      Top             =   3780
      Width           =   2670
   End
End
Attribute VB_Name = "TelaManEmpreendimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

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

Function funChecarCampos() As Boolean
  
  funChecarCampos = False
  
  If MskCodigo.Text = "" Then
    MsgBox "O preenchimento do Código do empreendimento é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Function
  End If
    
  If TxtNome.Text = "" Then
    MsgBox "O preenchimento do Nome do empreendimento é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
    Exit Function
  End If
  
  If cboAtividadeSped.ListIndex = -1 Then
    MsgBox "Atividade SPED é campo obrigatório", vbCritical + vbOKOnly, "ATENÇÃO"
    cboAtividadeSped.SetFocus
    Exit Function
  End If
  
  funChecarCampos = True
  
End Function


Private Sub subHabilitaBotoes()
    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True
    CmdImoveis.Enabled = True

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       CmdRemoverFiltro.Enabled = False
    Else
       CmdRemoverFiltro.Enabled = True
    End If
End Sub

Private Sub subDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
    CmdImoveis.Enabled = False
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

Private Sub PreencheInformacoesOriginaisLog()
    
    ReDim XGM_MATRIZLOG(18, 3) As Variant
                
    XGM_MATRIZLOG(0, 1) = MskCodigo.Text
    XGM_MATRIZLOG(1, 1) = TxtNome.Text
    XGM_MATRIZLOG(2, 1) = FunNuloVal(TxtTotUnidades.Text)
    XGM_MATRIZLOG(3, 1) = TxtEndereco.Text
    XGM_MATRIZLOG(4, 1) = TxtBairro.Text
    XGM_MATRIZLOG(5, 1) = TxtCidade.Text
    XGM_MATRIZLOG(6, 1) = MskEstado.Text
    XGM_MATRIZLOG(7, 1) = txtBloco.Text
    XGM_MATRIZLOG(8, 1) = MskCep.Text
    XGM_MATRIZLOG(9, 1) = Replace(txtPais.Text, "'", "''")
    XGM_MATRIZLOG(10, 1) = TxtResponsavel.Text
    XGM_MATRIZLOG(11, 1) = TxtFone.Text
    XGM_MATRIZLOG(12, 1) = TxtFax.Text
    XGM_MATRIZLOG(13, 1) = TxtCga.Text
    XGM_MATRIZLOG(14, 1) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
    XGM_MATRIZLOG(15, 1) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
    XGM_MATRIZLOG(16, 1) = FunNulo(CboCCorrente.BoundText)
    XGM_MATRIZLOG(17, 1) = txtCodigoMunicipioIbge.Text
        
End Sub

Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
    Dim XLI_INDICE As Integer
            
    If XLT_OPERACAO = EXCLUIR Then
        
        ReDim XGM_MATRIZLOG(2, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Código"
        XGM_MATRIZLOG(1, 0) = "Nome"
                        
        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("empd_cd_empreendimento")
        XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("empd_tx_Nome")
                
    Else
    
        If XLT_OPERACAO <> ALTERAR Then
            ReDim XGM_MATRIZLOG(18, 2) As Variant
            XLI_INDICE = 1
        Else
            XLI_INDICE = 2
        End If
    
        'Campos
        XGM_MATRIZLOG(0, 0) = "Código"
        XGM_MATRIZLOG(1, 0) = "Nome"
        XGM_MATRIZLOG(2, 0) = "Total Unidades"
        XGM_MATRIZLOG(3, 0) = "Endereço"
        XGM_MATRIZLOG(4, 0) = "Bairro"
        XGM_MATRIZLOG(5, 0) = "Cidade"
        XGM_MATRIZLOG(6, 0) = "Estado"
        XGM_MATRIZLOG(7, 0) = "Bloco"
        XGM_MATRIZLOG(8, 0) = "Cep"
        XGM_MATRIZLOG(9, 0) = "Pais"
        XGM_MATRIZLOG(10, 0) = "Responsável"
        XGM_MATRIZLOG(11, 0) = "Fone"
        XGM_MATRIZLOG(12, 0) = "Fax"
        XGM_MATRIZLOG(13, 0) = "CGA"
        XGM_MATRIZLOG(14, 0) = "Conta Contábil"
        XGM_MATRIZLOG(15, 0) = "Centro de Custo"
        XGM_MATRIZLOG(16, 0) = "Conta Corrente"
        XGM_MATRIZLOG(17, 0) = "Cód. Município IBGE"
            
        'Dados
        XGM_MATRIZLOG(0, XLI_INDICE) = MskCodigo.Text
        XGM_MATRIZLOG(1, XLI_INDICE) = TxtNome.Text
        XGM_MATRIZLOG(2, XLI_INDICE) = FunNuloVal(TxtTotUnidades.Text)
        XGM_MATRIZLOG(3, XLI_INDICE) = TxtEndereco.Text
        XGM_MATRIZLOG(4, XLI_INDICE) = TxtBairro.Text
        XGM_MATRIZLOG(5, XLI_INDICE) = TxtCidade.Text
        XGM_MATRIZLOG(6, XLI_INDICE) = MskEstado.Text
        XGM_MATRIZLOG(7, XLI_INDICE) = txtBloco.Text
        XGM_MATRIZLOG(8, XLI_INDICE) = MskCep.Text
        XGM_MATRIZLOG(9, XLI_INDICE) = Replace(txtPais.Text, "'", "''")
        XGM_MATRIZLOG(10, XLI_INDICE) = TxtResponsavel.Text
        XGM_MATRIZLOG(11, XLI_INDICE) = TxtFone.Text
        XGM_MATRIZLOG(12, XLI_INDICE) = TxtFax.Text
        XGM_MATRIZLOG(13, XLI_INDICE) = TxtCga.Text
        XGM_MATRIZLOG(14, XLI_INDICE) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
        XGM_MATRIZLOG(15, XLI_INDICE) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
        XGM_MATRIZLOG(16, XLI_INDICE) = FunNulo(CboCCorrente.BoundText)
        XGM_MATRIZLOG(17, XLI_INDICE) = txtCodigoMunicipioIbge.Text
    
    End If
    
    Call subRegistraLog("TelaManEmpreendimentos", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, EMPREENDIMENTO_CAPI))
    
End Sub

Private Sub CmdConfirmar_Click()
  Dim XLT_SQL As String
  Dim XLI_HOUVEERRO As Integer
  Dim XLI_BOOKMARK As Integer
  Dim XLT_PCR As String
  Dim XLT_CONTA As String
  
  XLI_HOUVEERRO = False
  
  If funChecarCampos = True Then
  
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
    
    On Error GoTo RotuloErro
    
    If PanEmpreendimento.Tag = "I" Then
        
        XLT_SQL = "INSERT INTO Empreendimentos (empd_cd_empreendimento, empr_cd_empresa, empd_tx_nome, " & _
                        "empd_nr_unidades, empd_nr_unvendidas, empd_tx_endereco, empd_tx_bairro, " & _
                        "empd_tx_Cidade, empd_tx_estado, empd_nr_bloco,empd_nr_Cep, empd_nr_fone, empd_nr_fax, " & _
                        "empd_tx_responsavel,plco_cd_Conta,cere_cd_Pcr,coco_cd_Codigo, empd_nr_cga, empd_tx_pais, empd_cd_MunicipioIbge, empd_tx_atividadesped) " & _
                        "VALUES ('" & MskCodigo.Text & "'," & PCodEmpresa & ",'" & TxtNome.Text & "'," & _
                        "" & FunNuloVal(TxtTotUnidades.Text) & ", 0 ,'" & TxtEndereco.Text & "','" & TxtBairro.Text & "'," & _
                        "'" & TxtCidade.Text & "','" & MskEstado.Text & "','" & Me.txtBloco.Text & "','" & MskCep.Text & "','" & TxtFone.Text & "'," & _
                        "'" & TxtFax.Text & "','" & TxtResponsavel.Text & "'," & FunNuloBanco(XLT_CONTA) & "," & FunNuloBanco(XLT_PCR) & "," & FunNuloBanco(CboCCorrente.BoundText) & ",'" & TxtCga.Text & "'" & ", '" & Replace(txtPais.Text, "'", "''") & "', '" & txtCodigoMunicipioIbge.Text & "','" & IIf(cboAtividadeSped.ListIndex = 0, "I", "A") & "')"
        
        Conexao.Execute XLT_SQL
        
        GravaLog CADASTRAR, "2"

    Else    'Alteração
        
        If CInt(TxtTotUnidades.Text) < Adodc1.Recordset.Fields("empd_nr_unvendidas") Then
          MsgBox "O campo total de unidades tem que ser igual ou maior que o número de Unidades vendidas ( " & Adodc1.Recordset.Fields("empd_nr_unvendidas") & " ) ", vbCritical + vbOKOnly, "ATENÇÃO"
          TxtTotUnidades.Text = Adodc1.Recordset.Fields("empd_nr_unvendidas")
          Exit Sub
        End If
        
        
        XLT_SQL = "UPDATE Empreendimentos SET empr_cd_empresa=" & PCodEmpresa & "," & _
                    "empd_tx_nome='" & TxtNome.Text & "'," & _
                    "empd_nr_unidades=" & TxtTotUnidades.Text & "," & _
                    "empd_tx_endereco='" & TxtEndereco.Text & "'," & _
                    "empd_cd_MunicipioIbge='" & txtCodigoMunicipioIbge.Text & "'," & _
                    "empd_tx_bairro='" & TxtBairro.Text & "'," & _
                    "empd_tx_Cidade='" & TxtCidade.Text & "'," & _
                    "empd_tx_estado='" & MskEstado.Text & "'," & _
                    "empd_nr_bloco='" & txtBloco.Text & "'," & _
                    "empd_nr_Cep='" & MskCep.Text & "'," & _
                    "empd_nr_fone='" & TxtFone.Text & "'," & _
                    "empd_nr_fax='" & TxtFax.Text & "'," & _
                    "empd_tx_responsavel='" & TxtResponsavel.Text & "'," & _
                    "plco_cd_Conta=" & FunNuloBanco(XLT_CONTA) & "," & _
                    "cere_cd_Pcr=" & FunNuloBanco(XLT_PCR) & "," & _
                    "empd_nr_cga='" & TxtCga.Text & "'," & _
                    "coco_cd_Codigo= " & FunNuloBanco(CboCCorrente.BoundText) & ", " & _
                    "empd_tx_pais = '" & Replace(txtPais.Text, "'", "''") & "', " & _
                    "empd_tx_atividadesped = '" & IIf(cboAtividadeSped.ListIndex = 0, "I", "A") & "' " & _
                    " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento = '" & PanEmpreendimento.Tag & "'"
        Conexao.Execute XLT_SQL
        
        GravaLog ALTERAR, "3"

    End If
    
    PanEmpreendimento.Visible = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    If PanEmpreendimento.Tag <> "I" Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        cboLocalizar.BoundText = PanEmpreendimento.Tag
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    Else
        FRM_Quantid = funOperacaoQuantid("+")
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        cboLocalizar.BoundText = MskCodigo.Text
        If Not IsNull(cboLocalizar.SelectedItem) Then
          Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
        End If
    End If
    
    cboLocalizar.Enabled = True
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
  End If
  
  txtPais.Text = ""
  cboAtividadeSped.ListIndex = -1
  Exit Sub
    
RotuloErro:
    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Resume Next
    Exit Sub
   
End Sub

Private Sub CmdImoveis_Click()
  
  If Not FunVerifica_Permissao("TelaManImoveis", "1") Then
        Exit Sub
  End If
  
  XOrigem = "TelaManEmpreendimentos"
  Chave = Adodc1.Recordset("empd_cd_Empreendimento")
  TelaManImoveis.Show
  If FunLocalizaJanelasAtivas("TelaManImoveis") <> -1 Then
    TelaManImoveis.SetFocus
  End If

End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaManEmpreendimentos
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    
    If Filtrou = True Then
        XLI_FIL = ContFil
        subContaReg Adodc1
        Adodc1.Refresh
        subrecarregadados Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = XLI_FIL
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

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    
    If Not FunVerifica_Permissao(Me.Name, "3") Then
        Exit Sub
    End If
    
    PanBotoes.Enabled = False
    cboLocalizar.Enabled = False
    PanEmpreendimento.Visible = True
    PanEmpreendimento.Enabled = True
    PanEmpreendimento.Left = (TDBGrid1.Width - PanEmpreendimento.Width) / 2
    PanEmpreendimento.Top = (TDBGrid1.Height - PanEmpreendimento.Height) / 2
    PanEmpreendimento.Tag = Adodc1.Recordset.Fields("empd_cd_empreendimento")
    PanEmpreendimento.Caption = " Alteração de Empreendimento"
    subDesabilitaBotoes
    subPreencheCampos
    TDBGrid1.Enabled = False
    LblUnVendidas.Visible = True
    SubDesabHabCampos (True)
    TxtNome.SetFocus
    
    'Log
    PreencheInformacoesOriginaisLog
        
End Sub

Private Sub CmdAlterar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
End Sub

Sub subPreencheCampos()
    MskCodigo.Text = Adodc1.Recordset.Fields("empd_cd_empreendimento")
    MskCodigo.Enabled = False
    LblCodigo.Enabled = False
    TxtNome.Text = Adodc1.Recordset.Fields("empd_tx_nome")
    TxtEndereco.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_endereco"))
    txtCodigoMunicipioIbge.Text = FunNulo(Adodc1.Recordset.Fields("empd_cd_MunicipioIbge"))
    TxtBairro.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_bairro"))
    TxtCidade.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_Cidade"))
    If Not IsNull(Adodc1.Recordset.Fields("empd_tx_pais")) Then
        txtPais.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_pais"))
    End If
    MskEstado.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_estado"))
    If IsNull(MskCep.Text = (Adodc1.Recordset.Fields("empd_nr_Cep"))) Then
      MskCep.Text = "  .   -   "
    Else
      MskCep.Text = Adodc1.Recordset.Fields("empd_nr_Cep")
    End If
    TxtResponsavel.Text = FunNulo(Adodc1.Recordset.Fields("empd_tx_responsavel"))
    TxtFone.Text = FunNulo(Adodc1.Recordset.Fields("empd_nr_fone"))
    TxtFax.Text = FunNulo(Adodc1.Recordset.Fields("empd_nr_fax"))
    TxtTotUnidades = FunNulo(Adodc1.Recordset.Fields("empd_nr_unidades"))
    LblUnVendidas2.Caption = FunNulo(Adodc1.Recordset.Fields("empd_nr_unvendidas"))
    TxtCga.Text = FunNulo(Adodc1.Recordset.Fields("empd_nr_cga"))
    txtBloco.Text = FunNulo(Adodc1.Recordset.Fields("empd_nr_bloco"))
  
    If Not DatConta.Recordset.EOF Then
      DatConta.Recordset.MoveFirst
      If Not IsNull(Adodc1.Recordset.Fields!plco_cd_Conta) Then
        DatConta.Recordset.Find "plco_cd_Conta =" & Adodc1.Recordset.Fields!plco_cd_Conta
        MskConta.Text = DatConta.Recordset.Fields("plco_cd_Codreduzido")
        CboPlanoContas.BoundText = MskConta.Text
      Else
        MskConta.PromptInclude = False
        MskConta.Text = ""
        MskConta.PromptInclude = True
        CboPlanoContas.BoundText = ""
      End If
    End If
      
    If Not DatPcr.Recordset.EOF Then
      DatPcr.Recordset.MoveFirst
      If Not IsNull(Adodc1.Recordset.Fields!cere_cd_Pcr) Then
          DatPcr.Recordset.Find "cere_cd_pcr =" & Adodc1.Recordset.Fields!cere_cd_Pcr
          MskPcr.Text = DatPcr.Recordset.Fields("cere_cd_estruturado")
          CboPcr.BoundText = MskPcr.Text
      Else
        MskPcr.PromptInclude = False
        MskPcr.Text = ""
        MskPcr.PromptInclude = True
        CboPcr.BoundText = ""
      End If
    End If
    
    If Not DatContaCorrente.Recordset.EOF Then
      DatContaCorrente.Recordset.MoveFirst
      If Not IsNull(Adodc1.Recordset.Fields!coco_cd_codigo) Then
          DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & Adodc1.Recordset.Fields!coco_cd_codigo
          CboCCorrente.BoundText = FunNulo(Adodc1.Recordset.Fields("coco_cd_Codigo"))
      Else
        CboCCorrente.BoundText = ""
      End If
    End If
    
    If (Adodc1.Recordset.Fields!empd_tx_atividadesped = "I") Then
        cboAtividadeSped.ListIndex = 0
    Else
        cboAtividadeSped.ListIndex = 1
    End If
    

End Sub

Private Sub CmdRetornar_Click()
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanEmpreendimento.Visible = False
    PanEmpreendimento.Enabled = False
    subHabilitaBotoes
    cboLocalizar.Enabled = True
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    txtPais.Text = ""
    cboAtividadeSped.ListIndex = -1
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Empreendimentos")
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaManEmpreendimentos
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()

    If Not FunVerifica_Permissao(Me.Name, "4") Then
        Exit Sub
    End If
    
    Dim XLI_OPCAO As Integer
    Dim XLI_BOOKMARK As Integer
        
    On Error GoTo RotuloErro
    
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
      Conexao.BeginTrans
      Conexao.Execute "DELETE FROM Empreendimentos WHERE empd_cd_empreendimento='" & Adodc1.Recordset.Fields("empd_cd_empreendimento") & "' AND empr_cd_Empresa =" & PCodEmpresa & ""
      Conexao.CommitTrans
      
      GravaLog EXCLUIR, "4"
      
      FRM_Quantid = funOperacaoQuantid("-")
      XLI_BOOKMARK = TDBGrid1.bookmark
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      
      If Not Adodc1.Recordset.EOF Then
         If XLI_BOOKMARK <= Adodc1.Recordset.RecordCount Then
           TDBGrid1.bookmark = XLI_BOOKMARK
         Else
           TDBGrid1.bookmark = XLI_BOOKMARK - 1
         End If
      End If
       
      subHabilitaBotoes
      
      If FunTabelaVazia(Adodc1) = True Then
         CmdAlterar.Enabled = False
         CmdExcluir.Enabled = False
         CmdImprimir.Enabled = False
         CmdOrdem.Enabled = False
         CmdImoveis.Enabled = False
      End If
       
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Esse empreendimento não pode ser excluído pois já possui imóveis relacionados!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
    
End Sub

Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
End Sub

Private Sub CmdInserir_Click()
    
    If Not FunVerifica_Permissao(Me.Name, "2") Then
        Exit Sub
    End If
    
    PanBotoes.Enabled = False
    PanEmpreendimento.Left = (TDBGrid1.Width - PanEmpreendimento.Width) / 2
    PanEmpreendimento.Top = (TDBGrid1.Height - PanEmpreendimento.Height) / 2
    PanEmpreendimento.Caption = " Inserção de Empreendimento"
    PanEmpreendimento.Visible = True
    PanEmpreendimento.Enabled = True
    subPrepararInsercao
    PanEmpreendimento.Tag = "I"
    subDesabilitaBotoes
    cboLocalizar.Text = ""
    cboLocalizar.Enabled = False
    TDBGrid1.Enabled = False
    LblUnVendidas.Visible = False
    MskCodigo.Enabled = True
    If BarraFerramentasPressionada = False Then
      MskCodigo.SetFocus
    End If
    SubDesabHabCampos (False)
    
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
End Sub

Sub subPrepararInsercao()
    LblCodigo.Enabled = True
    MskCodigo.Enabled = True
    MskCodigo.Text = ""
    TxtNome.Text = ""
    TxtEndereco.Text = ""
    txtCodigoMunicipioIbge.Text = ""
    TxtBairro.Text = ""
    TxtCidade.Text = ""
    TxtResponsavel.Text = ""
    txtPais.Text = ""
    MskEstado.Mask = ""
    MskEstado.Text = ""
    MskCep.Mask = ""
    MskCep.Text = ""
    MskCep.Mask = "##.###-###"
    TxtFone.Text = ""
    TxtFax.Text = ""
    TxtTotUnidades.Text = ""
    LblUnVendidas2.Caption = ""
        
    CboCCorrente.BoundText = ""
    CboPcr.BoundText = ""
    CboPlanoContas.BoundText = ""
        
End Sub
  
Private Sub Adodc1_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case funTrataErros(ComMensagem)
        Case ErrChaveDuplicada
            CancelDisplay = False
        Case ErrRegistroExcluido
            CancelDisplay = False
            Adodc1.Refresh
            CmdInserir.Caption = "&Inserir"
            CmdExcluir.Caption = "&Excluir"
            subHabilitaBotoes
        Case ErrExclusaoRelacionamento
            CancelDisplay = False
    End Select
End Sub

Private Sub MskCodigo_GotFocus()
   
   Call subSelecionaMSK(MskCodigo)
   'SubDesabHabCampos (True)
   
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaManEmpreendimentos
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImoveis.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaManEmpreendimentos
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaManEmpreendimentos
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaManEmpreendimentos"
         
    'Associa um banco de dados à tabela profissões
    subCarregaDadosNV Adodc1, "empreendimentos", "empd_cd_empreendimento", _
                            "", "SELECT * FROM ConsCAPEmpreendimentos WHERE empr_cd_Empresa= " & PCodEmpresa & ""
    
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
    subHabilitaBotoes
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImoveis.Enabled = False
    End If
    
    If BarraFerramentasPressionada = True Then
      Call CmdInserir_Click
      BarraFerramentasPressionada = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaManEmpreendimentos"
End Sub


Sub SubDesabHabCampos(XHab As Boolean)
    'Desabilita/Habilita os campos em caso de inserção
    TxtNome.Enabled = XHab
    TxtEndereco.Enabled = XHab
    txtCodigoMunicipioIbge.Enabled = XHab
    TxtBairro.Enabled = XHab
    TxtCidade.Enabled = XHab
    TxtResponsavel.Enabled = XHab
    TxtTotUnidades.Enabled = XHab
    TxtFax.Enabled = XHab
    TxtFone.Enabled = XHab
    MskEstado.Enabled = XHab
    MskCep.Enabled = XHab
    CboCCorrente.Enabled = XHab
    CboPcr.Enabled = XHab
    CboPlanoContas.Enabled = XHab
    txtPais.Enabled = XHab
    LblNome.Enabled = XHab
    LblEndereco.Enabled = XHab
    LblBairro.Enabled = XHab
    LblCidade.Enabled = XHab
    LblResponsavel.Enabled = XHab
    LblTotUnidades.Enabled = XHab
    LblUnVendidas.Enabled = XHab
    LblFone.Enabled = XHab
    LblFax.Enabled = XHab
    LblEstado.Enabled = XHab
    LblCep.Enabled = XHab
    LblPcr.Enabled = XHab
    LblCCorrente.Enabled = XHab
    LblConta.Enabled = XHab

End Sub

Private Sub MskCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
    SubDesabHabCampos (True)
End Sub

Private Sub MskCodigo_LostFocus()
    Dim XLO_EMPREENDIMENTO As Object
    Dim XLT_SQL As String
    
    If PanEmpreendimento.Tag = "I" Then
        
        SubPreencheComZero MskCodigo
        SubQOpenRecordset XLO_EMPREENDIMENTO, "SELECT * FROM Empreendimentos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento = '" & MskCodigo.Text & "'", "Estatico"
    
        If Not (XLO_EMPREENDIMENTO.BOF And XLO_EMPREENDIMENTO.EOF) Then
          PanEmpreendimento.Tag = XLO_EMPREENDIMENTO!empd_cd_Empreendimento
          TxtNome.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Nome)
          TxtEndereco.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Endereco)
          txtCodigoMunicipioIbge.Text = FunNulo(XLO_EMPREENDIMENTO!empd_cd_MunicipioIbge)
          TxtBairro.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Bairro)
          TxtCidade.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Cidade)
          MskEstado.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Bairro)
          
          MskCep.PromptInclude = False
          MskCep.Text = FunNulo(XLO_EMPREENDIMENTO!empd_nr_Cep)
          MskCep.PromptInclude = True
          
          TxtResponsavel.Text = FunNulo(XLO_EMPREENDIMENTO!empd_tx_Responsavel)
          TxtFone.Text = FunNulo(XLO_EMPREENDIMENTO!empd_nr_Fone)
          TxtFax.Text = FunNulo(XLO_EMPREENDIMENTO!empd_nr_Fax)
          TxtTotUnidades.Text = FunNuloVal(XLO_EMPREENDIMENTO!empd_nr_Unidades)
          LblUnVendidas.Visible = True
          LblUnVendidas2.Caption = FunNuloVal(XLO_EMPREENDIMENTO!empd_nr_unvendidas)
          
          XLO_EMPREENDIMENTO.Close
          Set XLO_EMPREENDIMENTO = Nothing
          
          cboLocalizar.BoundText = PanEmpreendimento.Tag
          If Not cboLocalizar.Text = "" Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
          subPreencheCampos
          PanEmpreendimento.Caption = " Alteração de Empreendimento"
        End If
        
    End If
End Sub

Private Sub TxtTotUnidades_KeyPress(KeyAscii As Integer)
  
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
     KeyAscii = 0
  End If
  
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

Private Sub CmdLimparplanoContas_Click()
  
  CboPlanoContas.BoundText = ""
  MskConta.Text = ""
  
End Sub

Private Sub CmdLimparPcr_Click()
  
  CboPcr.BoundText = ""
  MskPcr.PromptInclude = False
  MskPcr.Text = ""
  MskPcr.PromptInclude = True
  
End Sub

Private Sub CmdLimparCCorrente_Click()
  CboCCorrente.BoundText = ""
End Sub
