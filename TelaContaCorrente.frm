VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaContaCorrente 
   Caption         =   "Tabela de Contas Correntes"
   ClientHeight    =   5955
   ClientLeft      =   240
   ClientTop       =   1320
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4590
      Top             =   3060
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin Threed.SSPanel PanContaCorrente 
      Height          =   5055
      Left            =   1200
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   9450
      _Version        =   65536
      _ExtentX        =   16669
      _ExtentY        =   8916
      _StockProps     =   15
      Caption         =   "Cadastro de Conta Corrente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Height          =   345
         Left            =   6840
         TabIndex        =   15
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   345
         Left            =   7920
         TabIndex        =   16
         Top             =   4440
         Width           =   975
      End
      Begin TabDlg.SSTab TabContaCorrente 
         Height          =   4470
         Left            =   150
         TabIndex        =   28
         Top             =   435
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   7885
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "TelaContaCorrente.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblDescricao(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblCContabil(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblBanco(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LblCdAgencia(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LblAgencia(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LblConta(4)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LblDag(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LblDac(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "LblConvenioFinanceiro(6)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LblCarteira(7)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "LblConvenioFinanceiro(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "LblConvenioFinanceiro(1)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label2"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "LblCGCRaiz(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "LblCContabil(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "MskContaDespesa"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "cboPlanoContasDespesa"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "MskCGCRaiz"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cboBancos"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "MskBanco"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "MskContaCusto"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "cboPlanoContasCusto"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "TxtDescricao"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "TxtCodAgencia"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "TxtAgencia"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TxtNumeroConta"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TxtDag"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TxtDac"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TxtConvenioFinanceiro"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "TxtCarteira"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtConvenioCobranca"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtGerente"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "OptAtiva"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "OptInativa"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "CboProvisao"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).ControlCount=   36
         TabCaption(1)   =   "Instruções Bancárias"
         TabPicture(1)   =   "TelaContaCorrente.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSPanel1"
         Tab(1).Control(1)=   "PanInstrucao"
         Tab(1).Control(2)=   "FraArquivos"
         Tab(1).ControlCount=   3
         Begin VB.ComboBox CboProvisao 
            Height          =   315
            ItemData        =   "TelaContaCorrente.frx":0038
            Left            =   1545
            List            =   "TelaContaCorrente.frx":004E
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   3555
            Width           =   1935
         End
         Begin VB.OptionButton OptInativa 
            Caption         =   "Inativa"
            Height          =   375
            Left            =   5355
            TabIndex        =   64
            Top             =   3555
            Width           =   1095
         End
         Begin VB.OptionButton OptAtiva 
            Caption         =   "Ativa"
            Height          =   375
            Left            =   4560
            TabIndex        =   63
            Top             =   3555
            Width           =   975
         End
         Begin VB.TextBox txtGerente 
            Height          =   315
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1920
            Width           =   7150
         End
         Begin VB.TextBox txtConvenioCobranca 
            Height          =   315
            Left            =   4560
            MaxLength       =   20
            TabIndex        =   13
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Frame FraArquivos 
            Caption         =   "Número Sequencial dos Arquivos:"
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
            Left            =   -70440
            TabIndex        =   53
            Top             =   360
            Width           =   3855
            Begin VB.TextBox TxtArqForn 
               Height          =   315
               Left            =   3075
               MaxLength       =   6
               TabIndex        =   55
               Top             =   240
               Width           =   705
            End
            Begin VB.TextBox TxtArqCobranca 
               Height          =   315
               Left            =   1200
               MaxLength       =   6
               TabIndex        =   54
               Top             =   240
               Width           =   705
            End
            Begin VB.Label LblArqForn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fornecedores:"
               Height          =   255
               Index           =   0
               Left            =   1560
               TabIndex        =   57
               Top             =   285
               Width           =   1485
            End
            Begin VB.Label LblArqCogranca 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Cobrança:"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   56
               Top             =   285
               Width           =   1125
            End
         End
         Begin Threed.SSPanel PanInstrucao 
            Height          =   525
            Left            =   -74685
            TabIndex        =   45
            Top             =   420
            Width           =   4170
            _Version        =   65536
            _ExtentX        =   7355
            _ExtentY        =   926
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
            Begin VB.TextBox TxtInstrucao2 
               Height          =   315
               Left            =   3345
               MaxLength       =   2
               TabIndex        =   30
               Top             =   105
               Width           =   345
            End
            Begin VB.TextBox TxtInstrucao1 
               Height          =   315
               Left            =   1470
               MaxLength       =   2
               TabIndex        =   29
               Top             =   105
               Width           =   345
            End
            Begin VB.Label LblInstrucao2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Instrução 2:"
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
               Left            =   2160
               TabIndex        =   47
               Top             =   150
               Width           =   1125
            End
            Begin VB.Label LblInstrucao1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Instrução 1:"
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
               Left            =   285
               TabIndex        =   46
               Top             =   150
               Width           =   1125
            End
         End
         Begin VB.TextBox TxtCarteira 
            Height          =   315
            Left            =   7800
            MaxLength       =   5
            TabIndex        =   14
            Top             =   3120
            Width           =   910
         End
         Begin VB.TextBox TxtConvenioFinanceiro 
            Height          =   315
            Left            =   1545
            MaxLength       =   20
            TabIndex        =   12
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox TxtDac 
            Height          =   315
            Left            =   8385
            MaxLength       =   2
            TabIndex        =   5
            Top             =   1090
            Width           =   315
         End
         Begin VB.TextBox TxtDag 
            Height          =   315
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   7
            Top             =   1520
            Width           =   315
         End
         Begin VB.TextBox TxtNumeroConta 
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1090
            Width           =   1380
         End
         Begin VB.TextBox TxtAgencia 
            Height          =   315
            Left            =   4965
            MaxLength       =   30
            TabIndex        =   8
            Top             =   1520
            Width           =   3740
         End
         Begin VB.TextBox TxtCodAgencia 
            Height          =   315
            Left            =   1545
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1520
            Width           =   1575
         End
         Begin VB.TextBox TxtDescricao 
            Height          =   315
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   1
            Top             =   660
            Width           =   7155
         End
         Begin MSDataListLib.DataCombo cboPlanoContasCusto 
            Bindings        =   "TelaContaCorrente.frx":008B
            Height          =   330
            Left            =   2430
            TabIndex        =   11
            Top             =   2670
            Width           =   6285
            _ExtentX        =   11086
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
         Begin MSMask.MaskEdBox MskContaCusto 
            Height          =   330
            Left            =   1545
            TabIndex        =   10
            Top             =   2670
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskBanco 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   1090
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo cboBancos 
            Bindings        =   "TelaContaCorrente.frx":00A7
            Height          =   315
            Left            =   2055
            TabIndex        =   3
            Top             =   1095
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "banc_tx_descricao"
            BoundColumn     =   "banc_cd_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   1755
            Left            =   -74685
            TabIndex        =   48
            Top             =   1005
            Width           =   8145
            _Version        =   65536
            _ExtentX        =   14367
            _ExtentY        =   3096
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
            Begin VB.TextBox TxtMensagem4 
               Height          =   315
               Left            =   1485
               MaxLength       =   55
               TabIndex        =   34
               Top             =   1305
               Width           =   6375
            End
            Begin VB.TextBox TxtMensagem3 
               Height          =   315
               Left            =   1485
               MaxLength       =   55
               TabIndex        =   33
               Top             =   910
               Width           =   6375
            End
            Begin VB.TextBox TxtMensagem1 
               Height          =   315
               Left            =   1485
               MaxLength       =   55
               TabIndex        =   31
               Top             =   120
               Width           =   6375
            End
            Begin VB.TextBox TxtMensagem2 
               Height          =   315
               Left            =   1485
               MaxLength       =   55
               TabIndex        =   32
               Top             =   515
               Width           =   6375
            End
            Begin VB.Label LblMensagem4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensagem 4:"
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
               Index           =   5
               Left            =   300
               TabIndex        =   52
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label LblMensagem3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensagem 3:"
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
               Index           =   4
               Left            =   300
               TabIndex        =   51
               Top             =   955
               Width           =   1140
            End
            Begin VB.Label LblMensagem1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensagem 1:"
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
               Index           =   3
               Left            =   300
               TabIndex        =   50
               Top             =   165
               Width           =   1140
            End
            Begin VB.Label LblMensagem2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensagem 2:"
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
               Index           =   2
               Left            =   300
               TabIndex        =   49
               Top             =   560
               Width           =   1140
            End
         End
         Begin MSMask.MaskEdBox MskCGCRaiz 
            Height          =   315
            Left            =   1545
            TabIndex        =   67
            ToolTipText     =   "CPF do responsável perante a SRF"
            Top             =   3960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##.###.###"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo cboPlanoContasDespesa 
            Bindings        =   "TelaContaCorrente.frx":00BE
            Height          =   330
            Left            =   2430
            TabIndex        =   68
            Top             =   2300
            Width           =   6285
            _ExtentX        =   11086
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
         Begin MSMask.MaskEdBox MskContaDespesa 
            Height          =   330
            Left            =   1545
            TabIndex        =   69
            Top             =   2300
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin VB.Label LblCContabil 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "C/C Despesa:"
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
            Left            =   180
            TabIndex        =   70
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label LblCGCRaiz 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "CGC Raiz:"
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
            Left            =   120
            TabIndex        =   66
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   3735
            TabIndex        =   62
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label LblConvenioFinanceiro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Provisão:"
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
            Left            =   120
            TabIndex        =   61
            Top             =   3570
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Gerente:"
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
            Left            =   750
            TabIndex        =   59
            Top             =   1910
            Width           =   735
         End
         Begin VB.Label LblConvenioFinanceiro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Convênio Cobrança:"
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
            Index           =   0
            Left            =   2880
            TabIndex        =   58
            Top             =   3120
            Width           =   1650
         End
         Begin VB.Label LblCarteira 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Carteira de Cobrança:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   6720
            TabIndex        =   44
            Top             =   3120
            Width           =   1005
         End
         Begin VB.Label LblConvenioFinanceiro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Convênio Fin.:"
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
            Index           =   6
            Left            =   180
            TabIndex        =   43
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label LblDac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "DAC:"
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
            Left            =   7815
            TabIndex        =   42
            Top             =   1125
            Width           =   555
         End
         Begin VB.Label LblDag 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "DAG:"
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
            Index           =   6
            Left            =   3150
            TabIndex        =   41
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label LblConta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Conta:"
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
            Index           =   4
            Left            =   5595
            TabIndex        =   40
            Top             =   1125
            Width           =   870
         End
         Begin VB.Label LblAgencia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência:"
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
            Index           =   3
            Left            =   4110
            TabIndex        =   39
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label LblCdAgencia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Agência:"
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
            Index           =   2
            Left            =   105
            TabIndex        =   38
            Top             =   1510
            Width           =   1410
         End
         Begin VB.Label LblBanco 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   1
            Left            =   525
            TabIndex        =   37
            Top             =   1115
            Width           =   990
         End
         Begin VB.Label LblCContabil 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "C/C Custo:"
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
            Index           =   5
            Left            =   180
            TabIndex        =   36
            Top             =   2670
            Width           =   1335
         End
         Begin VB.Label LblDescricao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição:"
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
            Left            =   180
            TabIndex        =   35
            Top             =   640
            Width           =   1335
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatBanco 
      Height          =   375
      Left            =   2865
      Top             =   3060
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc DatContaCusto 
      Height          =   375
      Left            =   705
      Top             =   3060
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "DatContaCusto"
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
      Bindings        =   "TelaContaCorrente.frx":00DC
      Height          =   3690
      Left            =   75
      TabIndex        =   24
      Top             =   15
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   29
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "coco_cd_Codigo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Banco"
      Columns(1).DataField=   "banc_tx_descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Código Agência"
      Columns(2).DataField=   "coco_cd_agencia"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Agência"
      Columns(3).DataField=   "coco_nm_agencia"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DAG"
      Columns(4).DataField=   "coco_nr_Dag"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Conta Corrente"
      Columns(5).DataField=   "coco_tx_conta"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "DAC"
      Columns(6).DataField=   "coco_nr_Dac"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Descrição"
      Columns(7).DataField=   "coco_tx_descricao"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Conta Contábil"
      Columns(8).DataField=   "plco_tx_DescContabil"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Provisão"
      Columns(9).DataField=   "coco_tx_provisao"
      Columns(9).NumberFormat=   "Yes/No"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Status"
      Columns(10).DataField=   "coco_tx_status"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=4604"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4524"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(22)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2805"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8720"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(32)=   "Column(4).Width=900"
      Splits(0)._ColumnProps(33)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._WidthInPix=820"
      Splits(0)._ColumnProps(35)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=8209"
      Splits(0)._ColumnProps(37)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(39)=   "Column(5).Width=2540"
      Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2461"
      Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=8720"
      Splits(0)._ColumnProps(44)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(45)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(47)=   "Column(6).Width=873"
      Splits(0)._ColumnProps(48)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(6)._WidthInPix=794"
      Splits(0)._ColumnProps(50)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(6)._ColStyle=8209"
      Splits(0)._ColumnProps(52)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(53)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(54)=   "Column(7).Width=7064"
      Splits(0)._ColumnProps(55)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._WidthInPix=6985"
      Splits(0)._ColumnProps(57)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(7)._ColStyle=8720"
      Splits(0)._ColumnProps(59)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(60)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(61)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(62)=   "Column(8).Width=7064"
      Splits(0)._ColumnProps(63)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(8)._WidthInPix=6985"
      Splits(0)._ColumnProps(65)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(8)._ColStyle=8720"
      Splits(0)._ColumnProps(67)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(68)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(69)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(70)=   "Column(9).Width=1588"
      Splits(0)._ColumnProps(71)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(9)._WidthInPix=1508"
      Splits(0)._ColumnProps(73)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(74)=   "Column(9)._ColStyle=20"
      Splits(0)._ColumnProps(75)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(76)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(77)=   "Column(9)._MinWidth=117902480"
      Splits(0)._ColumnProps(78)=   "Column(10).Width=1588"
      Splits(0)._ColumnProps(79)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(10)._WidthInPix=1508"
      Splits(0)._ColumnProps(81)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(82)=   "Column(10)._ColStyle=532"
      Splits(0)._ColumnProps(83)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(84)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(85)=   "Column(10)._MinWidth=117902480"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=86,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=90,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
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
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   120
      TabIndex        =   23
      Top             =   4125
      Width           =   9120
      _Version        =   65536
      _ExtentX        =   16087
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
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   8640
         Picture         =   "TelaContaCorrente.frx":00F1
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   0
         Width           =   375
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
         Left            =   315
         TabIndex        =   0
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
         Left            =   7560
         TabIndex        =   22
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
         Left            =   6360
         TabIndex        =   21
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
         Left            =   5160
         TabIndex        =   20
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
         Left            =   3960
         TabIndex        =   19
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
         Left            =   2760
         TabIndex        =   18
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
         Left            =   1560
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaContaCorrente.frx":02A0
      Height          =   315
      Left            =   2760
      TabIndex        =   25
      Top             =   3750
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "banc_tx_descricao"
      BoundColumn     =   "coco_cd_codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DatContaDespesa 
      Height          =   375
      Left            =   720
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "DatContaDespesa"
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
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Código:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3780
      Width           =   2670
   End
End
Attribute VB_Name = "TelaContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Chave As Integer

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

Private Sub subHabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True

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
End Sub

Private Sub cboBancos_Change()
    MskBanco.Text = cboBancos.BoundText
End Sub

Private Sub cboBancos_Click(Area As Integer)
'    If Area = 2 Then
'        MskBanco.Text = cboBancos.BoundText
'    End If
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If

End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If

End Sub


Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub CboPlanoContasCusto_Change()
    If cboPlanoContasCusto.Text <> "" Then
        DatContaCusto.Recordset.Bookmark = cboPlanoContasCusto.SelectedItem
        If DatContaCusto.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma conta corrente.", vbCritical + vbOKOnly, "ATENÇÃO"
            cboPlanoContasCusto.BoundText = MskContaCusto.Text
            cboPlanoContasCusto.SetFocus
            Exit Sub
        End If
        MskContaCusto.Text = cboPlanoContasCusto.BoundText
    End If
End Sub

Private Sub CboPlanoContasDespesa_Change()
    If cboPlanoContasDespesa.Text <> "" Then
        DatContaDespesa.Recordset.Bookmark = cboPlanoContasDespesa.SelectedItem
        If DatContaDespesa.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma conta corrente.", vbCritical + vbOKOnly, "ATENÇÃO"
            cboPlanoContasDespesa.BoundText = MskContaDespesa.Text
            cboPlanoContasDespesa.SetFocus
            Exit Sub
        End If
        MskContaDespesa.Text = cboPlanoContasDespesa.BoundText
    End If
End Sub

Private Sub CmdConfirmar_Click()
    Dim Sql As String
    Dim Status As String
    
    'On Error GoTo MostraErro
    
    'Função para verificar preenchimento dos campos obrigatórios
        
    If MskBanco.Text = "" Or TxtNumeroConta.Text = "" Or TxtCodAgencia.Text = "" Then
        MsgBox "O preenchimento dos campos Banco, Número da Conta e Agência são obrigatórios.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    If MskContaDespesa.Text = "" Then
        MsgBox "O preenchimento do campo Conta Contábil Despesa é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
        
    'Inserido CGC Raiz no Cadastro da Conta Corrente - Samuel 26/09/18
    If MskCGCRaiz.Text = "" Then
        MsgBox "O preenchimento do campo CGC Raiz é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
        
    If OptAtiva.Value Then
        Status = "A"
    Else
        Status = "I"
    End If
    
    'Conexao.BeginTrans - Samuel 25/09/18 Alteração CGC
    'If PanContaCorrente.Tag = "I" Then
        'Sql = "INSERT INTO ContasCorrente (banc_cd_codigo, coco_cd_agencia, " & _
                    '"coco_nm_agencia, coco_tx_gerente, coco_tx_conta, plco_cd_conta, coco_tx_descricao, " & _
                    '"empr_cd_empresa, coco_nr_Dag, coco_nr_Dac, coco_nr_ConvenioFinanceiro, coco_nr_ConvenioCobranca, " & _
                    '"coco_nr_CarteiraCobranca, coco_nr_ArqCobranca, coco_nr_ArqForn, coco_tx_Instrucao1, coco_tx_Instrucao2,  " & _
                    '"coco_tx_Mensagem1, coco_tx_Mensagem2, coco_tx_Mensagem3, coco_tx_Mensagem4, coco_tx_provisao, coco_tx_status) " & _
              '"VALUES (" & MskBanco.Text & ",'" & TxtCodAgencia.Text & "'," & _
                    '"'" & TxtAgencia.Text & "','" & FunNulo(txtGerente.Text) & "','" & TxtNumeroConta.Text & "','" & DatConta.Recordset.Fields("plco_cd_conta") & "','" & FunTrataString(TxtDescricao.Text) & "', " & _
                    '"" & PCodEmpresa & ",'" & FunNulo(TxtDag.Text) & "','" & FunNulo(TxtDac.Text) & "','" & FunNulo(TxtConvenioFinanceiro.Text) & "','" & FunNulo(txtConvenioCobranca.Text) & "'," & _
                    '"'" & FunNulo(TxtCarteira.Text) & "','" & TxtArqCobranca.Text & "','" & TxtArqForn.Text & "','" & TxtInstrucao1.Text & "','" & FunNulo(TxtInstrucao2.Text) & "', " & _
                    '"'" & FunNulo(TxtMensagem1.Text) & "','" & FunNulo(TxtMensagem2.Text) & "','" & FunNulo(TxtMensagem3.Text) & "','" & FunNulo(TxtMensagem4.Text) & "','" & Provisao(CboProvisao) & "', '" & Status & "')"
        'Conexao.Execute Sql
        
        'MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) + "/" + CStr((Adodc1.Recordset.RecordCount + 1))
    
    Conexao.BeginTrans
    If PanContaCorrente.Tag = "I" Then
        Sql = "INSERT INTO ContasCorrente (banc_cd_codigo, coco_cd_agencia, " & _
                    "coco_nm_agencia, coco_tx_gerente, coco_tx_conta, plco_cd_conta, plco_cd_contaCusto, coco_tx_descricao, " & _
                    "empr_cd_empresa, coco_nr_Dag, coco_nr_Dac, coco_nr_ConvenioFinanceiro, coco_nr_ConvenioCobranca, " & _
                    "coco_nr_CarteiraCobranca, coco_nr_ArqCobranca, coco_nr_ArqForn, coco_tx_Instrucao1, coco_tx_Instrucao2,  " & _
                    "coco_tx_Mensagem1, coco_tx_Mensagem2, coco_tx_Mensagem3, coco_tx_Mensagem4, coco_tx_provisao, coco_tx_status, coco_tx_CGCRaiz) " & _
              "VALUES (" & MskBanco.Text & ",'" & TxtCodAgencia.Text & "'," & _
                    "'" & TxtAgencia.Text & "','" & FunNulo(txtGerente.Text) & "','" & TxtNumeroConta.Text & "','" & DatContaDespesa.Recordset.Fields("plco_cd_conta") & "','" & FunNulo(DatContaCusto.Recordset.Fields("plco_cd_conta")) & "','" & FunTrataString(TxtDescricao.Text) & "', " & _
                    "" & PCodEmpresa & ",'" & FunNulo(TxtDag.Text) & "','" & FunNulo(TxtDac.Text) & "','" & FunNulo(TxtConvenioFinanceiro.Text) & "','" & FunNulo(txtConvenioCobranca.Text) & "'," & _
                    "'" & FunNulo(TxtCarteira.Text) & "','" & TxtArqCobranca.Text & "','" & TxtArqForn.Text & "','" & TxtInstrucao1.Text & "','" & FunNulo(TxtInstrucao2.Text) & "', " & _
                    "'" & FunNulo(TxtMensagem1.Text) & "','" & FunNulo(TxtMensagem2.Text) & "','" & FunNulo(TxtMensagem3.Text) & "','" & FunNulo(TxtMensagem4.Text) & "','" & PROVISAO(CboProvisao) & "', '" & Status & "', " & _
                    "'" & MskCGCRaiz.Text & "')"
        Conexao.Execute Sql
        
        MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) + "/" + CStr((Adodc1.Recordset.RecordCount + 1))
    'Else    'Alteração - Samuel 25/09/18 Alteração CGC
    
        'Sql = "UPDATE ContasCorrente " & _
                      '"SET banc_cd_codigo = '" & Str(CInt(MskBanco.Text)) & "'," & _
                         '" coco_cd_agencia = '" & TxtCodAgencia.Text & "'," & _
                         '" coco_nm_agencia = '" & TxtAgencia.Text & "'," & _
                         '" coco_tx_gerente = '" & FunNulo(txtGerente.Text) & "'," & _
                         '" coco_tx_conta = '" & TxtNumeroConta.Text & "'," & _
                         '" plco_cd_conta = '" & DatConta.Recordset.Fields("plco_cd_conta") & "'," & _
                         '" coco_tx_descricao = '" & FunTrataString(TxtDescricao.Text) & "'," & _
                         '" coco_nr_Dag='" & FunNulo(TxtDag.Text) & "'," & _
                         '" coco_nr_Dac='" & FunNulo(TxtDac.Text) & "'," & _
                         '" coco_nr_ArqCobranca ='" & FunNulo(Me.TxtArqCobranca) & "', " & _
                         '" coco_nr_ArqForn ='" & FunNulo(Me.TxtArqForn) & "', " & _
                         '"coco_nr_ConvenioFinanceiro = '" & FunNulo(TxtConvenioFinanceiro.Text) & "'," & _
                         '"coco_nr_ConvenioCobranca = '" & FunNulo(txtConvenioCobranca.Text) & "'," & _
                         '"coco_nr_CarteiraCobranca= '" & FunNulo(TxtCarteira.Text) & "'," & _
                         '"coco_tx_Instrucao1='" & TxtInstrucao1.Text & "'," & _
                         '"coco_tx_Instrucao2='" & FunNulo(TxtInstrucao2.Text) & "'," & _
                         '"coco_tx_Mensagem1='" & FunNulo(TxtMensagem1.Text) & "'," & _
                         '"coco_tx_Mensagem2='" & FunNulo(TxtMensagem2.Text) & "'," & _
                         '"coco_tx_Mensagem3='" & FunNulo(TxtMensagem3.Text) & "'," & _
                         '"coco_tx_Mensagem4 ='" & FunNulo(TxtMensagem4.Text) & "', " & _
                         '"coco_tx_provisao = '" & Provisao(CboProvisao) & "', " & _
                         '"coco_tx_status = '" & Status & "'" & _
                     '" WHERE coco_cd_codigo = " & PanContaCorrente.Tag
        'Conexao.Execute Sql
    'End If
    
    Else    'Alteração
    
        Sql = "UPDATE ContasCorrente " & _
                      "SET banc_cd_codigo = '" & Str(CInt(MskBanco.Text)) & "'," & _
                         " coco_cd_agencia = '" & TxtCodAgencia.Text & "'," & _
                         " coco_nm_agencia = '" & TxtAgencia.Text & "'," & _
                         " coco_tx_gerente = '" & FunNulo(txtGerente.Text) & "'," & _
                         " coco_tx_conta = '" & TxtNumeroConta.Text & "'," & _
                         " plco_cd_conta = '" & DatContaDespesa.Recordset.Fields("plco_cd_conta") & "'," & _
                         " plco_cd_contaCusto = '" & FunNulo(DatContaCusto.Recordset.Fields("plco_cd_conta")) & "'," & _
                         " coco_tx_descricao = '" & FunTrataString(TxtDescricao.Text) & "'," & _
                         " coco_nr_Dag='" & FunNulo(TxtDag.Text) & "'," & _
                         " coco_nr_Dac='" & FunNulo(TxtDac.Text) & "'," & _
                         " coco_nr_ArqCobranca ='" & FunNulo(Me.TxtArqCobranca) & "', " & _
                         " coco_nr_ArqForn ='" & FunNulo(Me.TxtArqForn) & "', " & _
                         "coco_nr_ConvenioFinanceiro = '" & FunNulo(TxtConvenioFinanceiro.Text) & "'," & _
                         "coco_nr_ConvenioCobranca = '" & FunNulo(txtConvenioCobranca.Text) & "',"
                Sql = Sql & _
                         "coco_nr_CarteiraCobranca= '" & FunNulo(TxtCarteira.Text) & "'," & _
                         "coco_tx_Instrucao1='" & TxtInstrucao1.Text & "'," & _
                         "coco_tx_Instrucao2='" & FunNulo(TxtInstrucao2.Text) & "'," & _
                         "coco_tx_Mensagem1='" & FunNulo(TxtMensagem1.Text) & "'," & _
                         "coco_tx_Mensagem2='" & FunNulo(TxtMensagem2.Text) & "'," & _
                         "coco_tx_Mensagem3='" & FunNulo(TxtMensagem3.Text) & "'," & _
                         "coco_tx_Mensagem4 ='" & FunNulo(TxtMensagem4.Text) & "', " & _
                         "coco_tx_provisao = '" & PROVISAO(CboProvisao) & "', " & _
                         "coco_tx_status = '" & Status & "', " & _
                         "coco_tx_CGCRaiz = '" & MskCGCRaiz.Text & "'" & _
                     " WHERE coco_cd_codigo = " & PanContaCorrente.Tag
        Conexao.Execute Sql
    End If
    Conexao.CommitTrans
    
    PanContaCorrente.Visible = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    'Adodc1.Refresh TestesubrecarregadadosNV
    If PanContaCorrente.Tag <> "I" Then
'        subLocalizaData Adodc1, "coco_cd_codigo", CLng(PanContaCorrente.Tag)
        'cboLocalizar.BoundText = Chave
        cboLocalizar.BoundText = PanContaCorrente.Tag
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    Exit Sub
    
MostraErro:
    Conexao.RollbackTrans
   ' Enumera a coleção de erros e apresenta
   ' as propriedades de cada erro ocorrido
'   For Each errLoop In Conexao.Errors
'      StrError = "Error #" & errLoop.Number & vbCr & _
'         "   " & errLoop.Description & vbCr & _
'         "   (Source: " & errLoop.Source & ")" & vbCr & _
'         "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
'         "   (NativeError: " & errLoop.NativeError & ")" & vbCr
'
'   MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
'   Next
'
'   Resume Next
   MsgBox "Este código reduzido já foi atribuído a outro item do plano de contas.", vbCritical + vbOKOnly, "ATENÇÃO"
   Exit Sub
End Sub

Private Sub cmdExport_Click()
    Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    TDBGrid1.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To TDBGrid1.Columns.Count - 1
         If i = 0 And TDBGrid1.Columns(i).Visible Then
             sLine = Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         ElseIf TDBGrid1.Columns(i).Visible Then
             sLine = sLine & ";" & Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & Adodc1.Recordset.RecordCount & " registros", vbInformation, "Informação"

End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Contas Correntes")
End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim Fil As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaContaCorrente
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    
    If Filtrou = True Then
        Fil = ContFil
        subContaReg Adodc1
        'Adodc1.Refresh TestesubrecarregadadosNV
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
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

Sub subPrepararInsercao()
    MskBanco.Mask = ""
    MskBanco.Text = ""
    MskBanco.Mask = "9999"
    cboBancos.Text = ""
    
    MskContaDespesa.Mask = ""
    MskContaDespesa.Text = ""
    cboPlanoContasDespesa.Text = ""
    MskContaCusto.Mask = ""
    MskContaCusto.Text = ""
    cboPlanoContasCusto.Text = ""
    TxtDescricao.Text = ""
    TxtCodAgencia.Text = ""
    TxtAgencia.Text = ""
    txtGerente.Text = ""
    TxtDag.Text = ""
    TxtNumeroConta.Text = ""
    TxtDac.Text = ""
    TxtConvenioFinanceiro.Text = ""
    TxtCarteira.Text = ""
    TxtInstrucao1.Text = ""
    TxtInstrucao2.Text = ""
    TxtMensagem1.Text = ""
    TxtMensagem2.Text = ""
    TxtMensagem3.Text = ""
    TxtMensagem4.Text = ""
End Sub
Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    If Not FunVerifica_Permissao("TelaContaCorrente", "3") Then
        Exit Sub
    End If
    subPrepararInsercao
    PanBotoes.Enabled = False
    PanContaCorrente.Visible = True
    PanContaCorrente.Enabled = True
    PanContaCorrente.Left = (TDBGrid1.Width - PanContaCorrente.Width) / 2
    PanContaCorrente.Top = (TDBGrid1.Height - PanContaCorrente.Height) / 2
    PanContaCorrente.Tag = Adodc1.Recordset.Fields("coco_cd_codigo")
    PanContaCorrente.Caption = " Alteração de Conta Corrente"
    subDesabilitaBotoes
    OptAtiva.Value = True
    MskBanco.Text = Format$(Adodc1.Recordset.Fields("banc_cd_codigo"), "000")
    mskBanco_LostFocus
    
    If Not IsNull(Adodc1.Recordset.Fields("plco_cd_conta")) Then
        MskContaDespesa.Text = Adodc1.Recordset.Fields("plco_cd_codreduzido")
    End If
    MskContaDespesa_LostFocus
    
    If (Not IsNull(Adodc1.Recordset.Fields("plco_cd_contaCusto")) And Adodc1.Recordset.Fields("plco_cd_contaCusto") <> "") Then
        MskContaCusto.Text = Adodc1.Recordset.Fields("plco_cd_codreduzidoCusto")
    End If
    MskContaCusto_LostFocus
    
    'Configurar tamanho da cx de texto em relação ao banco
    Select Case MskBanco.Text
        Case "237", "104", "230" 'Bradesco,CEF,Bandeirante
            TxtMensagem1.Enabled = True
            TxtMensagem2.Enabled = True
            TxtMensagem3.Enabled = True
            TxtMensagem4.Enabled = True
            TxtInstrucao1.Enabled = False
            TxtInstrucao2.Enabled = False
            LblMensagem1(3).Enabled = True
            LblMensagem2(2).Enabled = True
            LblMensagem3(4).Enabled = True
            LblMensagem4(5).Enabled = True
            LblInstrucao1(0).Enabled = False
            LblInstrucao2(1).Enabled = False
            TxtMensagem1.MaxLength = 40
            TxtMensagem2.MaxLength = 40
            TxtMensagem3.MaxLength = 40
            TxtMensagem4.MaxLength = 40
        Case "237" 'Rural
            TxtMensagem1.Enabled = True
            TxtMensagem2.Enabled = True
            TxtMensagem3.Enabled = True
            TxtMensagem4.Enabled = True
            TxtInstrucao1.Enabled = False
            TxtInstrucao2.Enabled = False
            LblMensagem1(3).Enabled = True
            LblMensagem2(2).Enabled = True
            LblMensagem3(4).Enabled = True
            LblMensagem4(5).Enabled = True
            LblInstrucao1(0).Enabled = False
            LblInstrucao2(1).Enabled = False
            TxtMensagem1.MaxLength = 50
            TxtMensagem2.MaxLength = 50
            TxtMensagem3.MaxLength = 50
            TxtMensagem4.MaxLength = 50
        Case "399", "341"  'HSBC,Itaú
            TxtMensagem1.Enabled = False
            TxtMensagem2.Enabled = False
            TxtMensagem3.Enabled = False
            TxtMensagem4.Enabled = False
            TxtInstrucao1.Enabled = True
            TxtInstrucao2.Enabled = True
            LblMensagem1(3).Enabled = False
            LblMensagem2(2).Enabled = False
            LblMensagem3(4).Enabled = False
            LblMensagem4(5).Enabled = False
            LblInstrucao1(0).Enabled = True
            LblInstrucao2(1).Enabled = True
        Case "347" 'Sudameris
            TxtMensagem1.Enabled = True
            TxtMensagem2.Enabled = True
            TxtMensagem3.Enabled = True
            TxtMensagem4.Enabled = True
            TxtInstrucao1.Enabled = False
            TxtInstrucao2.Enabled = False
            LblMensagem1(3).Enabled = True
            LblMensagem2(2).Enabled = True
            LblMensagem3(4).Enabled = True
            LblMensagem4(5).Enabled = True
            LblInstrucao1(0).Enabled = False
            LblInstrucao2(1).Enabled = False
            TxtMensagem1.MaxLength = 55
            TxtMensagem2.MaxLength = 55
            TxtMensagem3.MaxLength = 55
            TxtMensagem4.MaxLength = 55
    End Select
    
    TxtCodAgencia.Text = Adodc1.Recordset.Fields("coco_cd_agencia")
    If Not IsNull(Adodc1.Recordset.Fields("coco_nm_agencia")) Then
        TxtAgencia.Text = Adodc1.Recordset.Fields("coco_nm_agencia")
    End If
    If Not IsNull(Adodc1.Recordset.Fields("coco_tx_gerente")) Then
        txtGerente.Text = Adodc1.Recordset.Fields("coco_tx_gerente")
    Else
        txtGerente.Text = ""
    End If
    TxtNumeroConta.Text = Adodc1.Recordset.Fields("coco_tx_conta")
    If Not IsNull(Adodc1.Recordset.Fields("coco_tx_descricao")) Then
        TxtDescricao.Text = Adodc1.Recordset.Fields("coco_tx_descricao")
    End If
    
    TxtDag.Text = FunNulo(Adodc1.Recordset.Fields("coco_nr_Dag"))
    TxtDac.Text = FunNulo(Adodc1.Recordset.Fields("coco_nr_Dac"))
    Me.TxtArqCobranca = FunNulo(Adodc1.Recordset.Fields("coco_nr_ArqCobranca"))
    Me.TxtArqForn = FunNulo(Adodc1.Recordset.Fields("coco_nr_ArqForn"))
    TxtConvenioFinanceiro.Text = FunNulo(Adodc1.Recordset.Fields("coco_nr_ConvenioFinanceiro"))
    txtConvenioCobranca.Text = FunNulo(Adodc1.Recordset.Fields("coco_nr_ConvenioCobranca"))
    TxtCarteira.Text = FunNulo(Adodc1.Recordset.Fields("coco_nr_CarteiraCobranca"))
    TxtInstrucao1.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Instrucao1"))
    TxtInstrucao2.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Instrucao2"))
    TxtMensagem1.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Mensagem1"))
    TxtMensagem2.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Mensagem2"))
    TxtMensagem3.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Mensagem3"))
    TxtMensagem4.Text = FunNulo(Adodc1.Recordset.Fields("coco_tx_Mensagem4"))
    'Alteração CGCRaiz - Samuel 26/09/18
    MskCGCRaiz.PromptInclude = False
    MskCGCRaiz.Text = FunNuloVal(Adodc1.Recordset.Fields("coco_tx_CGCRaiz"))
    MskCGCRaiz.PromptInclude = True
    
    If FunNulo(Adodc1.Recordset.Fields("coco_tx_provisao")) <> "" Then
        CboProvisao.Text = Adodc1.Recordset.Fields("coco_tx_provisao")
    Else
        'CboProvisao.Text = ""
        CboProvisao.ListIndex = -1
    End If
    
    If Adodc1.Recordset.Fields("coco_tx_status") = "Ativa" Then
        OptAtiva.Value = True
    ElseIf Adodc1.Recordset.Fields("coco_tx_status") = "Inativa" Then
        OptInativa.Value = True
    End If
    
    TDBGrid1.Enabled = False
    TxtDescricao.SetFocus

End Sub

Private Sub CmdCancelar_Click()
    If Not FunTabelaVazia(Adodc1) Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanContaCorrente.Visible = False
    PanContaCorrente.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If FunObrigatorioMSK(MskBanco, "Código do Banco é Obrigatório") Then Exit Function
       
    If FunObrigatorioTXT(TxtCodAgencia, "Código da Agência é Obrigatório") Then Exit Function
    
    If FunObrigatorioTXT(TxtNumeroConta, "Número da Conta é Obrigatória") Then Exit Function
    
    If FunObrigatorioMSK(MskContaDespesa, "Conta Contábil é Obrigatória") Then Exit Function
    
    If FunObrigatorioMSK(MskCGCRaiz, "CGC Raiz é Obrigatória") Then Exit Function
    
    funChecarCampos = False
End Function

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaContaCorrente
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

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    If Not FunVerifica_Permissao("TelaContaCorrente", "4") Then
        Exit Sub
    End If
    
    If CmdExcluir.Caption = "&Excluir" Then
        'If Not funVerifica_Permissao("FrmExemplo", OrdemRotina) Then Exit Sub
    End If

    Dim opcao As Integer
    Dim Houve_Erro As Integer
    Dim Posicao As Variant
    
    Houve_Erro = False
    
    Posicao = Adodc1.Recordset.Bookmark
    subDesabilitaBotoes
    
    opcao = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If opcao = 1 Then  'Confirma exclusão
       On Error GoTo RotuloErro
       Conexao.BeginTrans
       Conexao.Execute "DELETE FROM ContasCorrente WHERE coco_cd_codigo=" & Adodc1.Recordset.Fields("coco_cd_codigo")
'       Adodc1.Recordset.Delete
       Conexao.CommitTrans
       If Houve_Erro = False Then
          subConectarControleDadosNV Adodc1, FRM_FiltroAtual, "Estatico"
          Adodc1.Refresh
'          DataGrid1.Refresh
          'Quantid = Quantid - 1
          'MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)
          MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
          If ContFil <> 0 Then
             subContaReg Adodc1
          End If
       End If
       
       subHabilitaBotoes
       
       If FunTabelaVazia(Adodc1) = True Then
          CmdAlterar.Enabled = False
          CmdExcluir.Enabled = False
          CmdImprimir.Enabled = False
          CmdOrdem.Enabled = False
       End If
       
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.Bookmark = Posicao - 1
    End If
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Essa Conta Corrente não pode ser excluída pois já existem registros relacionados!", vbInformation, "ATENÇÃO"
    End If
    subHabilitaBotoes
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
    If Not FunVerifica_Permissao("TelaContaCorrente", "2") Then
        Exit Sub
    End If
    
    PanBotoes.Enabled = False
    PanContaCorrente.Visible = True
    PanContaCorrente.Enabled = True
    PanContaCorrente.Left = (TDBGrid1.Width - PanContaCorrente.Width) / 2
    PanContaCorrente.Top = (TDBGrid1.Height - PanContaCorrente.Height) / 2
    PanContaCorrente.Caption = " Inserção de Conta Corrente"
    PanContaCorrente.Tag = "I"
    'Chave = -1
    'XInserir = "I"
    subDesabilitaBotoes
    'Adodc1.Recordset.AddNew
    subPrepararInsercao
    cboLocalizar.Text = ""
    TDBGrid1.Enabled = False
    TxtDescricao.SetFocus
    OptAtiva.Value = True
    
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Grava na tabela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    End If
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


Private Sub MskCGCRaiz_LostFocus()
  MskCGCRaiz.PromptInclude = False
  If MskCGCRaiz.Text <> "" Then
      MskCGCRaiz.PromptInclude = True
      Exit Sub
    End If
  MskCGCRaiz.PromptInclude = True
  
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaContaCorrente
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, X
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False 'Desabilita o botão Excluir
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    Set Formulario = TelaContaCorrente
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaContaCorrente
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaContaCorrente"
    'Pesquisa por CGC Raiz - Samuel 26/09/18
    'Associa um banco de dados à tabela Descontos e Acréscimos
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subCarregaDadosNV Adodc1, "ConsFINContaCorrente", "banc_tx_descricao", _
                                "", "SELECT * FROM ConsFINContaCorrente WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "'"
    Else
        subCarregaDadosNV Adodc1, "ConsFINContaCorrente", "banc_tx_descricao", _
                            "", "SELECT * FROM ConsFINContaCorrente where empr_cd_empresa=" & PCodEmpresa
    End If
    
    'Inicializar Controle de Dados do GrdConta e do GrdBanco
    subConectarControleDadosNV DatBanco, "SELECT * FROM Bancos ORDER BY banc_tx_descricao", Estatico
    'subConectarControleDados DatConta, "SELECT * FROM PlanoContas ORDER BY plco_tx_desccontabil", Estatico
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV DatContaDespesa, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_Tipo = 'A' ORDER BY plco_tx_desccontabil", Estatico
        subConectarControleDadosNV DatContaCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_Tipo = 'A' ORDER BY plco_tx_desccontabil", Estatico
    Else
        subConectarControleDados DatContaDespesa, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY plco_cd_estruturado", Estatico
        subConectarControleDados DatContaCusto, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY plco_cd_estruturado", Estatico
    End If
    
    subHabilitaBotoes
    
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaContaCorrente"
End Sub


Private Sub MskBanco_GotFocus()
    Call subSelecionaMSK(MskBanco)
End Sub

Private Sub mskBanco_LostFocus()
    If IsNumeric(MskBanco.Text) Then
        cboBancos.BoundText = CLng(MskBanco.Text)
        If cboBancos.Text = "" Then
            MsgBox "Não existe Banco cadastrado com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskBanco.SetFocus
        End If
        'PnlBanco.Caption = TextoExisteGrid(DatBanco, "banc_cd_codigo", CLng(MskBanco.Text), "N", MskBanco, "Este Banco Não Existe !", "banc_tx_descricao", "0000")
    End If
    
    'Configurar tamanho da cx de texto em relação ao banco
    Select Case MskBanco.Text
        Case 237 Or 104 Or 230 'Bradesco,CEF,Bandeirante
            TxtMensagem1.MaxLength = 40
            TxtMensagem2.MaxLength = 40
            TxtMensagem3.MaxLength = 40
            TxtMensagem4.MaxLength = 40
        Case 237 'Rural
            TxtMensagem1.MaxLength = 50
            TxtMensagem2.MaxLength = 50
            TxtMensagem3.MaxLength = 50
            TxtMensagem4.MaxLength = 50
        Case 399 Or 341 'HSBC,Itaú
            TxtMensagem1.Enabled = False
            TxtMensagem2.Enabled = False
            TxtMensagem3.Enabled = False
            TxtMensagem4.Enabled = False
        Case 347 'Sudameris
            TxtMensagem1.MaxLength = 55
            TxtMensagem2.MaxLength = 55
            TxtMensagem3.MaxLength = 55
            TxtMensagem4.MaxLength = 55
    End Select
End Sub

Private Sub MskContaDespesa_GotFocus()
    Call subSelecionaMSK(MskContaDespesa)
End Sub

Private Sub MskContaCusto_GotFocus()
    Call subSelecionaMSK(MskContaCusto)
End Sub

Private Sub MskContaDespesa_LostFocus()
    If MskContaDespesa.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskContaDespesa.Text = Format(MskContaDespesa.Text, XGT_MSKCODREDUZIDO)
        End If
        cboPlanoContasDespesa.BoundText = MskContaDespesa.Text
        If cboPlanoContasDespesa.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskContaDespesa.SetFocus
        End If
        'PnlConta.Caption = TextoExisteGrid(DatConta, "plco_cd_Conta", MskConta.Text, "T", MskConta, "Esta Conta Não Existe !", "plco_tx_DescContabil", "")
    End If
End Sub

Private Sub MskContaCusto_LostFocus()
    If MskContaCusto.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskContaCusto.Text = Format(MskContaCusto.Text, XGT_MSKCODREDUZIDO)
        End If
        cboPlanoContasCusto.BoundText = MskContaCusto.Text
        If cboPlanoContasCusto.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskContaCusto.SetFocus
        End If
        'PnlConta.Caption = TextoExisteGrid(DatConta, "plco_cd_Conta", MskConta.Text, "T", MskConta, "Esta Conta Não Existe !", "plco_tx_DescContabil", "")
    End If
End Sub

Private Function PROVISAO(obj As ComboBox) As String
    If obj.ListIndex = 0 Then
        PROVISAO = "A"
    ElseIf obj.ListIndex = 1 Then
         PROVISAO = "B"
    ElseIf obj.ListIndex = 2 Then
        PROVISAO = "C"
    ElseIf obj.ListIndex = 3 Then
        PROVISAO = "M"
    ElseIf obj.ListIndex = 4 Then
        PROVISAO = "G"
    ElseIf obj.ListIndex = 5 Then
        PROVISAO = "P"
    End If
End Function
