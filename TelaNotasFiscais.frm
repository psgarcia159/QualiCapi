VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TelaNotasFiscais 
   Caption         =   "Manutenção e Consulta de Contas a Pagar"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   1245
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   9765
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanEstorno 
      Height          =   2745
      Left            =   10080
      TabIndex        =   39
      Top             =   4800
      Visible         =   0   'False
      Width           =   8640
      _Version        =   65536
      _ExtentX        =   15240
      _ExtentY        =   4842
      _StockProps     =   15
      Caption         =   " Estorno Financeiro"
      ForeColor       =   128
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
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox TxtProcPagto 
         Height          =   330
         Left            =   1545
         TabIndex        =   54
         Top             =   2115
         Width           =   1500
      End
      Begin VB.TextBox TxtCodTipoDoc 
         Height          =   315
         Left            =   4425
         MaxLength       =   4
         TabIndex        =   41
         Top             =   495
         Width           =   495
      End
      Begin MSDataListLib.DataCombo CboHistorico 
         Bindings        =   "TelaNotasFiscais.frx":0000
         Height          =   315
         Left            =   2070
         TabIndex        =   44
         Top             =   900
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "hist_tx_descricao"
         BoundColumn     =   "hist_cd_codigo"
         Text            =   ""
      End
      Begin VB.TextBox TxtHistorico 
         Height          =   315
         Left            =   1545
         MaxLength       =   4
         TabIndex        =   43
         Top             =   900
         Width           =   495
      End
      Begin VB.CommandButton CmdRetEstorno 
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
         Height          =   300
         Left            =   7365
         TabIndex        =   48
         Top             =   2190
         Width           =   1140
      End
      Begin VB.CommandButton CmdConfEstorno 
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
         Height          =   300
         Left            =   6135
         TabIndex        =   47
         Top             =   2190
         Width           =   1140
      End
      Begin MSAdodcLib.Adodc DatHistorico 
         Height          =   330
         Left            =   3360
         Top             =   2760
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
      Begin MSMask.MaskEdBox Mskconta 
         Height          =   315
         Left            =   1545
         TabIndex        =   45
         Top             =   1740
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo CboPlanoContas 
         Bindings        =   "TelaNotasFiscais.frx":001B
         Height          =   330
         Left            =   2415
         TabIndex        =   46
         Top             =   1740
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
      Begin MSAdodcLib.Adodc DatContaEstorno 
         Height          =   330
         Left            =   600
         Top             =   2760
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
      Begin MSComCtl2.DTPicker DtpEntrada 
         Height          =   315
         Left            =   1545
         TabIndex        =   40
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81723393
         CurrentDate     =   37460
      End
      Begin MSAdodcLib.Adodc DatTipoDoc 
         Height          =   330
         Left            =   3360
         Top             =   3240
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
         Bindings        =   "TelaNotasFiscais.frx":0039
         Height          =   315
         Left            =   4950
         TabIndex        =   42
         Top             =   495
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tido_tx_descricao"
         BoundColumn     =   "tido_cd_tipodoc"
         Text            =   "CboTipoDoc"
      End
      Begin MSMask.MaskEdBox MskPcr 
         Height          =   330
         Left            =   1545
         TabIndex        =   55
         Top             =   1305
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboPcr 
         Bindings        =   "TelaNotasFiscais.frx":0052
         Height          =   330
         Left            =   2415
         TabIndex        =   56
         Top             =   1305
         Width           =   6090
         _ExtentX        =   10742
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
      Begin MSAdodcLib.Adodc DatCREstorno 
         Height          =   330
         Left            =   600
         Top             =   3240
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
         Height          =   315
         Left            =   180
         TabIndex        =   57
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proc. Pagto.:"
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
         Left            =   135
         TabIndex        =   53
         Top             =   2175
         Width           =   1380
      End
      Begin VB.Label LblTipoDoc 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2925
         TabIndex        =   52
         Top             =   555
         Width           =   1515
      End
      Begin VB.Label LblEntrada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   330
         TabIndex        =   51
         Top             =   555
         Width           =   1260
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
         Left            =   135
         TabIndex        =   50
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label LblHistorico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   570
         TabIndex        =   49
         Top             =   960
         Width           =   945
      End
   End
   Begin Threed.SSPanel PanAltDistribRecolhimento 
      Height          =   2865
      Left            =   10080
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   9060
      _Version        =   65536
      _ExtentX        =   15981
      _ExtentY        =   5054
      _StockProps     =   15
      Caption         =   "Alteração de Distribuição de Recolhimento"
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
      Begin VB.CommandButton CmdDesisteDistribuicao 
         BackColor       =   &H00000000&
         Caption         =   "&Desistir"
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
         Left            =   7815
         TabIndex        =   38
         Top             =   2355
         Width           =   1140
      End
      Begin VB.CommandButton CmdGravaDistribuicao 
         BackColor       =   &H00000000&
         Caption         =   "&Gravar"
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
         Left            =   6585
         TabIndex        =   37
         Top             =   2355
         Width           =   1140
      End
      Begin Threed.SSPanel PanDistribuicao 
         Height          =   1575
         Left            =   105
         TabIndex        =   20
         Top             =   2800
         Width           =   8850
         _Version        =   65536
         _ExtentX        =   15610
         _ExtentY        =   2778
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
            Height          =   285
            Left            =   6465
            TabIndex        =   29
            Top             =   1230
            Width           =   1080
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
            Height          =   285
            Left            =   7635
            TabIndex        =   30
            Top             =   1230
            Width           =   1080
         End
         Begin VB.TextBox TxtVldistribuicao 
            Height          =   330
            Left            =   2820
            TabIndex        =   26
            Top             =   840
            Width           =   1365
         End
         Begin VB.TextBox TxtPercDistribuicao 
            Height          =   330
            Left            =   1485
            MaxLength       =   10
            TabIndex        =   25
            Top             =   840
            Width           =   660
         End
         Begin MSMask.MaskEdBox MskConta2 
            Height          =   330
            Left            =   1485
            TabIndex        =   23
            Top             =   465
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContas2 
            Bindings        =   "TelaNotasFiscais.frx":006D
            Height          =   330
            Left            =   2370
            TabIndex        =   24
            Top             =   465
            Width           =   6360
            _ExtentX        =   11218
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
            Height          =   315
            Left            =   1485
            TabIndex        =   21
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr2 
            Bindings        =   "TelaNotasFiscais.frx":0084
            Height          =   330
            Left            =   2370
            TabIndex        =   22
            Top             =   90
            Width           =   6360
            _ExtentX        =   11218
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
            Bindings        =   "TelaNotasFiscais.frx":0099
            Height          =   315
            Left            =   5415
            TabIndex        =   27
            Top             =   840
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "iden_tx_descricao"
            BoundColumn     =   "iden_cd_Projeto"
            Text            =   ""
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   315
            Left            =   8370
            TabIndex        =   28
            Top             =   840
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "TelaNotasFiscais.frx":00B8
            Picture         =   "TelaNotasFiscais.frx":050A
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   35
            Top             =   510
            Width           =   1350
         End
         Begin VB.Label Label4 
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
            TabIndex        =   34
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label3 
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
            Left            =   1995
            TabIndex        =   33
            Top             =   885
            Width           =   810
         End
         Begin VB.Label Label2 
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
            Left            =   480
            TabIndex        =   32
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Left            =   4245
            TabIndex        =   31
            Top             =   885
            Width           =   1155
         End
      End
      Begin VB.CommandButton CmdAlteraDistribuicao 
         Caption         =   "&Al&terar"
         Height          =   285
         Left            =   7860
         TabIndex        =   17
         Top             =   690
         Width           =   1080
      End
      Begin VB.CommandButton CmdRemoverDistribuicao 
         Caption         =   "Excl&uir"
         Height          =   285
         Left            =   7860
         TabIndex        =   18
         Top             =   1050
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDistribuicao 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ins&erir"
         Height          =   285
         Left            =   7860
         TabIndex        =   16
         Top             =   345
         Width           =   1080
      End
      Begin MSAdodcLib.Adodc DatPcr 
         Height          =   330
         Left            =   240
         Top             =   4680
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
         Height          =   405
         Left            =   2760
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   714
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
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   1830
         Left            =   120
         TabIndex        =   19
         Top             =   375
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   3228
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
      Begin VB.Label LblValorTotalDistrib 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7815
         TabIndex        =   36
         Top             =   1890
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel panCancelar 
      Height          =   1290
      Left            =   13200
      TabIndex        =   58
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   2275
      _StockProps     =   15
      Caption         =   " Cancelamento"
      ForeColor       =   12582912
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
      BevelInner      =   1
      Alignment       =   0
      Begin VB.CommandButton cmdRetornarCancelamento 
         Caption         =   "Retornar"
         Height          =   255
         Left            =   2880
         TabIndex        =   59
         Top             =   960
         Width           =   870
      End
      Begin MSComCtl2.DTPicker dtpCancelamento 
         Height          =   315
         Left            =   2400
         TabIndex        =   61
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81723393
         CurrentDate     =   37460
      End
      Begin VB.CommandButton cmdConfirmarCancelamento 
         Caption         =   "&Confirmar"
         Height          =   255
         Left            =   1920
         TabIndex        =   60
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblCancelamento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data de Cancelamento:"
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
         TabIndex        =   62
         Top             =   540
         Width           =   2175
      End
   End
   Begin Threed.SSPanel PanTextoRecibo 
      Height          =   2220
      Left            =   2880
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   6930
      _Version        =   65536
      _ExtentX        =   12224
      _ExtentY        =   3916
      _StockProps     =   15
      Caption         =   " Texto do Recibo"
      ForeColor       =   128
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
      BevelInner      =   1
      Alignment       =   0
      Begin VB.CommandButton CmdConfirmaRecibo 
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
         Height          =   300
         Left            =   4425
         TabIndex        =   14
         Top             =   1785
         Width           =   1140
      End
      Begin VB.CommandButton CmdSairRecibo 
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
         Height          =   300
         Left            =   5655
         TabIndex        =   13
         Top             =   1785
         Width           =   1140
      End
      Begin VB.TextBox TxtTextoRecibo 
         Height          =   1335
         Left            =   150
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "TelaNotasFiscais.frx":061C
         Top             =   300
         Width           =   6630
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaNotasFiscais.frx":0622
      Height          =   3435
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6059
      _LayoutType     =   4
      _RowHeight      =   30
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Entrada"
      Columns(0).DataField=   "nofi_dt_entrada"
      Columns(0).NumberFormat=   "dd/mm/yy"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   80
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   "P"
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   "Previsto"
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(1)._DefaultItem=   0
      Columns(1).ValueItems(1).Value=   "C"
      Columns(1).ValueItems(1).Value.vt=   8
      Columns(1).ValueItems(1).DisplayValue=   "Compromissado"
      Columns(1).ValueItems(1).DisplayValue.vt=   8
      Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(2)._DefaultItem=   0
      Columns(1).ValueItems(2).Value=   "R"
      Columns(1).ValueItems(2).Value.vt=   8
      Columns(1).ValueItems(2).DisplayValue=   "Realizado"
      Columns(1).ValueItems(2).DisplayValue.vt=   8
      Columns(1).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   3
      Columns(1).Caption=   "Status"
      Columns(1).DataField=   "nofi_tx_status"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fornecedor"
      Columns(2).DataField=   "focl_tx_fantasia"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo Doc."
      Columns(3).DataField=   "tido_tx_descricao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Título QualiCapi"
      Columns(4).DataField=   "nofi_nr_TituloCapi"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Nº Doc"
      Columns(5).DataField=   "nofi_nr_documento"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Série"
      Columns(6).DataField=   "nofi_tx_serie"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Valor"
      Columns(7).DataField=   "nofi_vl_valor"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Emissão"
      Columns(8).DataField=   "nofi_dt_emissao"
      Columns(8).NumberFormat=   "dd/mm/yy"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Nº Proc. Pagto."
      Columns(9).DataField=   "nofi_nr_fatura"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Histórico"
      Columns(10).DataField=   "nofi_tx_historico"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Exportado"
      Columns(11).DataField=   "nofi_tx_contabil"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Codigo"
      Columns(12).DataField=   "nofi_cd_notafiscal"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Controle"
      Columns(13).DataField=   "nofi_tx_controle"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Estoque"
      Columns(14).DataField=   "Estoque"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Provisao"
      Columns(15).DataField=   "Provisao"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(0)._MinWidth=128"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=2170"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(1)._MinWidth=103613908"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=4471"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=4392"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2461"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2381"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=528"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(3)._MinWidth=104758304"
      Splits(0)._ColumnProps(32)=   "Column(4).Width=3413"
      Splits(0)._ColumnProps(33)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._WidthInPix=3334"
      Splits(0)._ColumnProps(35)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=17"
      Splits(0)._ColumnProps(37)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(39)=   "Column(4)._MinWidth=104734912"
      Splits(0)._ColumnProps(40)=   "Column(5).Width=1931"
      Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=1852"
      Splits(0)._ColumnProps(43)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(5)._ColStyle=528"
      Splits(0)._ColumnProps(45)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(47)=   "Column(5)._MinWidth=104734912"
      Splits(0)._ColumnProps(48)=   "Column(6).Width=979"
      Splits(0)._ColumnProps(49)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(6)._WidthInPix=900"
      Splits(0)._ColumnProps(51)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(52)=   "Column(6)._ColStyle=17"
      Splits(0)._ColumnProps(53)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(54)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(55)=   "Column(7).Width=2090"
      Splits(0)._ColumnProps(56)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(7)._WidthInPix=2011"
      Splits(0)._ColumnProps(58)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(7)._ColStyle=786"
      Splits(0)._ColumnProps(60)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(61)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(62)=   "Column(8).Width=1508"
      Splits(0)._ColumnProps(63)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(8)._WidthInPix=1429"
      Splits(0)._ColumnProps(65)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(8)._ColStyle=529"
      Splits(0)._ColumnProps(67)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(68)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(69)=   "Column(9).Width=1561"
      Splits(0)._ColumnProps(70)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(9)._WidthInPix=1482"
      Splits(0)._ColumnProps(72)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(73)=   "Column(9)._ColStyle=17"
      Splits(0)._ColumnProps(74)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(75)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(76)=   "Column(10).Width=9022"
      Splits(0)._ColumnProps(77)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(10)._WidthInPix=8943"
      Splits(0)._ColumnProps(79)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(10)._ColStyle=528"
      Splits(0)._ColumnProps(81)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(83)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(87)=   "Column(11)._ColStyle=528"
      Splits(0)._ColumnProps(88)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(89)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(90)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(91)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(92)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(94)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(95)=   "Column(12)._ColStyle=528"
      Splits(0)._ColumnProps(96)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(98)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(99)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(100)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(102)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(103)=   "Column(13)._ColStyle=528"
      Splits(0)._ColumnProps(104)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(105)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(106)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(107)=   "Column(14).Width=2249"
      Splits(0)._ColumnProps(108)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(14)._WidthInPix=2170"
      Splits(0)._ColumnProps(110)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(111)=   "Column(14)._ColStyle=532"
      Splits(0)._ColumnProps(112)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(113)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(114)=   "Column(14)._MinWidth=116606384"
      Splits(0)._ColumnProps(115)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(116)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(117)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(118)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(119)=   "Column(15)._ColStyle=20"
      Splits(0)._ColumnProps(120)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(121)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(122)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(123)=   "Column(15)._MinWidth=-1"
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
      MultiSelect     =   2
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
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=98,.parent=13,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=90,.parent=13,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14,.alignment=1"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=0"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=0"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=0"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=76,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=77,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=86,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=83,.parent=14,.alignment=2"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=84,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=85,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(100) =   "Named:id=33:Normal"
      _StyleDefs(101) =   ":id=33,.parent=0"
      _StyleDefs(102) =   "Named:id=34:Heading"
      _StyleDefs(103) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   ":id=34,.wraptext=-1"
      _StyleDefs(105) =   "Named:id=35:Footing"
      _StyleDefs(106) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(107) =   "Named:id=36:Selected"
      _StyleDefs(108) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(109) =   "Named:id=37:Caption"
      _StyleDefs(110) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(111) =   "Named:id=38:HighlightRow"
      _StyleDefs(112) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(113) =   "Named:id=39:EvenRow"
      _StyleDefs(114) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(115) =   "Named:id=40:OddRow"
      _StyleDefs(116) =   ":id=40,.parent=33"
      _StyleDefs(117) =   "Named:id=41:RecordSelector"
      _StyleDefs(118) =   ":id=41,.parent=34"
      _StyleDefs(119) =   "Named:id=42:FilterBar"
      _StyleDefs(120) =   ":id=42,.parent=33"
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaNotasFiscais.frx":0637
      Height          =   315
      Left            =   2490
      TabIndex        =   8
      Top             =   3495
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "nofi_dt_entrada"
      BoundColumn     =   "nofi_cd_notafiscal"
      Text            =   ""
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   930
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   9540
      _Version        =   65536
      _ExtentX        =   16828
      _ExtentY        =   1640
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
         Left            =   0
         Picture         =   "TelaNotasFiscais.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkTotalFiltro 
         Caption         =   "Somar Valores"
         Height          =   255
         Left            =   4560
         TabIndex        =   64
         Top             =   0
         Width           =   1455
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
         Height          =   285
         Left            =   4095
         TabIndex        =   6
         Top             =   240
         Width           =   1155
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
         Height          =   285
         Left            =   7695
         TabIndex        =   7
         Top             =   240
         Width           =   1110
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
         Height          =   285
         Left            =   6510
         TabIndex        =   5
         Top             =   240
         Width           =   1130
      End
      Begin VB.CommandButton Botao_Fil 
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
         Height          =   285
         Left            =   5310
         TabIndex        =   4
         Top             =   240
         Width           =   1130
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
         Height          =   285
         Left            =   2910
         TabIndex        =   3
         Top             =   240
         Width           =   1130
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
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Top             =   270
         Width           =   1130
      End
      Begin VB.CommandButton CmdAlterar 
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
         Height          =   285
         Left            =   1710
         TabIndex        =   2
         Top             =   240
         Width           =   1130
      End
      Begin VB.Label lblTotalFiltro 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
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
         TabIndex        =   63
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7665
      Top             =   3480
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
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Entrada:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3540
      Width           =   2400
   End
End
Attribute VB_Name = "TelaNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor3 As New XArray
Dim XPcr As String  'Armazena o pcr digitado na distribuicao
Dim xcc As String   'Armazena a conta contábil digitada na distribuicao
Dim XDistribuicao As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDistribuido As String 'verifica se o valor total foi distribuido
Dim ValorTotalDistribuicao As Double 'Acumula o total de distribuições até compor o valor total da Nota Fiscal
Dim xNumRecibo As String
Dim XFT_TEXTO As String     'Texto do recibo
Dim XLT_ESTORNO As String

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

'Variável utilizada para controle de Contas associadas a CRs (CboPcr e CboPlanoContas) - 18/03/2022 - Paulo Garcia
Dim XFT_CR As String

Function FunVerificaRecolhimento() As Boolean
    
    Dim XLT_SQL As String
    Dim XLO_CHECK As ADODB.Recordset
    
    XLT_SQL = "SELECT desc_cd_desconto FROM DescontosNF WHERE (deen_dt_recolhimento <> '' OR NOT deen_dt_recolhimento IS NULL) AND deen_nr_fatura = '" & Adodc1.Recordset.Fields("nofi_nr_fatura") & "'"
    
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    
    If Not XLO_CHECK.EOF Then
        FunVerificaRecolhimento = False
        Exit Function
    End If
    
    FunVerificaRecolhimento = True
    
End Function
Function FunVerificacoes() As Boolean
    
    Dim XLT_SQL As String
    Dim XLO_CHECK As ADODB.Recordset
    
    'prestacoes de contas e recibos de adiantamento não são provisionados portanto tb não sao cancelados
     
    'Verifica se o mês está fechado
    If funMesFechado(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy")) <> "A" Then
        MsgBox "Não é possível realizar a operação nessa data, pois o mês contábil já foi fechado. " _
            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
        Me.panCancelar.Visible = False
        FunVerificacoes = False
        Exit Function
    End If
    
    'Verificar se o tipo de documento permite cancelamento
    'notas que não tem provisão não devem ser canceladas
    XLT_SQL = "SELECT tido_tx_Nec FROM tipodocumentos WHERE tido_cd_tipodoc = '" & Adodc1.Recordset.Fields("tido_cd_tipodoc") & "'"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If Not XLO_CHECK.EOF Then
        If XLO_CHECK!tido_tx_Nec = 1 Then
            MsgBox "Este tipo de documento não pode ser cancelado pois ele não é exportado para a contabilidade.", vbCritical, "ATENÇÃO"
            Me.panCancelar.Visible = False
            XLO_CHECK.Close
            FunVerificacoes = False
            Exit Function
        End If
    End If
    XLO_CHECK.Close
    
    'verificar se trata-se de um imposto. impostos não são provisionados
    If Me.Adodc1.Recordset("focl_tx_razaosocial").Value = "" Then
        MsgBox "Este tipo de documento não pode ser cancelado!", vbCritical, "ATENÇÃO"
         Me.panCancelar.Visible = False
         FunVerificacoes = False
         Exit Function
    End If
    
'    'verifica se a empresa possui contabilidade base caixa ou provisao e caixa
'    If XGT_TIPOEXPORTACAO <> "PC" Then
'        MsgBox "Esta operação só pode ser realizada para clientes" _
'            + Chr(13) + "que usam a contabilidade base provisão e caixa.!", vbCritical, "ATENÇÃO"
'         Me.panCancelar.Visible = False
'         FunVerificacoes = False
'         Exit Function
'    End If
    
    'Verifica se a data de cancelamento à data de entrada da nota
    If Me.dtpCancelamento.Value < Me.Adodc1.Recordset("nofi_dt_entrada") Then
        MsgBox "A data de cancelamento é posterior à data de entrada da nota. " _
            + Chr(13) + "Esta operação não pode ser realizada!", vbCritical, "ATENÇÃO"
         Me.panCancelar.Visible = False
         FunVerificacoes = False
         Exit Function
    End If
    
'    'verifica se exite alguma provisão para a nota/duplicata selecionada
'    XLT_SQL = "SELECT lancamentos.lanc_dt_data FROM lancamentos LEFT JOIN notasfiscais ON lancamentos.lanc_cd_origemfin = notasfiscais.nofi_cd_notafiscal " & _
'                    "WHERE (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T') AND nofi_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & ""
'    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
'    If XLO_CHECK.EOF Then
'            MsgBox "Não foi encontrada nenhum lançamento de provisão na contabilidade para essa duplicata." & _
'            Chr(13) + "Verifique se é um lançamento base caixa. Lançamentos base caixa não podem ser cancelados.", vbCritical, "ATENÇÃO"
'            Me.panCancelar.Visible = False
'            XLO_CHECK.Close
'            FunVerificacoes = False
'            Exit Function
'    End If
'    XLO_CHECK.Close
    
    FunVerificacoes = True
    
End Function

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

Sub subContabilizaRecolhimento(XLL_FATURA As Long)
    '********************DECRIÇÃO DA PROCEDURE***********************************
    'Contabilização do recolhimento em virtude da alteração da distribuição dos mesmos
    '****************************************************************************
    Dim XLL_CODCCORRENTE As Long    'Código da conta corrente
    Dim XLI_CODLANC As Integer      'Código para a tabela de Lançamentos
    Dim XLI_LOTE As Integer         'Nº do Lote
    Dim XLI_CODLANCBANCO As Integer 'Código de lançamento do banco
    Dim XLI_CONT As Integer         'Contador para o nº de parcelas
    Dim XLO_PARCELA As Object       'Recordset com os dados das parcelas
    Dim XLO_DESCONTOS As Object     'Recordset com os dados das parcelas
    Dim XLO_CHECK As Object         'Recordset para  definir o lote
    Dim XLT_HISTBANCO As String     'Histórico utilizado para o Banco
    Dim XLT_TIPOPAGTO As String     'Tipo de Pagto/rec (se cheque, dinheiro...)
    Dim XLT_PCRDEBITO As String, XLT_PCRCREDITO As String         'C. Custo de débito e crédito
    Dim XLT_CONTADEBITO As String, XLT_CONTACREDITO As String     'C. Contábil de débito e crédito
    Dim XLT_CONTADEBITOCC As String, XLT_CONTACREDITOCC As String 'C. Contábil (Banco ou forn) de débito e crédito
    Dim XLT_HISTORICO As String     'Histórico Contábil
    Dim XLT_HISTCUSTO As String     'Histórico do Custos = financeiro
    Dim XLT_DOCPAGREC As String     'Nº do doc pag/rec
    Dim XLF_VALOR As Currency       'Valor do lançamento
    Dim XLF_VALORBANCO As Currency  'Valor da parte do Banco
    Dim XLF_VALORDESC As Currency   'Valor de Desconto
    Dim XLF_VALORDESCACRESC As Currency  'Valor do desconto acrescimo
    Dim XLB_GRAVARBANCO As Boolean  'INDICA SE A PARTE DO BANCO JÁ PODE SER GRAVADA
    Dim XLB_GRAVARDESCACRESC As Boolean 'INDICA SE A PARTE DE DESC ACREC JÁ PODE SER GRAVADA
    Dim XLL_DUPLICATA As Long       'Identificar os descontos acréscimos dessa duplicata
    Dim XLL_CODDESC As Long, XLL_CODDESCPROX As Long
    Dim XLF_TOTALPORCODDESC As Currency
    Dim XLI_CONTDESC As Integer
    Dim XLF_TOTALGERALDESC As Currency
    Dim XLD_PAGTOREC As Date
    Dim XLB_CONSOLIDADO As Boolean
    Dim XLT_SQL As String
    Dim XLT_TIPO As String
    
    XLB_GRAVARBANCO = False
    XLB_GRAVARDESCACRESC = False
    
    XLT_SQL = "SELECT * FROM ConsFINExpContabilidade2 " & _
                "WHERE dupl_nr_fatura = " & XLL_FATURA & "" & _
                    " AND (dupl_tx_Estorno is null OR dupl_tx_Estorno = 'N' OR dupl_tx_Estorno = '') " & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
    SubQOpenRecordset XLO_PARCELA, XLT_SQL, Estatico
    If Not XLO_PARCELA.EOF Then
        XLI_CONT = XLO_PARCELA.RecordCount
        XLT_TIPO = XLO_PARCELA!dupl_tx_debcred
        XLT_TIPOPAGTO = XLO_PARCELA!dupl_tx_tipopagrec
        XLL_FATURA = XLO_PARCELA!dupl_nr_fatura
        XLL_CODCCORRENTE = FunNuloVal(XLO_PARCELA!coco_cd_codigo)
        XLT_DOCPAGREC = XLO_PARCELA!dupl_tx_docpagrec
        XLD_PAGTOREC = XLO_PARCELA!dupl_dt_pagrec
        XLL_DUPLICATA = XLO_PARCELA!dupl_cd_Duplicata
        If XLO_PARCELA!dupl_tx_consolidado = "S" Then
            XLB_CONSOLIDADO = True
        Else
            XLB_CONSOLIDADO = False
        End If

        'Pegar o número do Lote e Lancamento para cadastrar no mesmo lote do dia
        'Origem = FN - Financeiro Nota ou FD - Financeiro Duplicata
        XLT_SQL = "SELECT * FROM Lancamentos " & _
                    "WHERE lanc_dt_data = " & FunNuloData(XLO_PARCELA!dupl_dt_pagrec, NomeSgbd) & " AND " & _
                        "lanc_tx_origem LIKE 'F%' AND " & _
                        "empr_cd_empresa = " & PCodEmpresa & _
                     " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        If XLO_CHECK.EOF Then
            XLI_LOTE = (FunDefinirUltimoLote(XLO_PARCELA!dupl_dt_pagrec))
        Else
            XLI_LOTE = XLO_CHECK!lanc_nr_Lote
        End If
        XLO_CHECK.Close
       
        'Apagar os Lançamentos contábeis relacionados com esta Duplicata para gravar novos
        XLT_SQL = "DELETE FROM Lancamentos " & _
                    "WHERE lanc_cd_OrigemFin = " & XLO_PARCELA!dupl_cd_Duplicata & " AND " & _
                            "lanc_tx_Origem = 'FD' AND " & _
                            "not plco_cd_Debito is null AND " & _
                            "empr_cd_empresa = " & PCodEmpresa
        Conexao.Execute XLT_SQL
        
        XLI_CONTDESC = 0
        While Not XLO_PARCELA.EOF
            '***********Verificações sobre a parte do banco**************
            XLT_CONTACREDITO = XLO_PARCELA!PlcoContaCC
            If XLI_CONT > 1 Then
                If XLB_CONSOLIDADO = False Then
                    'Inserir os lançamentos a débito e um único registro no banco
                    XLB_GRAVARBANCO = True
                    XLT_CONTACREDITO = ""
                    XLT_CONTACREDITOCC = XLO_PARCELA!PlcoContaCC
                Else
                    'Inserir os lançamentos a débito e não precisa lançar a parte do banco
                    'por que o valor já foi contabilizado anteriormente no total consolidado e esse valor não será modificado
                    XLB_GRAVARBANCO = False
                    XLT_CONTACREDITO = ""
                    XLT_CONTACREDITOCC = ""
                End If
            Else
                If XLB_CONSOLIDADO = False Then
                    'Inserir o lançamento a débito e a crédito
                    XLB_GRAVARBANCO = False
                    XLT_CONTACREDITO = XLO_PARCELA!PlcoContaCC
                    XLT_CONTACREDITOCC = ""
                Else
                    'Inserir o lançamento a débito e não precisa lançar a parte do banco
                    'por que o valor já foi contabilizado anteriormente no total consolidado e esse valor não será modificado
                    XLB_GRAVARBANCO = False
                    XLT_CONTACREDITO = ""
                    XLT_CONTACREDITOCC = ""
                End If
            End If
            
            If (XLO_PARCELA!desconto <> 0 Or XLO_PARCELA!Acrescimo <> 0) Then
                'Inserir os lançamentos a débito e um único registro no banco
                XLB_GRAVARDESCACRESC = True
            End If
            '************************************************************
            XLI_CODLANC = funDefinirCodigoLanc(XLD_PAGTOREC, XLI_LOTE)
            
            XLF_VALORDESC = "0"
            XLF_VALORDESC = XLF_VALORDESC / XLI_CONT
            XLT_PCRDEBITO = ""
            XLT_CONTADEBITO = XLO_PARCELA!ContaRecDesp
            XLF_VALOR = XLO_PARCELA!distvalor
                    
            '**********MONTAR HISTÓRICO***********
            If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                If XLT_TIPOPAGTO = "1" Then 'Cheque
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. CHQ. Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. CHQ. Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    End If
                ElseIf XLT_TIPOPAGTO = "2" Or XLT_TIPOPAGTO = "3" Then 'Borderô
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. BORDERÔ Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. BORDERÔ Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    End If
                Else
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                    End If
                End If
            Else
                If XLT_TIPOPAGTO = "1" Then 'Cheque
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. CHQ. Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. CHQ. Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    End If
                ElseIf XLT_TIPOPAGTO = "2" Or XLT_TIPOPAGTO = "3" Then 'Borderô
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. BORDERÔ Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. BORDERÔ Nº " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    End If
                Else
                    If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then 'Nº automatico
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-" + "PAGTO. " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    Else
                        XLT_HISTORICO = Mid(XLO_PARCELA!tido_tx_descricao + "-Doc: " + XLO_PARCELA!nofi_nr_documento + "-" + "PAGTO. " + XLO_PARCELA!dupl_tx_docpagrec + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                    End If
                End If
            End If
            '*********************************
            
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
                        "'" & FunNulo(Left(FunTrataString(XLT_HISTORICO), 60)) & "'," & FunTrataFloat(XLF_VALOR) & "," & _
                        "" & PCodEmpresa & ",'" & FunNulo(FunTrataString(XLO_PARCELA!nofi_tx_historico)) & "'," & _
                        "" & FunNuloBanco(XLO_PARCELA!iden_cd_Projeto) & "," & XLO_PARCELA!dupl_cd_Duplicata & ",'FD')"
            Conexao.Execute XLT_SQL
    
            XLF_VALORBANCO = XLF_VALORBANCO + XLF_VALOR
            XLI_CODLANCBANCO = XLI_CODLANC
            XLT_HISTBANCO = XLT_HISTORICO
            XLF_VALOR = 0
                                
            XLO_PARCELA.MoveNext
            If XLO_PARCELA.EOF Then
                XLO_PARCELA.MovePrevious
                '******************************DESCONTOS/ACRÉSCIMOS******************************
                'Se existir atualizar o valor da duplicata e gravar o registro do desconto/acrescimo
                If XLB_GRAVARDESCACRESC And (XLO_PARCELA!desconto <> 0 Or XLO_PARCELA!Acrescimo <> 0) Then
                    XLT_SQL = "SELECT * FROM ConsFINExpDescacresc WHERE dupl_cd_duplicata = " & XLL_DUPLICATA & " and empr_cd_empresa = " & PCodEmpresa & ""
                    SubQOpenRecordset XLO_DESCONTOS, XLT_SQL, Estatico
                        
                    XLF_TOTALGERALDESC = 0
                    XLF_TOTALPORCODDESC = 0
                    While Not XLO_DESCONTOS.EOF
                        XLL_CODDESC = XLO_DESCONTOS!deac_cd_descacresc
                        XLL_CODDESCPROX = XLO_DESCONTOS!deac_cd_descacresc
                        If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                            If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                                XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                            Else
                                XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
                            End If
                        Else
                            If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                                XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                            Else
                                XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
                            End If
                        End If
                        
                        'Se for receita ou despesa, deve informar o Centro de custo
                        If XLO_DESCONTOS!natu_cd_Natureza = 3 Or XLO_DESCONTOS!natu_cd_Natureza = 4 Then
                            If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                                XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_conta
                                XLT_PCRDEBITO = XLO_DESCONTOS!cere_cd_Pcr
                                XLT_PCRCREDITO = ""
                                XLT_CONTACREDITO = ""
                            Else 'Desconto
                                XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_conta
                                XLT_PCRCREDITO = XLO_DESCONTOS!cere_cd_Pcr
                                XLT_PCRDEBITO = ""
                                XLT_CONTADEBITO = ""
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
                            If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                                XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_conta
                                XLT_PCRDEBITO = ""
                                XLT_PCRCREDITO = ""
                                XLT_CONTACREDITO = ""
                            Else 'Desconto
                                XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_conta
                                XLT_PCRCREDITO = ""
                                XLT_PCRDEBITO = ""
                                XLT_CONTADEBITO = ""
                            End If
                            
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
                                                                                
                        XLF_TOTALGERALDESC = XLF_TOTALGERALDESC + XLF_VALOR
                        If (XLF_TOTALGERALDESC - (XLO_PARCELA!desconto * -1)) = 0.01 Then
                            XLF_VALOR = XLF_VALOR - 0.01
                        End If
                        
                        XLI_CODLANC = funDefinirCodigoLanc(XLD_PAGTOREC, XLI_LOTE)
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
                                "" & PCodEmpresa & ",'" & FunNulo(FunTrataString(XLO_PARCELA!nofi_tx_historico)) & "'," & _
                                "" & FunNuloBanco(XLO_PARCELA!iden_cd_Projeto) & "," & XLO_PARCELA!dupl_cd_Duplicata & ",'FD')"
                        Conexao.Execute XLT_SQL
                        XLO_DESCONTOS.MoveNext
                    Wend
                    XLO_DESCONTOS.Close
                End If
            
                '***********************GRAVA REGISTRO DO BANCO**************************
                If XLB_GRAVARBANCO Then
                    XLF_VALORDESCACRESC = (XLO_PARCELA!desconto + XLO_PARCELA!Acrescimo)
                    XLF_VALORBANCO = XLF_VALORBANCO + XLF_VALORDESCACRESC
                    XLI_CODLANC = funDefinirCodigoLanc(XLD_PAGTOREC, XLI_LOTE)
                    
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_nr_Lote, lanc_cd_Lanc, " & _
                                    " lanc_dt_Data, " & _
                                    " plco_cd_Credito, plco_cd_Debito," & _
                                    " lanc_tx_historico, lanc_vl_valor," & _
                                    " empr_cd_empresa, lanc_tx_DescCustos, " & _
                                    " iden_cd_Projeto, lanc_cd_OrigemFin, lanc_tx_Origem) " & _
                                " VALUES (" & XLI_LOTE & ", " & XLI_CODLANC & "," & _
                                    "" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "," & _
                                    "" & FunNuloBanco(XLT_CONTACREDITOCC) & "," & FunNuloBanco(XLT_CONTADEBITOCC) & "," & _
                                    "'" & FunNulo(XLT_HISTBANCO) & "'," & FunTrataFloat(XLF_VALORBANCO) & "," & _
                                    "" & PCodEmpresa & ",'" & FunTrataString(FunNulo(XLO_PARCELA!nofi_tx_historico)) & "'," & _
                                    "" & FunNuloBanco(XLO_PARCELA!iden_cd_Projeto) & "," & XLO_PARCELA!dupl_cd_Duplicata & ",'FD')"
                    Conexao.Execute XLT_SQL
                    XLB_GRAVARBANCO = False
                End If
                XLO_PARCELA.MoveNext
            End If
        Wend
        XLO_PARCELA.Close
    End If
End Sub


'Sub subImprimeRecibo(NumRecibo As Double)
'    Dim XCidade As String, XLT_SQL As String, XLT_DATAPAGREC As String
'    Dim XFornecedor As String, Xhistorico As String
'    Dim XRecibo As String, XExtenso As String, XLT_TEMDESC As String
'    Dim XLF_DA As Double, XLF_DESCNF As Double, Xvalor As Double, XLF_VLBRUTO As Double
'    Dim XLF_VALORGLOSA As Double
'    Dim XLO_DESC As Object
'    Dim XLO_SUBRELATORIO As ADODB.Recordset
'    Dim XLO_SUBRELATORIO2 As ADODB.Recordset
'    Dim XLO_SUBRELATORIO3 As ADODB.Recordset
'    Dim XLO_SUBRELATORIO4 As ADODB.Recordset
'
'    TelaNotasFiscais.MousePointer = vbHourglass
'
'    'Preenche os parâmetros a serem passados para o relatório
'    XRecibo = CStr(NumRecibo)
'    Xvalor = CStr(TDBGrid1.Columns("Valor"))
'    XLF_VLBRUTO = CStr(TDBGrid1.Columns("Valor"))
'    XExtenso = funExtenso(TDBGrid1.Columns("Valor"))
'    Xhistorico = TDBGrid1.Columns("Histórico")
'    XCidade = "Salvador"
'    XFornecedor = Adodc1.Recordset.Fields("focl_tx_RazaoSocial")
'    XLT_TEMDESC = "N"
'
'    XLT_SQL = "SELECT dupl_dt_PagRec FROM DUPLICATAS WHERE dupl_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_Fatura") & " and Not (dupl_dt_pagrec) IS NULL ORDER BY dupl_dt_pagrec DESC"
'    SubQOpenRecordset XLO_SUBRELATORIO4, XLT_SQL, Estatico
'    If Not XLO_SUBRELATORIO4.EOF Then
'        XLT_DATAPAGREC = Format(XLO_SUBRELATORIO4!dupl_dt_pagrec, "dd/MM/yyyy")
'    End If
'    XLO_SUBRELATORIO4.Close
'
'    If Not IsNull(Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")) Then
'        SubQOpenRecordset XLO_SUBRELATORIO4, "SELECT * FROM ConsFINNotasFiscais WHERE nofi_cd_NotaFiscal = " & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev") & "", Estatico
'        If Not XLO_SUBRELATORIO4.EOF Then
'            XLF_VALORGLOSA = XLO_SUBRELATORIO4!nofi_vl_valor
'        End If
'        XLO_SUBRELATORIO4.Close
'    End If
'
'    If Adodc1.Recordset.Fields("nofi_tx_TipoMov") = "S" Then
'       'DESC/ACRES
'       SubQOpenRecordset XLO_DESC, "Select (sum(Desconto) + sum(Acrescimo)) as DA FROM ConsFINRelFinanceiroDescAcresc WHERE dupl_nr_fatura= " & Adodc1.Recordset.Fields("nofi_nr_fatura"), Estatico
'       XLF_DA = XLO_DESC!DA
'       If XLF_DA <> 0 Then XLT_TEMDESC = "S"
'       XLO_DESC.Close
'       'DESCONTOS NF
'       SubQOpenRecordset XLO_DESC, "Select sum(deen_vl_valor) as Descontos FROM ConsFINDescontosNFRecibo WHERE deen_nr_fatura= " & Adodc1.Recordset.Fields("nofi_nr_fatura"), Estatico
'       If Not XLO_DESC.EOF And Not IsNull(XLO_DESC!Descontos) Then
'          XLF_DESCNF = XLO_DESC!Descontos
'          XLT_TEMDESC = "S"
'       End If
'       XLO_DESC.Close
'       Xvalor = CStr(Format(CDbl(TDBGrid1.Columns("Valor")) + XLF_DA - XLF_DESCNF, "standard"))
'       XExtenso = funExtenso(CDbl(Xvalor))
'
'       Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINNotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & "")
'       Set XLO_SUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINRelFinanceiroDescAcresc")
'       Set XLO_SUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDescontosNFRecibo")
'       Set XLO_SUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque")
'
'       'Abre o relatório
'       FrmVisRelatoriosEXT.Caption = "Recibo (Contas a Pagar)"
'       Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ReciboDespesa.rpt")
'
'      'Passa o recordset para o relatório
'       XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
'
'       XFT_TEXTO = FunTiraAspas(XFT_TEXTO)
'       'Ativa os Relatórios
'       With XGR_RELATORIO.FormulaFields
'         .GetItemByName("Recibo").Text = Chr(34) + XRecibo + Chr(34)
'         .GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
'         .GetItemByName("Texto1").Text = Chr(34) + Mid(XFT_TEXTO, 1, 254) + Chr(34)
'         If Len(XFT_TEXTO) > 254 Then
'            .GetItemByName("Texto2").Text = Chr(34) + Mid(XFT_TEXTO, 255, (Len(XFT_TEXTO) - 254)) + Chr(34)
'         End If
'         .GetItemByName("Valor").Text = Chr(34) + CStr(Format(Xvalor, "standard")) + Chr(34)
'         .GetItemByName("Glosa").Text = Chr(34) + CStr(XLF_VALORGLOSA) + Chr(34)
'         .GetItemByName("ValorExtenso").Text = Chr(34) + XExtenso + Chr(34)
'         .GetItemByName("Cidade").Text = Chr(34) + XCidade + Chr(34)
'         .GetItemByName("TemDescAcresc").Text = Chr(34) + XLT_TEMDESC + Chr(34)
'         If XLT_DATAPAGREC <> "" Then
'            .GetItemByName("DataPagRec").Text = Chr(34) + Mid(XLT_DATAPAGREC, 1, 2) + " de " + MonthName(CLng(Mid(XLT_DATAPAGREC, 4, 2))) + " de " + Right(XLT_DATAPAGREC, 4) + Chr(34)
'         End If
'       End With
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO, "Desconto"
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "DescontoNF"
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO3, "Distribuicao"
'       FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
'    Else
'       Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINNotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & "")
'       Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas")
'       Set XLO_SUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDuplicatas WHERE dupl_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & "")
'       Set XLO_SUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque")
'       Xvalor = CStr(Format(CDbl(TDBGrid1.Columns("Valor")) - XLF_VALORGLOSA, "standard"))
'       XExtenso = funExtenso(CDbl(Xvalor))
'
'       'Abre o relatório
'       FrmVisRelatoriosEXT.Caption = "Recibo (Contas a Receber)"
'       Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ReciboReceita.rpt")
'
'       'Passa o recordset para o relatório
'       XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
'
'       'Ativa os Relatórios
'       With XGR_RELATORIO.FormulaFields
'         .GetItemByName("Recibo").Text = Chr(34) + XRecibo + Chr(34)
'         .GetItemByName("Empresa").Text = Chr(34) + CStr(PCodEmpresa) + Chr(34)
'         .GetItemByName("Valor").Text = Chr(34) + CStr(Format(Xvalor, "standard")) + Chr(34)
'         .GetItemByName("Glosa").Text = Chr(34) + CStr(XLF_VALORGLOSA) + Chr(34)
'         .GetItemByName("ValorExtenso").Text = Chr(34) + XExtenso + Chr(34)
'         .GetItemByName("Historico").Text = Chr(34) + Xhistorico + Chr(34)
'         .GetItemByName("Cidade").Text = Chr(34) + XCidade + Chr(34)
'         .GetItemByName("Fornecedor").Text = Chr(34) + XFornecedor + Chr(34)
'         If XLT_DATAPAGREC <> "" Then
'            .GetItemByName("DataPagRec").Text = Chr(34) + Mid(XLT_DATAPAGREC, 1, 2) + " de " + MonthName(CLng(Mid(XLT_DATAPAGREC, 4, 2))) + " de " + Right(XLT_DATAPAGREC, 4) + Chr(34)
'         End If
'       End With
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO, "ContaCorrente"
'       FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO3, "Distribuicao"
'       FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
'    End If
'    TelaNotasFiscais.MousePointer = vbDefault
'End Sub

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
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True
'    CmdDuplicata.Enabled = True
'    CmdRecibo.Enabled = True
'    CmdDevGlosa.Enabled = True
    chkTotalFiltro.Visible = True 'alterado *
    
'    If EntSaida = "S" Then
'        CmdRecibo.Enabled = True
'        CmdProcPagto.Enabled = True
'    End If

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
        Botao_Fil.Enabled = False
        chkTotalFiltro.Visible = False
        chkTotalFiltro.Value = 0
        lblTotalFiltro.Caption = ""
    Else
        Botao_Fil.Enabled = True
        chkTotalFiltro.Visible = True
        chkTotalFiltro_Click
    End If
    
'    If Adodc1.Recordset.Fields("tido_tx_Descricao") = "NOTA PROMISSÓRIA" Then
    If TDBGrid1.Columns("Tipo Doc.") = "NOTA PROMISSÓRIA" Then
        CmdExcluir.Enabled = False
'        CmdEstorno.Enabled = False
'        cmdCancela.Enabled = False
        CmdAlteraDistribuicao.Enabled = False
'        CmdDevGlosa.Enabled = False
'        CmdPPSequencia.Enabled = False
'        CmdProcPagto.Enabled = False
    End If
    
End Sub

Private Sub subDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    Botao_Fil.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
'    CmdDuplicata.Enabled = False
'    CmdRecibo.Enabled = False
'    CmdDevGlosa.Enabled = False
'    If EntSaida = "S" Then
'        CmdProcPagto.Enabled = False
'    End If
End Sub

Private Sub botao_conf_Click()
    Dim XLF_VALORNOTA As Double, XLF_VALORDISTRIB As Double
    Dim Xvalor

    'Função para verificar preenchimento dos campos obrigatórios
    If FunObrigatorioTXT(TxtVldistribuicao, "O preenchimento do campo Valor é obrigatório.") Then Exit Sub
    If FunObrigatorioCBO(CboPcr2, "O preenchimento do campo Centro Custo é obrigatório.") Then Exit Sub
    If FunObrigatorioMSK(MskConta2, "O preenchimento do campo Centro Contábil é obrigatório.") Then Exit Sub

    DatPCR.Recordset.Bookmark = CboPcr2.SelectedItem
    DatConta.Recordset.Bookmark = CboPlanoContas2.SelectedItem
    
    If XDistribuicao = "A" Then
       ValorTotalDistribuicao = ValorTotalDistribuicao - CDbl(TxtVldistribuicao.Text)
    End If
    
    XLF_VALORDISTRIB = (ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text))
    If CDbl(Format(((Adodc1.Recordset.Fields("nofi_vl_Valor") * 100) / 100), "##,##0.00")) < CDbl(Format((((ValorTotalDistribuicao + CDbl(TxtVldistribuicao.Text)) * 100) / 100), "##,##0.00")) Then
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
                            TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                        End If
                    Else
                        TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text
                    End If
                    TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                    TDBGrid3.Columns(8).Text = CboIdentificador.BoundText
                    TDBGrid3.Columns(9).Text = CboIdentificador.Text
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
                        TDBGrid3.Columns(6).Text = TxtVldistribuicao.Text - TDBGrid3.Columns("Devolução").Text
                    End If
                    TDBGrid3.Columns(7).Text = (TxtPercDistribuicao.Text / 100)
                    TDBGrid3.Columns(8).Text = CboIdentificador.BoundText
                    TDBGrid3.Columns(9).Text = CboIdentificador.Text
                    TDBGrid3.Update
                End If
            End If
        End If
    End If
    PanDistribuicao.Enabled = False
    PanDistribuicao.Visible = False
    CmdGravaDistribuicao.Visible = True
    CmdDesisteDistribuicao.Visible = True
    PanAltDistribRecolhimento.Height = 2760
            
    If Vetor3(0, 0) <> Empty Then
        CmdInsereDistribuicao.Enabled = True
        CmdRemoverDistribuicao.Enabled = True
        CmdAlteraDistribuicao.Enabled = True
        
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    Else
        LblValorTotalDistrib.Caption = ""
        ValorTotalDistribuicao = 0
        CmdInsereDistribuicao.Enabled = True
    End If
    
    TDBGrid3.Enabled = True
    
    If XInserir = "I" Or XDistribuido = "N" Then
        If Adodc1.Recordset.Fields("nofi_vl_Valor") = CDbl(LblValorTotalDistrib.Caption) Then
           XDistribuido = "S"
        Else
           XDistribuido = "N"
        End If
    End If
End Sub



Private Sub Botao_Fil_Click()
    subDesabilitaBotoes
    
    Set Formulario = TelaNotasFiscais
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        subContaReg Adodc1
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
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

Private Sub CboHistorico_Change()
    If CboHistorico.Text <> "" And Not IsNull(CboHistorico.SelectedItem) Then
        DatHistorico.Recordset.Bookmark = CboHistorico.SelectedItem
        TxtHistorico.Text = CboHistorico.BoundText
    End If
End Sub

Private Sub CboHistorico_KeyPress(KeyAscii As Integer)
    If Len(CboHistorico) = 50 Then
        If KeyAscii <> 8 And CboHistorico.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub cboLocalizar_Change()
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
'    If Area = 2 Then
'        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
'    End If
End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
'    If cboLocalizar.BoundText <> "" Then
'        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
'    End If
End Sub

Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If cboLocalizar.BoundText <> "" Then
'        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
'    End If
End Sub

'
' - Trata combo Centro de Custos da tela de Estorno
' --------------------------------------------------------------
Private Sub CboPcr_LostFocus()

    Dim XLO_CHECAR As ADODB.Recordset

    If CboPcr.Text <> "" Then
        ' DatContaEstorno.Recordset.Bookmark = CboPcr.SelectedItem
        DatCREstorno.Recordset.Bookmark = CboPcr.SelectedItem
        MskPcr.Text = CboPcr.BoundText
        
'        ' Se informação de vinculação a nível de empresa for verdadeira
'        If pVinculaCrCc = True Then
         
            SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatCREstorno.Recordset.Fields("cere_cd_pcr"), Estatico
            
            If Not (XLO_CHECAR.EOF) Then
               
                If EntSaida = "E" Then
                    subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatCREstorno.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                ElseIf EntSaida = "S" Then
                    subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatCREstorno.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                End If
                                                    
            End If
            
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
            
            ' Alterado em 21/03/22 - Paulo Garcia
            ' Força a (re)carga do CboPlanoContas sempre que haja mudança na opção do centro de custo,
            ' obrigando a escolha de nova conta que esteja associada ao centro de custo escolhido.
            ' Testa se o formulário está carregado, única garantia de que a execução desta subrotina
            ' será em função da opção do usuários
            ' ----------------------------------------------------------------------------------------
            If (CboPcr.Visible = True And CboPcr.Enabled = True) Then
            
                If (XFT_CR <> DatCREstorno.Recordset.Fields("cere_cd_pcr")) Then
                
                    MskConta.Text = ""
                    CboPlanoContas.Text = ""
                    CboPlanoContas.BoundText = 0
                        
                End If
                    
                XFT_CR = DatCREstorno.Recordset.Fields("cere_cd_pcr")
                
            End If
                    
        ' End If
                    
    End If
                
End Sub

'
' - Trata combo Plano de Contas da tela de Estorno
' --------------------------------------------------------------
Private Sub CboPlanoContas_LostFocus()
    If CboPlanoContas.Text <> "" Then
        DatContaEstorno.Recordset.Bookmark = CboPlanoContas.SelectedItem
        If DatContaEstorno.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = MskConta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        MskConta.Text = CboPlanoContas.BoundText
    End If
End Sub

Private Sub CboTipoDoc_Change()
    If CboTipoDoc.Text <> "" Then
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        TxtCodTipoDoc.Text = CboTipoDoc.BoundText
    End If
End Sub

Private Sub chkTotalFiltro_Click()
 Dim total As Double

    total = 0
    If (chkTotalFiltro.Value = 1) Then
        TDBGrid1.MoveFirst
    
        While Not TDBGrid1.EOF
            total = total + CDbl(TDBGrid1.Columns("Valor"))
            TDBGrid1.MoveNext
        Wend
        
        lblTotalFiltro.Caption = "R$ " & Format(CStr("R$ " & total), "##,##0.00")
        TDBGrid1.MoveFirst
    Else
        lblTotalFiltro.Caption = ""
        
    End If
End Sub

Private Sub CmdAlterar_Click()

    Dim XLO_RS As ADODB.Recordset
    Dim RESDUPL As Object
    Dim ResDesc As Object
    Dim ResTipoDoc As Object
    Dim xsigla As String
    Dim XLI_PROVISAO As Long
    
    XGB_AlteraTransfCusto = True
    Chave = Adodc1.Recordset.Fields("nofi_cd_notafiscal")
    NFatura = Adodc1.Recordset.Fields("nofi_nr_fatura")
    
    If Not IsNull(Adodc1.Recordset.Fields("nofi_tx_controle")) Then
        XControleTransf = Adodc1.Recordset.Fields("nofi_tx_controle")
    End If
    
    xChave = Adodc1.Recordset.Bookmark
    
    XLI_PROVISAO = 0
    
    If Not IsNull(Adodc1.Recordset.Fields("Provisao")) Then
        XLI_PROVISAO = Adodc1.Recordset.Fields("Provisao")
    End If
    
    If EntSaida = "E" Then
        If Not FunVerifica_Permissao("TelaNotasFiscais", "5") Then
            Exit Sub
        End If
    Else
        If Not FunVerifica_Permissao("TelaNotasFiscais", "8") Then
           Exit Sub
        End If
    End If
    
    
    'Se não existe fornecedor o registro é uma transferencia bancária ou recolhimentos
    If IsNull(Adodc1.Recordset.Fields("focl_cd_forncli")) Then
        
        SubQOpenRecordset ResTipoDoc, "SELECT TipoDocumentos.tido_tx_sigla FROM TipoDocumentos WHERE tipodocumentos.tido_cd_tipodoc =" & Adodc1.Recordset.Fields("tido_cd_tipodoc"), "Estatico"
        
        If IsNull(ResTipoDoc!tido_tx_sigla) Then
            xsigla = ""
        Else
            xsigla = ResTipoDoc!tido_tx_sigla
        End If
        
        ResTipoDoc.Close
        
'        If xsigla = "TB" Then
'            XFormulario = "TelaNotasFiscais"
'            FrmTransfBancaria.Show 1
'            Set Formulario = TelaNotasFiscais
'            Exit Sub
'        ElseIf (xsigla = "TBCP" Or xsigla = "TBCR") Then
'            If (XLI_PROVISAO > 0) Then
'
'                xChave = XLI_PROVISAO
'                XFormulario = "TelaNotasFiscais"
'
'                XGT_SQL = _
'                "select distinct planocontas.plco_tx_DescContabil  from notasfiscais " & _
'                 "inner join distribuicao on distribuicao.nofi_cd_notafiscal = notasfiscais.nofi_cd_notafiscal " & _
'                 "inner join PlanoContas on planocontas.plco_cd_Conta = distribuicao.plco_cd_Conta " & _
'                 "where " & _
'                 "notasfiscais.empr_cd_empresa = " & CStr(PCodEmpresa) & " AND " & _
'                 "nofi_nr_provisao = " & CStr(XLI_PROVISAO) & " AND " & _
'                 "( " & _
'                 "    plco_tx_DescContabil  = '13. Salário' OR " & _
'                 "    plco_tx_DescContabil  = 'Férias' OR " & _
'                 "    plco_tx_DescContabil  = 'FGTS' OR " & _
'                 "    plco_tx_DescContabil  = 'INSS' " & _
'                 ") "
'
'                SubQOpenRecordset XLO_RS, XGT_SQL, "Estatico"
'
'                If Not XLO_RS.EOF Then
'                    XLO_RS.Close
'                    FrmProvisaoFolhaPagamento.Show 1
'                Else
'                    XLO_RS.Close
'                    frmProvisaoImpostos.Show 1
'                End If
'                Set Formulario = TelaNotasFiscais
'
'            Else
'                XFormulario = "TelaNotasFiscais"
'                frmTransferenciaComCusto.Show 1
'                Set Formulario = TelaNotasFiscais
'            End If
'
'            Exit Sub
'
'        ElseIf xsigla = "TRCC" Then
'
'            If MsgBox("Não é permitido alterar uma Transferência de Custo. Deseja visualizá-la ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
'                XFormulario = "TelaNotasFiscais"
'                FrmTransfCusto.Show 1
'                Set Formulario = TelaNotasFiscais
'                Exit Sub
'            Else
'                Exit Sub
'            End If
'
'        ElseIf xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then
'
'            '****Alterado para permitir alterar a distribuição de um recolhimento*******
'            If Not FunVerifica_Permissao("TelaNotasFiscais", "10") Then
'                Exit Sub
'            Else
'                subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
'                subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
'
'                Sql = "SELECT * FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & ""
'                subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "ValorDistrib", "dist_vl_GlosaDevolucao", "dist_vl_valor", "percentual", "iden_cd_Projeto", "iden_tx_Descricao"), Vetor3, TDBGrid3
'
'                ValorTotalDistribuicao = 0
'                PanAltDistribRecolhimento.Enabled = True
'                LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
'                PanAltDistribRecolhimento.Left = (TDBGrid1.Width - PanAltDistribRecolhimento.Width) / 2
'                PanAltDistribRecolhimento.Top = (TDBGrid1.Height - PanAltDistribRecolhimento.Height) / 2
'                PanAltDistribRecolhimento.Visible = True
'                Exit Sub
'            End If
'            '***************************************
'        End If
    End If
    
    If Adodc1.Recordset.Fields("tido_tx_Sigla") = "GLOS" Or Adodc1.Recordset.Fields("tido_tx_Sigla") = "NFDE" Then
        'Pegar nº do documento da nota de origem
        SubQOpenRecordset ResTipoDoc, "SELECT nofi_nr_documento FROM NotasFiscais WHERE nofi_cd_notafiscal =" & Adodc1.Recordset.Fields("nofi_cd_notaGlosaDev"), "Estatico"
        If EntSaida = "S" Then
            MsgBox "Não é possível alterar esse tipo de documento." + Chr(13) + "Visualize a Glosa através do seu Documento de origem nº " & ResTipoDoc!nofi_nr_documento & ", no Contas a Receber.", vbCritical, "ATENÇÃO"
        Else
            MsgBox "Não é possível alterar esse tipo de documento." + Chr(13) + "Visualize a Devolução através do seu Documento de origem nº " & ResTipoDoc!nofi_nr_documento & ", no Contas a Pagar.", vbCritical, "ATENÇÃO"
        End If
        ResTipoDoc.Close
        Exit Sub
    End If
    
    SubQOpenRecordset RESDUPL, "SELECT Duplicatas.dupl_cd_duplicata FROM Duplicatas where dupl_nr_fatura =" & NFatura, "Estatico"
    If Not (RESDUPL.BOF And RESDUPL.EOF) Then
        ChaveDp = RESDUPL!dupl_cd_Duplicata
    End If
    RESDUPL.Close
    
    SubQOpenRecordset ResDesc, "SELECT DescontosNF.deen_cd_desconto FROM DescontosNF where deen_nr_fatura =" & NFatura, "Estatico"
    If Not (ResDesc.BOF And ResDesc.EOF) Then
        ChaveD = ResDesc!deen_cd_desconto
    End If
    ResDesc.Close
    
    XInserir = "A"
    XFormulario = "TelaNotasFiscais" 'Inclusão
    
 
    
    If Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "3" Then
        FrmFatura.Show 1
        
        If FrmFatura.CARREGAGRID = True Then
            subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            cboLocalizar.BoundText = Chave
            If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
        End If
        
    Else
        
        If EntSaida = "E" Then
            FrmNotasFiscais.Caption = "Alteração de Contas a Receber"
        Else
            FrmNotasFiscais.Caption = "Alteração de Contas a Pagar"
        End If
        
        FrmNotasFiscais.Show 1
        
        'If FrmNotasFiscais.CARREGAGRID = True Then 26/11/09
            subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            cboLocalizar.BoundText = Chave
            If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
        'End If
        
    End If
    
    

End Sub

Private Sub cmdCancela_Click()
    
    Dim XLO_CHECK As ADODB.Recordset
    Dim XLT_SQL As String
    'verifica se existe imposto recolhido - 12/01/2010 - Patrícia
    XLT_SQL = "SELECT * FROM NOTASFISCAIS INNER JOIN DESCONTOSNF ON (NOTASFISCAIS.NOFI_NR_FATURA = DESCONTOSNF.DEEN_NR_FATURA AND NOFI_NR_FATURA = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & ")"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If Not (XLO_CHECK.EOF) Then
        MsgBox "Esta opção não é permitida para Notas com impostos já recolhidos.!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'verifica a mesma permissão do estorno
    If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("nofi_tx_status") = "Compromissado" Then
        Me.panCancelar.Visible = True
        Me.panCancelar.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
        Me.panCancelar.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
        Me.dtpCancelamento.Value = Now
    Else
        MsgBox "Essa operação só é permitida para notas fiscais compromissadas", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
End Sub

Private Sub CmdConfEstorno_Click()
    Dim XLO_NOVO            As ADODB.Recordset
    Dim XLO_DUPLICATA       As ADODB.Recordset
    Dim XLO_CHECK           As ADODB.Recordset
    Dim XLI_LOTE            As Integer
    Dim XLI_CODLANC         As Integer
    Dim XLI_CODLANCATUAL    As Integer
    Dim XLT_SQL             As String
    Dim XLT_NUMDOC          As String
    Dim XLL_NUMFAT          As Long
    
    Dim XLL_CODDUPLICATA    As Long
    Dim XLT_DOCPAGREC       As String
    Dim XLT_TIPOPAGREC      As String
    Dim XLD_PAGREC          As Date
    Dim XLL_CCORRENTE       As Long
    Dim XLT_HISTORICO As String
    
    If Left(Adodc1.Recordset.Fields("nofi_tx_status"), 1) <> "R" Then
        MsgBox "Esta opção só é permitida para registros com status 'Realizado'", vbCritical, "ATENÇÃO"
        PanEstorno.Enabled = False
        PanEstorno.Visible = False
        Exit Sub
    End If
    
    'CHECAR SE O MÊS ESTÁ FECHADO
    If funMesFechado(DtpEntrada.Value) <> "A" Then
        MsgBox "Não é possível realizar a operação nessa data, pois o mês contábil já foi fechado. " _
            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    'Pegar dados da duplicata de origem
    XLT_SQL = "SELECT * FROM Duplicatas " & _
                    " WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & ""
    SubQOpenRecordset XLO_DUPLICATA, XLT_SQL, Estatico
        
    XLL_CODDUPLICATA = XLO_DUPLICATA!dupl_cd_Duplicata
    XLT_DOCPAGREC = XLO_DUPLICATA!dupl_tx_docpagrec
    XLT_TIPOPAGREC = XLO_DUPLICATA!dupl_tx_tipopagrec
    XLD_PAGREC = XLO_DUPLICATA!dupl_dt_pagrec
    XLL_CCORRENTE = XLO_DUPLICATA!coco_cd_codigo
        
    XLO_DUPLICATA.Close
    
    On Error GoTo TrataErro
        
    Conexao.BeginTrans

    If XLT_ESTORNO = "I" Then
        DatContaEstorno.Recordset.Bookmark = CboPlanoContas.SelectedItem
        DatCREstorno.Recordset.Bookmark = CboPcr.SelectedItem
        XLT_SQL = "SELECT lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa FROM Lancamentos " & _
                    "WHERE lanc_dt_data = " & FunNuloData(DtpEntrada.Value, NomeSgbd) & " AND " & _
                        "lanc_tx_origem LIKE 'F%' AND " & _
                        "empr_cd_empresa = " & PCodEmpresa & _
                     " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        If XLO_CHECK.EOF Then
            XLI_LOTE = (FunDefinirUltimoLote(DtpEntrada.Value))
        Else
            XLI_LOTE = XLO_CHECK!lanc_nr_Lote
        End If
        XLO_CHECK.Close
                
        '**************CONTABILIDADE*****************
        'ABRIR TABELA DE LANÇAMENTOS
        XLT_SQL = "SELECT * FROM ConsFinEstornoFinanceiro " & _
                    " WHERE lanc_cd_origemfin = " & XLL_CODDUPLICATA & " " & _
                    " AND lanc_tx_Origem = 'FD'"
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    
        While Not XLO_CHECK.EOF
    '            '****************************************************
    '            'Este trecho de código serve para tratar casos onde o valor seja distribuido
    '            'por CRs diferentes, mas com valores iguais
    '            XLI_CODLANCATUAL = XLO_CHECK!lanc_cd_lanc
    '            If XLO_CHECK.Bookmark > 1 Then
    '                XLO_CHECK.MovePrevious
    '                If XLI_CODLANCATUAL = XLO_CHECK!lanc_cd_lanc Then
    '                    XLO_CHECK.MoveNext
    '                End If
    '                If Not XLO_CHECK.EOF Then XLO_CHECK.MoveNext
    '            End If
    '            '****************************************************
            
            If Not XLO_CHECK.EOF Then
            XLI_CODLANC = funDefinirCodigoLanc(DtpEntrada.Value, XLI_LOTE)
            If EntSaida = "E" Then
                'Se for cliente, Débito de cliente e Crédito da conta informada
                'Se for entrada, Débito da receita errada e Crédito da conta informada
    
                If Not IsNull(XLO_CHECK!plco_cd_credito) Then ' Se for nulo é porque é apenas a parte do banco - Não interessa nesse momento!
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                        "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                        "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                        "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                                " VALUES (" & FunNuloData(DtpEntrada.Value, NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                        FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunNuloBanco(DatContaEstorno.Recordset.Fields("plco_cd_conta")) & "," & FunNuloBanco(DatCREstorno.Recordset.Fields("cere_cd_Pcr")) & "," & FunTrataFloat(XLO_CHECK!lanc_vl_valor) & "," & _
                                        "'" & CboHistorico.Text & "','" & CboHistorico.Text & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & XLL_CODDUPLICATA & "," & _
                                        "'" & XLO_CHECK!lanc_tx_Origem & "', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
                    Conexao.Execute XLT_SQL
                End If
            Else
                'Se for fornecedor, Crédito de fornecedor e Débito da conta informada
                'Se for saída, Crédito da despesa e Débito da conta informada
                If Not IsNull(XLO_CHECK!plco_cd_Debito) Then ' Se for nulo é porque é apenas a parte do banco - Não interessa nesse momento!
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                        "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                        "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                        "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                             " VALUES (" & FunNuloData(DtpEntrada.Value, NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                        FunNuloBanco(DatContaEstorno.Recordset.Fields("plco_cd_conta")) & ", " & FunNuloBanco(DatCREstorno.Recordset.Fields("cere_cd_Pcr")) & " ," & FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunTrataFloat(XLO_CHECK!lanc_vl_valor) & "," & _
                                        "'" & CboHistorico.Text & "','" & CboHistorico.Text & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & XLL_CODDUPLICATA & "," & _
                                        "'" & XLO_CHECK!lanc_tx_Origem & "', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
                    Conexao.Execute XLT_SQL
                End If
            End If
            XLO_CHECK.MoveNext
            End If
        Wend
        
        'ATUALIZAR O DOCUMENTO DEVIDO
        If TxtProcPagto.Text <> "" Then
            Conexao.Execute "UPDATE Duplicatas SET dupl_tx_tipopagrec='" & XLT_TIPOPAGREC & "', " & _
                  "dupl_dt_pagrec =" & FunNuloData(XLD_PAGREC, NomeSgbd) & ", " & _
                  "dupl_tx_docpagrec = '" & XLT_DOCPAGREC & "', dupl_tx_status='R', " & _
                  "coco_cd_codigo=" & XLL_CCORRENTE & " WHERE dupl_nr_fatura=" & TxtProcPagto.Text
        
            Conexao.Execute "UPDATE NotasFiscais SET nofi_tx_status = 'R' " & _
                  " WHERE nofi_nr_fatura=" & TxtProcPagto.Text
        End If
        
        'ATUALIZAR O DOCUEMNTO ORIGINAL
    
        XLT_NUMDOC = "E" & Mid(Adodc1.Recordset.Fields("nofi_nr_documento"), 2, (Len(Adodc1.Recordset.Fields("nofi_nr_documento")) - 1))
        XLT_HISTORICO = "(ESTORNADO)" & Left(Adodc1.Recordset.Fields("nofi_tx_historico"), 89)
        Conexao.Execute "UPDATE NotasFiscais SET nofi_nr_documento='" & XLT_NUMDOC & "', " & _
              " tido_cd_tipodoc =" & CboTipoDoc.BoundText & ", " & _
              " nofi_tx_historico = '" & XLT_HISTORICO & "' " & _
              " WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
        
        Set XLO_CHECK = Nothing
        Set XLO_DUPLICATA = Nothing
        Set XLO_NOVO = Nothing
        
        PanEstorno.Enabled = False
        PanEstorno.Visible = False
        
        
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        
        ReDim XGM_MATRIZLOG(5, 2) As Variant
        XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
        XGM_MATRIZLOG(1, 0) = "Valor" 'Valor da nota fiscal.
        XGM_MATRIZLOG(2, 0) = "PP original" 'Valor da nota fiscal.
        XGM_MATRIZLOG(3, 0) = "PP novo" 'Valor da nota fiscal.
        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("nofi_nr_documento")
        XGM_MATRIZLOG(1, 1) = Format$(Adodc1.Recordset.Fields("nofi_vl_valor"), "##,##0.00")
        XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("nofi_nr_fatura")
        XGM_MATRIZLOG(3, 1) = FunNulo(TxtProcPagto)
        
        Call subRegistraLog("TelaNotasFiscais", "6", funCriaDescricaoLog(XGM_MATRIZLOG, ESTORNO, NOTAFISCAL_PAGAR))
        
        '*******************************************************************

        
    Else
        
        'CANCELAR O PAGTO DO PP DEVIDO
        'Atualiza a tabela de duplicatas
        Sql = "UPDATE Duplicatas SET dupl_tx_docpagrec=null,dupl_dt_pagrec=null,dupl_tx_status='C' WHERE dupl_nr_fatura=" & TxtProcPagto.Text
        Conexao.Execute Sql
        
        'Atualiza a tabela de notas fiscais
        Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_nr_fatura=" & TxtProcPagto.Text
        Conexao.Execute Sql
        
        'ATUALIZAR O DOCUMENTO ORIGINAL

        XLT_NUMDOC = "A" & Mid(Adodc1.Recordset.Fields("nofi_nr_documento"), 2, (Len(Adodc1.Recordset.Fields("nofi_nr_documento")) - 1))
        XLT_HISTORICO = Mid(Adodc1.Recordset.Fields("nofi_tx_historico"), 12, 88)
        Conexao.Execute "UPDATE NotasFiscais SET nofi_nr_documento='" & XLT_NUMDOC & "', " & _
              " tido_cd_tipodoc =" & CboTipoDoc.BoundText & ", " & _
              " nofi_tx_historico = '" & XLT_HISTORICO & "' " & _
              " WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
        
        
        'EXCLUIR O ESTORNO DA CONTABILIDADE
        '***************CONTABILIDADE*******************
        Conexao.Execute "DELETE FROM Lancamentos " & _
                        " WHERE lanc_cd_OrigemFin= " & XLL_CODDUPLICATA & "" & _
                        " AND lanc_tx_Origem = 'FD'" & _
                        " AND lanc_dt_data = " & DtpEntrada.Value & "" & _
                        " AND empr_cd_empresa = " & PCodEmpresa & ""
        '***********************************************
                
        Set XLO_CHECK = Nothing
        Set XLO_DUPLICATA = Nothing
        Set XLO_NOVO = Nothing
        
        PanEstorno.Enabled = False
        PanEstorno.Visible = False
        
                
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        
        ReDim XGM_MATRIZLOG(4, 2) As Variant
        XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
        XGM_MATRIZLOG(1, 0) = "Valor" 'Valor da nota fiscal.
        XGM_MATRIZLOG(2, 0) = "PP original" 'Número do PP original.
        XGM_MATRIZLOG(3, 0) = "PP novo" 'Número do novo PP.
        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("nofi_nr_documento")
        XGM_MATRIZLOG(1, 1) = Format$(Adodc1.Recordset.Fields("nofi_vl_valor"), "##,##0.00")
        XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("nofi_nr_fatura")
        XGM_MATRIZLOG(3, 1) = FunNulo(TxtProcPagto)
        
        Call subRegistraLog("TelaNotasFiscais", "6", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR_ESTORNO, NOTAFISCAL_PAGAR))
        
        '*******************************************************************
        
        
    End If
    Conexao.CommitTrans
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    If Chave <> -1 Then
        cboLocalizar.BoundText = Chave
        If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
    Exit Sub
    
TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub
End Sub


Private Sub CmdConfEstornoANTIGO_Click()
    Dim XLO_NOVO As ADODB.Recordset
    Dim XLO_DUPLICATA As ADODB.Recordset
    Dim XLO_CHECK As ADODB.Recordset
    Dim XLI_LOTE As Integer
    Dim XLI_CODLANC As Integer
    Dim XLI_CODLANCATUAL As Integer
    Dim XLT_SQL As String
    Dim XLT_NUMDOC As String
    Dim XLL_NUMFAT As Long
    Dim XLT_DESCRICAOLOG As String
    
    If Left(Adodc1.Recordset.Fields("nofi_tx_status"), 1) <> "R" Then
        MsgBox "Esta opção só é permitida para registros com status 'Realizado'", vbCritical, "ATENÇÃO"
        PanEstorno.Enabled = False
        PanEstorno.Visible = False
        Exit Sub
    End If
    
    'CHECAR SE O MÊS ESTÁ FECHADO
    If funMesFechado(DtpEntrada.Value) <> "A" Then
        MsgBox "Não é possível realizar a operação nessa data, pois o mês contábil já foi fechado. " _
            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    DatContaEstorno.Recordset.Bookmark = CboPlanoContas.SelectedItem
    XLT_SQL = "SELECT lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa FROM Lancamentos " & _
                "WHERE lanc_dt_data = " & FunNuloData(DtpEntrada.Value, NomeSgbd) & " AND " & _
                    "lanc_tx_origem LIKE 'F%' AND " & _
                    "empr_cd_empresa = " & PCodEmpresa & _
                 " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If XLO_CHECK.EOF Then
        XLI_LOTE = (FunDefinirUltimoLote(DtpEntrada.Value))
    Else
        XLI_LOTE = XLO_CHECK!lanc_nr_Lote
    End If
    XLO_CHECK.Close
    
    On Error GoTo TrataErro
    
    Conexao.BeginTrans

    'GERANDO O NOVO REGISTRO DA NOTA FISCAL
    XLT_SQL = "SELECT * FROM Notasfiscais WHERE nofi_cd_notafiscal = -1"
    SubQOpenRecordset XLO_NOVO, XLT_SQL, Dinamico

    XLO_NOVO.AddNew

    XLT_NUMDOC = Format(FunPegaGuardaUltimo("D"), "E000000000")
    XLL_NUMFAT = FunPegaGuardaUltimo("F")

    XLO_NOVO!nofi_nr_documento = XLT_NUMDOC
    XLO_NOVO!nofi_dt_emissao = Format$(DtpEntrada.Value, "DD/MM/YY")
    XLO_NOVO!nofi_dt_entrada = Format$(DtpEntrada.Value, "DD/MM/YY")
    XLO_NOVO!nofi_tx_tipomov = Adodc1.Recordset.Fields("nofi_tx_tipomov")
    XLO_NOVO!nofi_tx_status = Left(Adodc1.Recordset.Fields("nofi_tx_status"), 1)
    XLO_NOVO!focl_cd_forncli = Adodc1.Recordset.Fields("focl_cd_forncli")
    XLO_NOVO!nofi_tx_serie = Adodc1.Recordset.Fields("nofi_tx_serie")
    XLO_NOVO!tido_cd_TipoDoc = CboTipoDoc.BoundText
    XLO_NOVO!nofi_tx_historico = FunNulo(CboHistorico.Text)
    XLO_NOVO!nofi_vl_valor = FunNuloVal(Adodc1.Recordset.Fields("nofi_vl_valor")) * -1
    XLO_NOVO!nofi_tx_distribuir = Adodc1.Recordset.Fields("nofi_tx_distribuir")
    XLO_NOVO!nofi_nr_fatura = XLL_NUMFAT
    XLO_NOVO!nofi_tx_tipofatura = Adodc1.Recordset.Fields("nofi_tx_tipofatura")
    XLO_NOVO!empr_cd_empresa = PCodEmpresa

    XLO_NOVO.Update

    Chave = XLO_NOVO!nofi_cd_notafiscal

    XLO_NOVO.Close
    
    'Distribuição
    XLT_SQL = "SELECT * FROM Distribuicao WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_notafiscal") & ""
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico

    While Not XLO_CHECK.EOF
        Sql = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor," & _
                    "dist_vl_GlosaDevolucao,nofi_cd_notafiscal, iden_cd_Projeto)" & _
              " VALUES ('" & XLO_CHECK!cere_cd_Pcr & "','" & XLO_CHECK!plco_cd_conta & "'," & FunNuloVal(FunTrataFloat(XLO_CHECK!dist_vl_valor)) & "," & _
              "" & FunNuloVal(FunTrataFloat(XLO_CHECK!dist_vl_GlosaDevolucao)) & "," & Chave & ", " & FunNuloBancoVal(XLO_CHECK!iden_cd_Projeto) & ")"
        Conexao.Execute Sql

        XLO_CHECK.MoveNext
    Wend
    XLO_CHECK.Close

    'ABRIR AS DUPLICATAS DA NOTA
    XLT_SQL = "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & ""
    SubQOpenRecordset XLO_DUPLICATA, XLT_SQL, Estatico
    
    While Not XLO_DUPLICATA.EOF
        'NOVO REGISTRO DE DUPLICATA
        SubQOpenRecordset XLO_NOVO, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
        XLO_NOVO.AddNew
        
        XLO_NOVO!dupl_nr_duplicata = XLO_DUPLICATA!dupl_nr_duplicata
        XLO_NOVO!dupl_dt_vencimento = DtpEntrada.Value
        XLO_NOVO!dupl_vl_GlosaDevolucao = XLO_DUPLICATA!dupl_vl_GlosaDevolucao
        XLO_NOVO!dupl_vl_valor = (XLO_DUPLICATA!dupl_vl_valor * -1)
        XLO_NOVO!dupl_tx_status = XLO_DUPLICATA!dupl_tx_status
        XLO_NOVO!dupl_tx_tipopagrec = XLO_DUPLICATA!dupl_tx_tipopagrec
        XLO_NOVO!dupl_tx_docpagrec = FunNulo(XLO_DUPLICATA!dupl_tx_docpagrec)
        XLO_NOVO!coco_cd_codigo = FunNuloBanco(XLO_DUPLICATA!coco_cd_codigo)
        XLO_NOVO!dupl_tx_debcred = XLO_DUPLICATA!dupl_tx_debcred
        XLO_NOVO!dupl_nr_fatura = XLL_NUMFAT
        XLO_NOVO!dupl_dt_pagrec = DtpEntrada.Value
        
        XLO_NOVO.Update

        ChaveD = XLO_NOVO!dupl_cd_Duplicata

        XLO_NOVO.Close
        
        '**************CONTABILIDADE*****************
        'ABRIR TABELA DE LANÇAMENTOS
        XLT_SQL = "SELECT * FROM ConsFinEstornoFinanceiro " & _
                    " WHERE lanc_cd_origemfin = " & XLO_DUPLICATA!dupl_cd_Duplicata & " " & _
                    " AND lanc_tx_Origem = 'FD'"
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico

        While Not XLO_CHECK.EOF
            '****************************************************
            'Este trecho de código serve para tratar casos onde o valor seja distribuido
            'por CRs diferentes, mas com valores iguais
            XLI_CODLANCATUAL = XLO_CHECK!lanc_cd_Lanc
            If XLO_CHECK.Bookmark > 1 Then
                XLO_CHECK.MovePrevious
                If XLI_CODLANCATUAL = XLO_CHECK!lanc_cd_Lanc Then
                    XLO_CHECK.MoveNext
                End If
                If Not XLO_CHECK.EOF Then XLO_CHECK.MoveNext
            End If
            '****************************************************
            
            If Not XLO_CHECK.EOF Then
            XLI_CODLANC = funDefinirCodigoLanc(DtpEntrada.Value, XLI_LOTE)
            If EntSaida = "E" Then
                'Se for cliente, Débito de cliente e Crédito da conta informada
                'Se for entrada, Débito da receita errada e Crédito da conta informada

                If Not IsNull(XLO_CHECK!plco_cd_credito) Then ' Se for nulo é porque é apenas a parte do banco - Não interessa nesse momento!
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                        "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                        "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                        "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                                " VALUES (" & FunNuloData(DtpEntrada.Value, NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                        "" & FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunNuloBanco(DatContaEstorno.Recordset.Fields("plco_cd_conta")) & "," & FunNuloBanco(XLO_CHECK!cere_cd_Pcr) & "," & FunTrataFloat(XLO_CHECK!lanc_vl_valor) & "," & _
                                        "'" & CboHistorico.Text & "','" & CboHistorico.Text & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & ChaveD & "," & _
                                        "'" & XLO_CHECK!lanc_tx_Origem & "', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
                    Conexao.Execute XLT_SQL
                End If
            Else
                'Se for fornecedor, Crédito de fornecedor e Débito da conta informada
                'Se for saída, Crédito da despesa e Débito da conta informada
                If Not IsNull(XLO_CHECK!plco_cd_Debito) Then ' Se for nulo é porque é apenas a parte do banco - Não interessa nesse momento!
                    XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                        "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                        "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                        "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                             " VALUES (" & FunNuloData(DtpEntrada.Value, NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                        "" & FunNuloBanco(DatContaEstorno.Recordset.Fields("plco_cd_conta")) & "," & FunNuloBanco(XLO_CHECK!cere_cd_Pcr) & "," & FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunTrataFloat(XLO_CHECK!lanc_vl_valor) & "," & _
                                        "'" & CboHistorico.Text & "','" & CboHistorico.Text & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & ChaveD & "," & _
                                        "'" & XLO_CHECK!lanc_tx_Origem & "', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
                    Conexao.Execute XLT_SQL
                End If
            End If
            XLO_CHECK.MoveNext
            End If
        Wend
        XLO_DUPLICATA.MoveNext
    Wend
    XLO_DUPLICATA.Close
    
    Set XLO_CHECK = Nothing
    Set XLO_DUPLICATA = Nothing
    Set XLO_NOVO = Nothing
    
    Conexao.CommitTrans
    
    PanEstorno.Enabled = False
    PanEstorno.Visible = False
    
    'REGISTRO DO LOG
    XLT_DESCRICAOLOG = "Estorno Financeiro do doc " & Adodc1.Recordset.Fields("nofi_nr_documento") & " - " & Adodc1.Recordset.Fields("focl_tx_fantasia")
    Call subGravaLog(func_cd_operador, "TelaNotasFiscais", "6", Now(), XLT_DESCRICAOLOG)
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    If Chave <> -1 Then
        cboLocalizar.BoundText = Chave
        If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
    Exit Sub
    
TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub
End Sub

Private Sub cmdConfirmarCancelamento_Click()

    Dim XLO_CHECK As ADODB.Recordset
    Dim XLI_LOTE As Integer
    Dim XLI_CODLANC As Integer
    Dim XLT_SQL As String
    Dim XLT_HISTORICO As String
    Dim XLT_DESCRICAOLOG As String
    
    'Função que realiza diversas verificações antes do cancelamento
    If FunVerificacoes = False Then
        Exit Sub
    End If
    
    If Left(Adodc1.Recordset.Fields("nofi_tx_status"), 1) <> "C" Then
        MsgBox "Esta opção só é permitida para registros com status 'Compromissado'" + Chr(13) + _
                "Se for 'Realizado', utilize a opção de estorno!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    'CHECAR SE EXISTEM PARCELAS REALIZADAS
    SubQOpenRecordset XLO_CHECK, "SELECT dupl_cd_duplicata FROM Duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & " AND dupl_tx_status = 'R'", "Estatico"
    If Not (XLO_CHECK.BOF And XLO_CHECK.EOF) Then
        MsgBox "A Nota Fiscal não pode ser cancelada pois existem duplicatas com Status = 'Realizado' vinculadas a ela.", vbCritical, "ATENÇÃO"
        XLO_CHECK.Close
        Exit Sub
    End If
    XLO_CHECK.Close
    
    'CHECAR SE O MÊS ESTÁ FECHADO
    If funMesFechado(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy")) <> "A" Then
        MsgBox "Não é possível realizar a operação nessa data, pois o mês contábil já foi fechado. " _
            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    If MsgBox("Deseja cancelar o registro?", vbCritical + vbYesNo, "CUIDADO") = vbYes Then
        XLT_SQL = "SELECT lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa FROM Lancamentos " & _
                    "WHERE lanc_dt_data = " & FunNuloData(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy"), NomeSgbd) & " AND " & _
                        "lanc_tx_origem LIKE 'F%' AND " & _
                        "empr_cd_empresa = " & PCodEmpresa & _
                     " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        If XLO_CHECK.EOF Then
            XLI_LOTE = (FunDefinirUltimoLote(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy")))
        Else
            XLI_LOTE = XLO_CHECK!lanc_nr_Lote
        End If
        XLO_CHECK.Close
        
        On Error GoTo TrataErro
        
        Conexao.BeginTrans
    
        'ALTERANDO O STATUS PARA CANCELADO - X
        Conexao.Execute "UPDATE NotasFiscais SET nofi_tx_Status = 'X' WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura")
        Conexao.Execute "UPDATE Duplicatas SET dupl_tx_Status = 'X' WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura")
                
        'Cadastrar na contabilidade o inverso da provisão.
        'ABRIR TABELA DE LANÇAMENTOS
        XLT_SQL = "SELECT lancamentos.* FROM lancamentos LEFT JOIN notasfiscais ON lancamentos.lanc_cd_origemfin = notasfiscais.nofi_cd_notafiscal " & _
                    "WHERE (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T') AND nofi_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & ""
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        
        If XLO_CHECK.EOF Then
            XLT_SQL = "SELECT lancamentos.lanc_cd_lanc FROM Lancamentos " & _
                        " WHERE lanc_tx_Origem = 'EE' AND " & _
                        " empr_cd_empresa = " & PCodEmpresa & " AND " & _
                        " lanc_cd_Estoque =(SELECT enma_nr_Ordem FROM EntradaMateriais " & _
                        " WHERE  nofi_cd_NotaFiscal = (SELECT min(nofi_cd_NotaFiscal) " & _
                        " FROM NotasFiscais WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & "))"
            
            SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        End If
        
        XLT_HISTORICO = "CANCELAMENTO DO DOC: " & Adodc1.Recordset.Fields("NOFI_NR_DOCUMENTO") & " - " & Adodc1.Recordset.Fields("FOCL_TX_FANTASIA")
        While Not XLO_CHECK.EOF
            XLI_CODLANC = funDefinirCodigoLanc(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy"), XLI_LOTE)
             
            XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                        " VALUES (" & FunNuloData(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy"), NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                "" & FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunNuloBanco(XLO_CHECK!plco_cd_Debito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrDebitar) & "," & FunTrataFloat(XLO_CHECK!lanc_vl_valor) & "," & _
                                "'" & XLT_HISTORICO & "','" & XLT_HISTORICO & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & XLO_CHECK!lanc_cd_OrigemFin & "," & _
                                "'FN', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
            Conexao.Execute XLT_SQL
             
            XLO_CHECK.MoveNext
        Wend
    
        Set XLO_CHECK = Nothing
        
        Conexao.CommitTrans
        
        
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        
        ReDim XGM_MATRIZLOG(4, 2) As Variant
        XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
        XGM_MATRIZLOG(1, 0) = "PP" 'Número de fatura da nota fiscal.
        XGM_MATRIZLOG(2, 0) = "Valor" 'Valor da nota fiscal.
        XGM_MATRIZLOG(3, 0) = "Historico" 'Histórico da nota fiscal.
        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("nofi_nr_documento")
        XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("nofi_nr_fatura")
        XGM_MATRIZLOG(2, 1) = Format$(Adodc1.Recordset.Fields("nofi_vl_valor"), "##,##0.00")
        XGM_MATRIZLOG(3, 1) = FunNulo(Adodc1.Recordset.Fields("nofi_tx_historico"))
        
        If EntSaida = "S" Then Call subRegistraLog("TelaNotasFiscais", "6", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR, NOTAFISCAL_PAGAR))
        If EntSaida = "E" Then Call subRegistraLog("TelaNotasFiscais", "6", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR, NOTAFISCAL_RECEBER))
        '*******************************************************************

        
        MsgBox ("Cancelamento realizado c/ sucesso !")
        
        Me.panCancelar.Visible = False
        
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    End If
    Exit Sub
    
TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub
   
    
End Sub

Private Sub CmdConfirmaRecibo_Click()
    Dim ResTipoDoc As Object
        
    XFT_TEXTO = TxtTextoRecibo.Text
        
    SubQOpenRecordset ResTipoDoc, "SELECT nofi_tx_controle FROM NotasFiscais WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_notafiscal"), Dinamico
    If IsNull(ResTipoDoc!nofi_tx_controle) Then            'Emitir o recibo e gera o nº de controle
        If MsgBox("Confirma emissão do Recibo?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
            If Adodc1.Recordset.Fields("nofi_tx_TipoMov") = "S" Then
                ResTipoDoc!nofi_tx_controle = CStr(FunPegaGuardaUltimo("C"))
            Else
                ResTipoDoc!nofi_tx_controle = CStr(FunPegaGuardaUltimo("R"))
            End If
'            subImprimeRecibo (ResTipoDoc!nofi_tx_controle)
          Else
            Exit Sub
        End If
    Else
        If MsgBox("O recibo já foi emitido. Deseja reemitir ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
            'Reemitir sem gerar um novo número de controle
'            subImprimeRecibo (ResTipoDoc!nofi_tx_controle)
        Else
            Exit Sub
        End If
    End If
    ResTipoDoc.Update
    ResTipoDoc.Close
    
    PanTextoRecibo.Visible = False
End Sub




Private Sub CmdDesisteDistribuicao_Click()
    PanAltDistribRecolhimento.Visible = False
    PanAltDistribRecolhimento.Enabled = False
End Sub

Private Sub CmdDevGlosa_Click()
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("nofi_nr_documento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If (IsNull(Adodc1.Recordset.Fields("tido_tx_sigla")) Or _
        (Adodc1.Recordset.Fields("tido_tx_sigla") <> "GLOS" And Adodc1.Recordset.Fields("tido_tx_sigla") <> "NFDE" And Adodc1.Recordset.Fields("tido_tx_sigla") <> "TBCP" And Adodc1.Recordset.Fields("tido_tx_sigla") <> "TBCR")) Then
        Chave = Adodc1.Recordset.Fields("nofi_cd_notafiscal")
        NFatura = Adodc1.Recordset.Fields("nofi_nr_fatura")
        If Not IsNull(Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")) Then
            ChaveGlosaDev = Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")
            XInserir = "A"
        Else
            ChaveGlosaDev = -1
            XInserir = "I"
        End If
        
        If EntSaida = "E" Then
            FrmGlosaDevolucao.Caption = "Cadastro de Glosas"
        End If
        
        XFormulario = "TelaNotasFiscais"
        FrmGlosaDevolucao.Show 1
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        cboLocalizar.BoundText = Chave
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    ElseIf Adodc1.Recordset.Fields("tido_tx_sigla") = "GLOS" Then
        MsgBox "Não é possível cadastrar uma Devolução para esse tipo de documento.", vbCritical, "ATENÇÃO"
        Exit Sub
    ElseIf Adodc1.Recordset.Fields("tido_tx_sigla") = "NFDE" Then
        MsgBox "Não é possível cadastrar uma Glosa para esse tipo de documento.", vbCritical, "ATENÇÃO"
        Exit Sub
    ElseIf Adodc1.Recordset.Fields("tido_tx_sigla") = "TBCP" Or Adodc1.Recordset.Fields("tido_tx_sigla") = "TBCR" Then
        MsgBox "Não é possível cadastrar uma Devolução para esse tipo de documento.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
End Sub

Private Sub CmdDuplicata_Click()
    Dim RESDUPL As Object
    Dim ResTipo As Object
    Dim ResTipoDoc As Object
    Dim XTipo As String
    Dim xsigla As String
        
    Chave = Adodc1.Recordset.Fields("nofi_cd_notafiscal")
    SubQOpenRecordset ResTipo, "SELECT NotasFiscais.nofi_cd_notafiscal,Notasfiscais.tido_cd_tipodoc,tipodocumentos.tido_tx_duplicata FROM Notasfiscais LEFT JOIN TipoDocumentos ON Notasfiscais.tido_cd_tipodoc = Tipodocumentos.tido_cd_tipodoc where NotasFiscais.nofi_cd_notafiscal = " & Chave, "Estatico"
    XTipo = ResTipo!tido_tx_duplicata
    ResTipo.Close
    
    NFatura = Adodc1.Recordset.Fields("nofi_nr_fatura")
    xChave = Adodc1.Recordset.Bookmark

    If IsNull(Adodc1.Recordset.Fields("focl_cd_forncli")) Then
    'Se não existe fornecedor o registro é uma transferencia bancária ou recolhimentos
        SubQOpenRecordset ResTipoDoc, "SELECT TipoDocumentos.tido_tx_sigla FROM TipoDocumentos WHERE tipodocumentos.tido_cd_tipodoc =" & Adodc1.Recordset.Fields("tido_cd_tipodoc"), "Estatico"
        xsigla = ResTipoDoc!tido_tx_sigla
        ResTipoDoc.Close
        If xsigla = "TB" Then
            MsgBox "Não existem duplicatas para uma Transferência Bancária.", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf xsigla = "TRCC" Then
            MsgBox "Não existem duplicatas para uma Transferência de Custo.", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf xsigla = "TBCP" Or xsigla = "TBCP" Then
            MsgBox "Não existem duplicatas para uma Transferência Bancária com Custo", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then
            MsgBox "Registros de recolhimentos não possuem duplicatas.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If

    SubQOpenRecordset RESDUPL, "SELECT Duplicatas.dupl_cd_duplicata FROM Duplicatas where dupl_nr_fatura =" & NFatura, "Estatico"
    If Not (RESDUPL.BOF And RESDUPL.EOF) Then
        ChaveDp = RESDUPL!dupl_cd_Duplicata
    End If
    RESDUPL.Close
    
    XInserir = "A"
    XFormulario = "TelaNotasFiscais"
    If EntSaida = "E" Then
        FrmNotasFiscais.Caption = "Alteração de Contas a Receber"
    Else
        FrmNotasFiscais.Caption = "Alteração de Contas a Pagar"
    End If

    If Adodc1.Recordset.Fields("nofi_tx_status") = "P" Or XTipo = "1" Then
        FrmNotasFiscais.TabNotasFiscais.Tab = 0
        FrmNotasFiscais.TabAuxiliar.Tab = 1
    Else
        FrmNotasFiscais.TabNotasFiscais.Tab = 1
    End If
    FrmNotasFiscais.Show 1
    
    If (FrmNotasFiscais.CARREGAGRID = True) Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        TDBGrid1.Refresh
        Adodc1.Recordset.Bookmark = xChave
    End If
End Sub

Private Sub CmdEstorno_Click()

    If Left(Adodc1.Recordset.Fields("nofi_nr_documento"), 1) = "E" Then
    
        If MsgBox("Deseja cancelar o estorno financeiro?", vbCritical + vbYesNo, "CUIDADO") = vbYes Then
            
            XLT_ESTORNO = "C" 'Cancelar

            TxtCodTipoDoc.Text = Adodc1.Recordset.Fields("tido_cd_tipodoc")
            CboTipoDoc.BoundText = Adodc1.Recordset.Fields("tido_cd_tipodoc")
            TxtHistorico.Text = ""
            CboHistorico.Text = Adodc1.Recordset.Fields("nofi_tx_historico")
            
            subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') ORDER BY tido_tx_descricao", Estatico
            If EntSaida = "E" Then
                subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
            ElseIf EntSaida = "S" Then
                subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
            End If
                        
            TxtHistorico.Enabled = False
            CboHistorico.Enabled = False
            MskConta.Enabled = False
            CboPlanoContas.Enabled = False
            MskPcr.Enabled = False
            CboPcr.Enabled = False
            
            DtpEntrada.Value = Format(Now, "DD/MM/yyyy")
                      
            PanEstorno.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
            PanEstorno.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
            PanEstorno.Enabled = True
            PanEstorno.Visible = True
            
            Exit Sub
            
        End If
        
    Else
    
        'PREPARA INSERÇÃO DO ESTORNO FINANCEIRO
        XLT_ESTORNO = "I" 'Inserir
        
        TxtHistorico.Enabled = True
        CboHistorico.Enabled = True
        MskConta.Enabled = True
        CboPlanoContas.Enabled = True
        MskPcr.Enabled = True
        CboPcr.Enabled = True
            
        TxtCodTipoDoc.Text = ""
        CboTipoDoc.Text = ""
        TxtHistorico.Text = ""
        CboHistorico.Text = "ESTORNO DOC: " & Adodc1.Recordset.Fields("NOFI_NR_DOCUMENTO") & " - " & Adodc1.Recordset.Fields("FOCL_TX_FANTASIA")
        MskConta.Text = ""
        CboPlanoContas.Text = ""
    
        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') ORDER BY tido_tx_descricao", Estatico
        subConectarControleDadosNV DatCREstorno, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
        
        If EntSaida = "E" Then
            subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
            subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        ElseIf EntSaida = "S" Then
            subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
            subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        End If
        
        DtpEntrada.Value = Format(Now, "dd/MM/yyyy")
        
        PanEstorno.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
        PanEstorno.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
        PanEstorno.Enabled = True
        PanEstorno.Visible = True
        
        Exit Sub
        
    End If
    
End Sub


Private Sub CmdEstornoANTIGO_Click()
    If Left(Adodc1.Recordset.Fields("nofi_nr_documento"), 1) = "E" Then
        If MsgBox("Deseja cancelar o estorno financeiro?", vbCritical + vbYesNo, "CUIDADO") = vbYes Then
            Dim XLO_DUPLICATA As ADODB.Recordset
            
            Me.MousePointer = vbHourglass
            
            '***************CONTABILIDADE*******************
            'verificar se possui módulo de contabilidade e se o mês está fechado
            If XGB_PQUALICONT Then
                SubQOpenRecordset XLO_DUPLICATA, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & " and dupl_tx_Status='R'", Estatico
                While Not XLO_DUPLICATA.EOF
                    If funMesFechado(XLO_DUPLICATA!dupl_dt_pagrec) <> "A" Then
                        MsgBox "Não é possível concluir a operação, pois o mês contábil já foi fechado. " _
                            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                        Exit Sub
                    End If
                    XLO_DUPLICATA.MoveNext
                Wend
                XLO_DUPLICATA.Close
            End If
            '***********************************************
    
            'Exclui a parte da contabilidade
            SubQOpenRecordset XLO_DUPLICATA, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & " and dupl_tx_Status='R'", Estatico
            
            While Not XLO_DUPLICATA.EOF
                Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin= " & XLO_DUPLICATA!dupl_cd_Duplicata & "" & _
                            " AND lanc_tx_Origem = 'FD'  " & _
                            " AND empr_cd_empresa = " & PCodEmpresa & ""
                XLO_DUPLICATA.MoveNext
            Wend
            XLO_DUPLICATA.Close
            
            'Excluir a distribuição
            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
            'Exclui as Duplicatas
            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura")
            'Exclui as NotasFiscais
            Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
        
            'Atualiza a tela
            subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            
            Me.MousePointer = vbDefault
        End If
    Else
        'PREPARA INSERÇÃO DO ESTORNO FINANCEIRO
        TxtCodTipoDoc.Text = ""
        CboTipoDoc.Text = ""
        TxtHistorico.Text = ""
        CboHistorico.Text = "ESTORNO DOC: " & Adodc1.Recordset.Fields("NOFI_NR_DOCUMENTO") & " - " & Adodc1.Recordset.Fields("FOCL_TX_FANTASIA")
        MskConta.Text = ""
        CboPlanoContas.Text = ""
    
        subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') ORDER BY tido_tx_descricao", Estatico
        If EntSaida = "E" Then
            subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'R' ORDER BY hist_tx_descricao", Estatico
            subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        ElseIf EntSaida = "S" Then
            subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
            subConectarControleDadosNV DatContaEstorno, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
        End If
        DtpEntrada.Value = Format(Now, "dd/MM/yyyy")
        
        PanEstorno.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
        PanEstorno.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
        PanEstorno.Enabled = True
        PanEstorno.Visible = True
        Exit Sub
    End If
End Sub



Private Sub CmdExcluir_Click()
    Dim opcao As Integer
   Dim Houve_Erro As Integer
   Dim RESDUPL As Object
   Dim ResNF As Object
   Dim xNumFat As Long, XLL_FATGLOSA As Long
   Dim XLI_PROVISAO As Long
   
   Dim ResTipoDoc As Object
   Dim xdarfdam As String, xsigla As String
   Dim XData As String
   Dim Posicao As Variant
   Dim XLT_DESCRICAO As String
   Houve_Erro = False
    
   subDesabilitaBotoes
       
   If EntSaida = "E" Then
        If Not FunVerifica_Permissao("TelaNotasFiscais", "6") Then
            Exit Sub
        End If
   Else
        If Not FunVerifica_Permissao("TelaNotasFiscais", "9") Then
            Exit Sub
        End If
   End If
   
   Posicao = Adodc1.Recordset.Bookmark
    
    If FunVerificaRecolhimento = False Then
        MsgBox "Não é possível excluir a nota selecionada, pois existe(m) imposto(s) recolhido(s) associado(s) a ela. " _
                    + Chr(13) + "Para realizar essa operação cancele o recolhimento da nota.", vbInformation, "ATENÇÃO"
        Exit Sub
    End If
   
   xNumFat = 0
   XLI_PROVISAO = 0
   If Not IsNull(Adodc1.Recordset.Fields("nofi_nr_fatura")) Then
        xNumFat = Adodc1.Recordset.Fields("nofi_nr_fatura")
   End If
   
   If Not IsNull(Adodc1.Recordset.Fields("Provisao")) Then
        XLI_PROVISAO = Adodc1.Recordset.Fields("Provisao")
   End If
   
   '***************CONTABILIDADE*******************
   'verificar se possui módulo de contabilidade e  exporta base provisão
   If XGB_PQUALICONT Then
        If XGT_TIPOEXPORTACAO = "PC" Then
            If funMesFechado(Adodc1.Recordset.Fields("nofi_dt_entrada")) <> "A" Then
                MsgBox "Não é possível excluir a Nota selecionada, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                Exit Sub
            End If
        Else
            SubQOpenRecordset RESDUPL, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & xNumFat & " and dupl_tx_Status='R'", Estatico
            While Not RESDUPL.EOF
                If funMesFechado(RESDUPL!dupl_dt_pagrec) <> "A" Then
                    MsgBox "Não é possível excluir a Nota selecionada, pois existem parcelas na contabilidade cujo mês contábil já foi fechado. " _
                        + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    Exit Sub
                End If
                RESDUPL.MoveNext
            Wend
            RESDUPL.Close
        End If
   End If
   '***********************************************
   
   '*******NÃO PERMITE EXCLUIR UMA NOTA CADASTRADA PELO QUALIFAT**********
   If Adodc1.Recordset.Fields("nofi_tx_OrigemQualiFat") = "S" Then
        MsgBox "Não é permitido excluir uma fatura do QualiFat! Entre em contato com o faturamento.", vbCritical, "ATENÇÃO"
        Exit Sub
   End If
   '**********************************************************************
   
   If Adodc1.Recordset.Fields("Estoque") = "Pendente" Then
        MsgBox "A Nota Fiscal não pode ser excluída pois possui pendências no Estoque!", vbCritical, "ATENÇÃO"
        Exit Sub
   End If
   
   If (XLI_PROVISAO > 0) Then
           
        ExcluiProvisao XLI_PROVISAO
        
        If Houve_Erro = False Then
            Adodc1.Refresh
            MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
            If ContFil <> 0 Then
               subContaReg Adodc1
            End If
        End If
        subHabilitaBotoes
        If FunTabelaVazia(Adodc1) Then
           CmdAlterar.Enabled = False
           CmdExcluir.Enabled = False
           CmdImprimir.Enabled = False
           CmdOrdem.Enabled = False
'           CmdDuplicata.Enabled = False
'           CmdRecibo.Enabled = False
'           CmdProcPagto.Enabled = False
        End If
        
        'Posicionando o cursor no próximo ou no anterior
        If Posicao - 1 <= Adodc1.Recordset.RecordCount And Posicao > 1 Then
            Adodc1.Recordset.Bookmark = Posicao - 1
        Else
            subHabilitaBotoes
        End If
   Else
   
       If Adodc1.Recordset.Fields("nofi_tx_status") <> "R" Or Adodc1.Recordset.Fields("tido_tx_Sigla") = "GLOS" Or Adodc1.Recordset.Fields("tido_tx_Sigla") = "NFDE" _
            Or Adodc1.Recordset.Fields("nofi_tx_contabil") <> "S" Or Adodc1.Recordset.Fields("nofi_tx_tipofatura") <> "3" Then
            'Verifica se existe alguma duplicata REALIZADA, se sim não deve permitir apagar a nota
                If IsNull(Adodc1.Recordset.Fields("tido_tx_Sigla")) Or ((Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TB" Or Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TBCP" Or Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TBCR") _
                    And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "GLOS" And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "NFDE" _
                    And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TRCC" And Adodc1.Recordset.Fields("tido_tx_PrestContas") <> "1") Then
                    SubQOpenRecordset RESDUPL, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & " AND dupl_tx_status = 'R'", "Estatico"
                    If Not (RESDUPL.BOF And RESDUPL.EOF) Then
                        MsgBox "A Nota Fiscal não pode ser excluída pois existem duplicatas com Status = 'Realizado' vinculadas a ela.", vbCritical, "ATENÇÃO"
                        RESDUPL.Close
                        subHabilitaBotoes
                        Exit Sub
                    End If
                    RESDUPL.Close
                End If
                
                If Adodc1.Recordset.Fields("tido_tx_Sigla") = "GLOS" Or Adodc1.Recordset.Fields("tido_tx_Sigla") = "NFDE" Then
                    SubQOpenRecordset ResNF, "SELECT nofi_nr_documento,nofi_nr_fatura FROM NotasFiscais WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev") & "", Estatico
                    opcao = MsgBox("O Documento selecionado é uma Glosa ou Devolução." + Chr(13) + "O Documento origem nº " + ResNF!nofi_nr_documento + " será alterado. Deseja Continuar?", vbCritical + vbYesNo, "CUIDADO")
                    XLL_FATGLOSA = ResNF!nofi_nr_fatura
                    ResNF.Close
                Else
                    opcao = MsgBox("Confirma Remoção do Registro ?", vbCritical + vbYesNo, "CUIDADO")
                End If
                
                If opcao = 6 Then  'Confirma exclusão
                    On Error GoTo RotuloErro
                   
                    Conexao.BeginTrans
                   
                    '***************CONTABILIDADE*******************
                    'Exclui a parte da contabilidade
                    If IsNull(Adodc1.Recordset.Fields("tido_tx_Sigla")) Or (Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TB" And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TRCC" And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TBCP" And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "TBCR") Then
                        Conexao.Execute "DELETE FROM Lancamentos " & _
                                        " WHERE lanc_cd_OrigemFin= " & Adodc1.Recordset.Fields("nofi_cd_NotaFiscal") & "" & _
                                        " AND lanc_tx_Origem = 'FN'  " & _
                                        " AND empr_cd_empresa = " & Adodc1.Recordset.Fields("empr_cd_empresa") & ""
            
                        SubQOpenRecordset RESDUPL, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & xNumFat & " and dupl_tx_Status='R'", Estatico
                        While Not RESDUPL.EOF
                            Conexao.Execute "DELETE FROM Lancamentos " & _
                                        " WHERE lanc_cd_OrigemFin= " & RESDUPL!dupl_cd_Duplicata & "" & _
                                        " AND lanc_tx_Origem = 'FD'  " & _
                                        " AND empr_cd_empresa = " & PCodEmpresa & ""
                            RESDUPL.MoveNext
                        Wend
                    RESDUPL.Close
                    End If
                    '***********************************************
                    xsigla = IIf(IsNull(Adodc1.Recordset.Fields("tido_tx_Sigla")), "", Adodc1.Recordset.Fields("tido_tx_Sigla"))
                               
                    If IsNull(Adodc1.Recordset.Fields("focl_cd_forncli")) _
                    Or xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then   'Se não existe fornecedor o registro é uma transferencia bancária ou um recolhimento
                        If xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then
                            xdarfdam = Right(Adodc1.Recordset.Fields("nofi_nr_documento"), 7)
                            XData = "Null "
                               Sql = "UPDATE DescontosNF " & _
                                      "SET deen_nr_darfdam = ''," & _
                                      "deen_dt_recolhimento = " & XData & "" & _
                                      "WHERE deen_nr_darfDam = '" & xdarfdam & "'"
                                      Conexao.Execute Sql
                            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
                            Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
                            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & Str(xNumFat)
                        ElseIf xsigla = "TB" Or xsigla = "TRCC" Or xsigla = "TBCP" Or xsigla = "TBCR" Then
                            SubQOpenRecordset ResNF, "SELECT * FROM NotasFiscais WHERE nofi_tx_controle = '" & Adodc1.Recordset.Fields("nofi_tx_controle") & "'  and empr_cd_empresa = " & PCodEmpresa & "", "Estatico"
                            ResNF.MoveFirst
                            Do While Not ResNF.EOF
                                '***************CONTABILIDADE*******************
                                'Exclui a parte da contabilidade
                                Conexao.Execute "DELETE FROM Lancamentos " & _
                                                " WHERE lanc_cd_OrigemFin= " & ResNF!nofi_cd_notafiscal & "" & _
                                                " AND lanc_tx_Origem = 'FN'  " & _
                                                " AND empr_cd_empresa = " & PCodEmpresa & ""
                    
                                SubQOpenRecordset RESDUPL, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & ResNF!nofi_nr_fatura & " and dupl_tx_Status='R'", Estatico
                                If Not RESDUPL.EOF Then
                                    Conexao.Execute "DELETE FROM Lancamentos " & _
                                                " WHERE lanc_cd_OrigemFin= " & RESDUPL!dupl_cd_Duplicata & "" & _
                                                " AND lanc_tx_Origem = 'FD'  " & _
                                                " AND empr_cd_empresa = " & PCodEmpresa & ""
                                End If
                                RESDUPL.Close
                                '***********************************************
                                Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & ResNF!nofi_nr_fatura
                                Conexao.Execute "DELETE FROM DescontosNF WHERE deen_nr_fatura = " & ResNF!nofi_nr_fatura
                                                        
                                ResNF.MoveNext
                            Loop
                            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
                            Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_tx_controle='" & Adodc1.Recordset.Fields("nofi_tx_controle") & "' and nofi_cd_notafiscal='" & Adodc1.Recordset.Fields("nofi_cd_notafiscal") & "' and empr_cd_empresa = " & PCodEmpresa & ""
                            ResNF.Close
                        End If
                    Else
                        ' Verifcar se existe uma devolução ou glosa
                        If Not IsNull(Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")) And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "GLOS" And Adodc1.Recordset.Fields("tido_tx_Sigla") <> "NFDE" Then
                            ' Pegar Nº da Fatura da Nota de Glosa para apagar as duplicatas
                            SubQOpenRecordset ResNF, "SELECT nofi_nr_fatura FROM NotasFiscais WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev") & "", Estatico
                            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & ResNF!nofi_nr_fatura
                            
                            ResNF.Close
                            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")
                            Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev")
                        End If
                        If Adodc1.Recordset.Fields("tido_tx_Sigla") = "GLOS" Or Adodc1.Recordset.Fields("tido_tx_Sigla") = "NFDE" Then
                            'NOTA FISCAL
                            SubQOpenRecordset ResNF, "SELECT nofi_cd_notaGlosaDev FROM NotasFiscais WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev") & "", Dinamico
                            While Not ResNF.EOF
                                ResNF!nofi_cd_notaGlosaDev = Null
                                ResNF.Update
                                ResNF.MoveNext
                            Wend
                            ResNF.Close
                            'DUPLICATAS
                            SubQOpenRecordset ResNF, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATGLOSA & "", Dinamico '20/01/09 - Patricia. Troca do parâmetro de Estatico para Dinamico
                            
                            While Not ResNF.EOF
                                ResNF!dupl_vl_valor = ResNF!dupl_vl_valor + ResNF!dupl_vl_GlosaDevolucao
                                ResNF!dupl_vl_GlosaDevolucao = 0
                                ResNF.Update
                                ResNF.MoveNext
                            Wend
                            ResNF.Close
                            'DISTRIBUIÇÃO
                            SubQOpenRecordset ResNF, "SELECT * FROM Distribuicao WHERE nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_NotaGlosaDev") & "", Dinamico '20/01/09 - Patricia. Troca do parâmetro de Estatico para Dinamico
                            While Not ResNF.EOF
                                ResNF!dist_vl_valor = ResNF!dist_vl_valor + ResNF!dist_vl_GlosaDevolucao
                                ResNF!dist_vl_GlosaDevolucao = 0
                                ResNF.Update
                                ResNF.MoveNext
                            Wend
                            ResNF.Close
                        End If
                        Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
                        Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_cd_notafiscal=" & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
                        If xNumFat <> 0 Then
                             Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & Str(xNumFat)
                             Conexao.Execute "DELETE FROM DescontosNF WHERE deen_nr_fatura = " & Str(xNumFat)
                        End If
                    End If
                    
                    
                    '******************REGISTRA LOG DA OPERAÇÃO**************************
                    
                    ReDim XGM_MATRIZLOG(4, 2) As Variant
                    XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
                    XGM_MATRIZLOG(1, 0) = "PP" 'Número de fatura da nota fiscal.
                    XGM_MATRIZLOG(2, 0) = "Valor" 'Valor da nota fiscal.
                    XGM_MATRIZLOG(3, 0) = "Historico" 'Histórico da nota fiscal
                    XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("nofi_nr_documento")
                    XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("nofi_nr_fatura")
                    XGM_MATRIZLOG(2, 1) = Format$(Adodc1.Recordset.Fields("nofi_vl_valor"), "##,##0.00")
                    XGM_MATRIZLOG(3, 1) = FunNulo(Adodc1.Recordset.Fields("nofi_tx_historico"))
                    
                    If EntSaida = "E" Then
                        Call subRegistraLog("TelaNotasFiscais", "6", funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR, NOTAFISCAL_RECEBER))
                    Else
                        Call subRegistraLog("TelaNotasFiscais", "9", funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR, NOTAFISCAL_PAGAR))
                    End If
                    '*******************************************************************
        
                    
                    Conexao.CommitTrans
                    
                    If Houve_Erro = False Then
                        Adodc1.Refresh
                        MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
                        If ContFil <> 0 Then
                           subContaReg Adodc1
                        End If
                    End If
                    subHabilitaBotoes
                    If FunTabelaVazia(Adodc1) Then
                       CmdAlterar.Enabled = False
                       CmdExcluir.Enabled = False
                       CmdImprimir.Enabled = False
                       CmdOrdem.Enabled = False
'                       CmdDuplicata.Enabled = False
'                       CmdRecibo.Enabled = False
'                       CmdProcPagto.Enabled = False
                    End If
                    
                    'Posicionando o cursor no próximo ou no anterior
                    If Posicao > 1 Then
                        Adodc1.Recordset.Bookmark = Posicao - 1
                    Else
                        subHabilitaBotoes
                    End If
            ElseIf Adodc1.Recordset.Fields("nofi_tx_status") = "R" Then
                MsgBox "A Nota Fiscal não pode ser excluída pois seu Status = 'Realizado'!", vbCritical, "ATENÇÃO"
                Exit Sub
            ElseIf Adodc1.Recordset.Fields("nofi_tx_contabil") = "S" Then 'Uma nota não realizada pode ser exportada?
                MsgBox "A Nota Fiscal não pode ser excluída pois já foi exportada para a Contabilidade!", vbCritical, "ATENÇÃO"
                Exit Sub
            ElseIf Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "3" Then
                MsgBox "Para excluir esta Nota Fiscal é preciso excluir a Composição da Fatura Primeiro!", vbCritical, "ATENÇÃO"
                Exit Sub
            End If
        End If
    End If
    
    cboLocalizar.Text = ""
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Resume Next
          
End Sub

Private Sub ExcluiProvisao(XLI_PROVISAO As Long)
   On Error GoTo ExcluiProvisao_Error

    Dim XLO_RS As ADODB.Recordset
    Dim XLT_DUPLICATAS As String
    Dim XLT_NOTASFISCAIS As String
    
    If MsgBox("Confirma Remoção do Registro ?", vbCritical + vbYesNo, "CUIDADO") = vbYes Then   'Confirma exclusão
       
        Conexao.BeginTrans
       
        SubQOpenRecordset XLO_RS, "SELECT dupl_cd_duplicata FROM NotasFiscais INNER JOIN Duplicatas ON dupl_nr_fatura = nofi_nr_fatura WHERE nofi_nr_provisao = " & XLI_PROVISAO & " AND empr_cd_empresa = " & PCodEmpresa & "", Estatico
        While Not XLO_RS.EOF
            XLT_DUPLICATAS = XLT_DUPLICATAS & XLO_RS!dupl_cd_Duplicata & ", "
            XLO_RS.MoveNext
        Wend
        XLO_RS.Close
        XLT_DUPLICATAS = Mid(XLT_DUPLICATAS, 1, Len(XLT_DUPLICATAS) - 2)
        
        SubQOpenRecordset XLO_RS, "SELECT nofi_cd_notafiscal FROM NotasFiscais WHERE nofi_nr_provisao = " & XLI_PROVISAO & " AND empr_cd_empresa = " & PCodEmpresa & "", Estatico
        While Not XLO_RS.EOF
            XLT_NOTASFISCAIS = XLT_NOTASFISCAIS & XLO_RS!nofi_cd_notafiscal & ", "
            XLO_RS.MoveNext
        Wend
        XLO_RS.Close
        XLT_NOTASFISCAIS = Mid(XLT_NOTASFISCAIS, 1, Len(XLT_NOTASFISCAIS) - 2)
        
        Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin IN (" & XLT_NOTASFISCAIS & ") " & _
                            " AND lanc_tx_Origem = 'FN'  " & _
                            " AND empr_cd_empresa = " & CStr(PCodEmpresa) & ""
                            
        Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin IN (" & XLT_DUPLICATAS & ") " & _
                            " AND lanc_tx_Origem = 'FD'  " & _
                            " AND empr_cd_empresa = " & CStr(PCodEmpresa) & ""

        Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal IN (" & XLT_NOTASFISCAIS & ")"
        Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_cd_notafiscal IN (" & XLT_NOTASFISCAIS & ")"
        Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_cd_duplicata IN (" & XLT_DUPLICATAS & ")"
        
        If (XLI_PROVISAO > 0) Then
            Conexao.Execute "DELETE FROM PROVENTOSPROVISAO WHERE nofi_nr_provisao = " & XLI_PROVISAO
            Conexao.Execute "DELETE FROM PROVISAO WHERE nofi_nr_provisao = " & XLI_PROVISAO
            Conexao.Execute "DELETE FROM HistoricoProvisaoImpostosIR WHERE NumeroProvisao = " & XLI_PROVISAO
            Conexao.Execute "DELETE FROM HistoricoProvisaoImpostosISS WHERE NumeroProvisao = " & XLI_PROVISAO
        End If
        
        
        
        Conexao.CommitTrans
    End If
    

   On Error GoTo 0
   Exit Sub

ExcluiProvisao_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExcluiProvisao of Formulário TelaNotasFiscais"
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

Private Sub CmdGravaDistribuicao_Click()
    Dim XLT_SQL As String
    Dim XLO_REGISTRO As Object
    Dim XLT_DESCRICAO As String
                    
    If Format(Adodc1.Recordset.Fields("nofi_vl_Valor"), "standard") <> LblValorTotalDistrib Then
       MsgBox "O total da distribuição é diferente do total da Nota Fiscal", vbCritical, "ATENÇÃO"
       Exit Sub
    End If
               
    Conexao.Execute "DELETE from Distribuicao WHERE nofi_cd_notafiscal = " & Chave
    
    If Vetor3(0, 0) <> Empty Then
        subGravaVetor "Distribuicao", "cere_cd_pcr,plco_cd_conta,dist_vl_GlosaDevolucao,dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal", Array("T", "T", "N", "N", "N", "N"), 5, Array(0, 2, 5, 6, 8), "" & Chave & "", Vetor3, TDBGrid3
    End If

    'Descrição do Log
    'Empr = Código da empresa
    'PP = Nº do Processo Pagto
    'Valor = Valor do Registro
           
    XLT_DESCRICAO = "Empr:" & PCodEmpresa & _
                    " PP:" & Adodc1.Recordset.Fields("nofi_nr_fatura") & _
                    " Valor:" & Adodc1.Recordset.Fields("nofi_vl_valor")

    Call subGravaLog(func_cd_operador, "TelaNotasFiscais", "10", Now(), XLT_DESCRICAO)

    '***************CONTABILIDADE*******************
    If Adodc1.Recordset.Fields("nofi_tx_Status") = "Realizado" Then
        subContabilizaRecolhimento Adodc1.Recordset.Fields("nofi_nr_fatura")
    End If
    '***********************************************
    
    PanAltDistribRecolhimento.Visible = False
    PanAltDistribRecolhimento.Enabled = False
End Sub

Private Sub CmdImprimir_Click()
    If EntSaida = "E" Then
        Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Contas a Receber")
    Else
        Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Contas a Pagar")
    End If
End Sub

Private Sub CmdInserir_Click()
    Dim ResFat As Object
    
    Chave = -1
    ChaveD = -1
    ChaveDp = -1
    NFatura = -1
    XInserir = "I"
    XGB_AlteraTransfCusto = False
    
    If Not FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = True
       CmdExcluir.Enabled = True
       CmdImprimir.Enabled = True
       CmdOrdem.Enabled = True
    End If
    
'    If EntSaida = "E" Then
'        If Not FunVerifica_Permissao("TelaNotasFiscais", "4") Then
'            Exit Sub
'        End If
'
'        FrmNotasFiscais.Caption = "Cadastro de Contas a Receber"
'    Else
'        If Not FunVerifica_Permissao("TelaNotasFiscais", "7") Then
'            Exit Sub
'        End If
'    End If
    
    XFormulario = "TelaNotasFiscais"
    FrmNotasFiscais.Show 1
    
    If FrmNotasFiscais.CARREGAGRID = True Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        If Chave <> -1 Then
            cboLocalizar.BoundText = Chave
            If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
        End If
    End If
    
End Sub

Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
        
    TDBGrid1.SetFocus
    
End Sub

Private Sub CmdPPSequencia_Click()
    Dim ResNF As Object

    'PanPPSequencia.Left = (TDBGrid1.Width - PanPPSequencia.Width) / 2
    'PanPPSequencia.Top = (TDBGrid1.Height - PanPPSequencia.Height) / 2
    'TxtInicio.Text = Adodc1.Recordset.Fields("nofi_nr_fatura")
    'PanPPSequencia.Visible = True
    'TxtInicio.SetFocus
    
    FrmImpPPSequencia.TxtInicio.Text = Adodc1.Recordset.Fields("nofi_nr_fatura")
    FrmImpPPSequencia.Show 1
    
    'Atualiza a tabela de Duplicatas
    If FrmImpPPSequencia.TxtInicio.Text <> "" And FrmImpPPSequencia.TxtFim.Text <> "" Then
        SubQOpenRecordset ResNF, "SELECT dupl_dt_ProcPagamento FROM Duplicatas " & _
                                " LEFT JOIN NotasFiscais ON Duplicatas.dupl_nr_fatura = NotasFiscais.nofi_nr_fatura " & _
                                " WHERE dupl_nr_fatura >= " & CLng(FrmImpPPSequencia.TxtInicio.Text) & " " & _
                                " AND dupl_nr_fatura <= " & CLng(FrmImpPPSequencia.TxtFim.Text) & "" & _
                                " AND dupl_tx_DebCred = 'D' AND empr_cd_empresa = " & PCodEmpresa & "", "Dinamico"
        Unload FrmImpPPSequencia
        ResNF.MoveFirst
        While Not ResNF.EOF
          If IsNull(ResNF!dupl_dt_ProcPagamento) Then
              ResNF!dupl_dt_ProcPagamento = Format(Now, "DD/MM/YY")
          End If
          ResNF.MoveNext
        Wend
        ResNF.Close
    Else
        Sql = "SELECT dupl_dt_ProcPagamento FROM NotasFiscais LEFT OUTER JOIN " & _
                  " Duplicatas ON NotasFiscais.nofi_nr_fatura = Duplicatas.dupl_nr_fatura " & _
                  " WHERE nofi_tx_pendEstoque = 'N' and dupl_dt_ProcPagamento is null" & _
                  " AND empr_cd_empresa = " & PCodEmpresa & " "
        SubQOpenRecordset ResNF, Sql, Dinamico
        While Not ResNF.EOF
          ResNF!dupl_dt_ProcPagamento = Format(Now, "DD/MM/YY")
        
          ResNF.MoveNext
        Wend
        ResNF.Close
    End If
End Sub

Private Sub CmdProcPagto_Click()
  Dim FiltroCCusto As String
  Dim ResNF As Object
  Dim NumFat As Long
  Dim ResProc As Object
  Dim XFatura As Long
  Dim Qtd As Integer, x As Integer
  Dim XDtVenc As String, Xvalor As String
  Dim Xhistorico As String, Sql As String
  Dim XL0_RSSUBRELATORIO As ADODB.Recordset
  Dim XL0_RSSUBRELATORIO2 As ADODB.Recordset
  Dim XL0_RSSUBRELATORIO3 As ADODB.Recordset
  Dim XLD_VALORBRUTO As Double
  Dim XLT_EXISTE As String
  Dim XLT_PARCELA As String
  Dim XLT_DISTRIBUIDO As String
  Dim XLT_TIPODOC As String * 1
  Dim XL0_RSTABTEMP As ADODB.Recordset
  Dim XLT_TIPOPAG As String * 1
  Dim Cont As String ' Armazena a quantidade de parcelas.
  Dim NumChamado As String
    
  '******************************************************************************************
  'Comum a todas as impressões
  TelaNotasFiscais.MousePointer = vbHourglass
  '******************************************************************************************
  Cont = "0"
  
  NumFat = Adodc1.Recordset.Fields("nofi_nr_fatura")
  If Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "4" Then
    MsgBox "O Processo de Pagamento não pode ser emitido pois não existem duplicatas cadastradas.", vbCritical, "ATENÇÃO"
    TelaNotasFiscais.MousePointer = vbDefault
    Exit Sub
  'Solicitado pela C. Andrade para que fosse permitido emitir mesmo que realizado informando esse status
  'ElseIf Adodc1.Recordset.Fields("nofi_tx_status") <> "C" Then
  '  MsgBox "O Status da Nota Fiscal deve ser 'Compromissado'.", vbCritical, "ATENÇÃO"
  '  TelaNotasFiscais.MousePointer = vbDefault
  '  Exit Sub
  End If
  
  FiltroCCusto = ""
        
  'Preparação do Arquivo de Impressão
  If XControle <> "" Then
    Sql = "DELETE FROM ProcPagto_Temp WHERE prte_cd_controle = '" & XControle & "'"
    Conexao.Execute Sql
  End If
  XControle = Trim(CStr(func_cd_operador)) + CStr(Date) + CStr(Time())
        
  'XLD_VALORBRUTO = FunTrataFloat(Adodc1.Recordset.Fields("nofi_vl_valor"))
  XLD_VALORBRUTO = 0
  
  NumChamado = FunNuloVal(Adodc1.Recordset.Fields("nofi_nr_chamado"))
  
  'Inserção dos Registros de Duplicatas
  Sql = "SELECT * FROM ConsFINProcPagto WHERE nofi_nr_fatura = " & NumFat & " order by dupl_nr_duplicata"
  SubQOpenRecordset ResProc, Sql, Estatico
  x = 1
  Qtd = ResProc.RecordCount
  Do While Not ResProc.EOF
    '    If Adodc1.Recordset.Fields("nofi_tx_tipofatura") <> "3" Then
    '        XLD_VALORBRUTO = XLD_VALORBRUTO + ResProc!dupl_vl_valor
    '    Else 'FATURA
    '        XLD_VALORBRUTO = XLD_VALORBRUTO + ResProc!nofi_vl_valor
    '    End If
    XFatura = ResProc!nofi_nr_fatura
    Xhistorico = ResProc!Historico
    Xvalor = ResProc!dupl_vl_valor
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura,prte_nr_parcela, prte_nr_chamado) " _
        & "Values('" & XControle & "','" & FunTrataString(Xhistorico) & "'," & FunTrataFloat(Xvalor) & "," & XFatura & "," & x & ",'" & FunTrataString(NumChamado) & "')"
    Conexao.Execute Sql
    x = x + 1
    XDtVenc = ResProc!dupl_dt_vencimento
    XLT_PARCELA = ResProc!dupl_nr_duplicata
    ResProc.MoveNext
  Loop
    
  ResProc.Close
  Set ResProc = Nothing

  
  'Inserção dos Registros de Desconsto/Acréscimo
  Sql = "SELECT * FROM ConsFINProcPagtoDescAcresc WHERE dupl_nr_fatura = " & NumFat & " and dupl_nr_duplicata = '" & XLT_PARCELA & "'"
  SubQOpenRecordset ResProc, Sql, Estatico
  Do While Not ResProc.EOF
    XFatura = ResProc!dupl_nr_fatura
    Xhistorico = FunNulo(ResProc!deac_tx_descricao)
    'Xvalor = ResProc!dude_vl_descacresc
    Xvalor = FunNulo(ResProc!DA)
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura, prte_nr_chamado) " _
        & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(Xvalor) & "," & XFatura & ",'" & FunTrataString(NumChamado) & "')"
    Conexao.Execute Sql
    ResProc.MoveNext
  Loop
  ResProc.Close
  Set ResProc = Nothing
  
  'Inserção dos Registros de DescontosNF
  Sql = "SELECT * FROM ConsFINProcPagtoDesc WHERE nofi_nr_fatura = " & NumFat
  SubQOpenRecordset ResProc, Sql, Estatico
  Do While Not ResProc.EOF
    'XLD_VALORBRUTO = XLD_VALORBRUTO + ResProc!deen_vl_valor
    XFatura = ResProc!nofi_nr_fatura
    Xhistorico = FunNulo(ResProc!desc_tx_descricao)
    Xvalor = FunNulo(ResProc!deen_vl_valor)
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura, prte_nr_chamado) " _
        & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(Xvalor) & "," & XFatura & ",'" & FunTrataString(NumChamado) & "')"
    Conexao.Execute Sql
    ResProc.MoveNext
  Loop
  ResProc.Close
  Set ResProc = Nothing
  
  'Inserção dos registro de devolução
  Sql = "SELECT * FROM ConsFINNotasFiscais WHERE nofi_cd_notaglosadev = " & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
  SubQOpenRecordset ResProc, Sql, Estatico
  Do While Not ResProc.EOF
    XLD_VALORBRUTO = XLD_VALORBRUTO + ResProc!nofi_vl_valor
    'XFatura = ResProc!nofi_nr_fatura
    Xhistorico = FunNulo(ResProc!nofi_tx_historico)
    Xvalor = FunNulo(ResProc!nofi_vl_valor)
    NumChamado = FunNulo(ResProc!nofi_nr_chamado)
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura,prte_nr_chamado) " _
        & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(Xvalor) & "," & XFatura & ",'" & FunTrataString(NumChamado) & "')"
    Conexao.Execute Sql
    ResProc.MoveNext
  Loop
  ResProc.Close
  Set ResProc = Nothing
  
  
  'Verifica se Foi distribuído
  'SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir FROM NotasFiscais WHERE nofi_nr_Fatura = " & NumFat & "", Estatico
  SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir, sum(nofi_vl_valor) as VALOR FROM NotasFiscais WHERE nofi_nr_Fatura = " & NumFat & " group by  nofi_tx_distribuir", Estatico
  While Not ResNF.EOF
     XLD_VALORBRUTO = XLD_VALORBRUTO + ResNF!VALOR
     ResNF.MoveNext
  Wend
  
  ResNF.MoveFirst
  If Not ResNF.EOF Then
    If ResNF!nofi_tx_distribuir = "S" Or Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "3" Then
        XLT_DISTRIBUIDO = "S"
    Else
        XLT_DISTRIBUIDO = "N"
    End If
 '   XLD_VALORBRUTO = XLD_VALORBRUTO + ResNF!Valor
  End If
  ResNF.Close
  Set ResProc = Nothing
  'Verifica a quantidade de parcelas
  SubQOpenRecordset ResNF, "SELECT dupl_cd_duplicata FROM Duplicatas where dupl_nr_fatura =" & NumFat, "Estatico"
  If Not (ResNF.EOF) Then '31/03/09 - Patrícia
    ResNF.MoveFirst
  End If
  While Not ResNF.EOF
    Cont = CStr(CInt(Cont) + 1)
    ResNF.MoveNext
  Wend
  ResNF.Close
  Set ResProc = Nothing
       
  If Cont = 0 Then '31/03/09 - Patrícia
        MsgBox "Não existem Duplicatas cadastradas para o lançamento selecionado.", vbInformation, "Relatório sem dados"
        FrmRelChEmitidos.MousePointer = vbDefault
        Exit Sub
  End If
  XGT_SELECAO = "ConsFINFluxoCaixa.dupl_nr_fatura = " & NumFat & ""
  
    'Atualiza a tabela de Duplicatas
  SubQOpenRecordset ResNF, "SELECT dupl_dt_ProcPagamento FROM Duplicatas where dupl_nr_fatura = " & NumFat, Dinamico
  If Not (ResNF.EOF) Then
        ResNF.MoveFirst
  End If
  While Not ResNF.EOF
    'If Not IsDate(ResNF!dupl_dt_ProcPagamento = FunNuloData(Now, NomeSgbd)) Then
    If IsNull(ResNF!dupl_dt_ProcPagamento) Then
        XLT_EXISTE = "N" 'O processo foi impresso pela 1ª vez
        ResNF!dupl_dt_ProcPagamento = Format(Now, "DD/MM/YY")
    Else
        XLT_EXISTE = "S" 'O processo já foi impresso anteriormente
    End If
    ResNF.MoveNext
  Wend
  ResNF.Close
  Set ResNF = Nothing

  'Verifica se a data do Processo de pagamento foi preenchida.
  SubQOpenRecordset ResNF, "SELECT dupl_dt_ProcPagamento FROM Duplicatas where dupl_nr_fatura = " & NumFat, Dinamico
  If Not (ResNF.EOF) Then
    ResNF.MoveFirst
  End If
  While Not ResNF.EOF
    'If Not IsDate(ResNF!dupl_dt_ProcPagamento = FunNuloData(Now, NomeSgbd)) Then
    If IsNull(ResNF!dupl_dt_ProcPagamento) Then
            MsgBox "Data do PP inválida!", vbInformation, "Processo de Pagamento"
            ResNF.Close
            Set ResNF = Nothing
            Exit Sub
    End If
    ResNF.MoveNext
  Wend
  ResNF.Close
  Set ResNF = Nothing


  'Abre os recordsets e verificar se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINFluxoCaixa WHERE " & XGT_SELECAO)
  Set XL0_RSSUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ProcPagto_Temp WHERE prte_cd_Controle='" & XControle & "'")
  Set XL0_RSSUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_fatura=" & Adodc1.Recordset.Fields("nofi_nr_Fatura") & "")
  Set XL0_RSSUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDescAcresc")
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    FrmRelChEmitidos.MousePointer = vbDefault
    Exit Sub
  End If
  
  'Abre o relatório
  FrmVisRelatorios.Caption = "Processo de Pagamento"
  
  'Inicio da alteração para imprimir o relatório específico da APAE - Lucas Santiago 20.01.09
  If (UCase(PCliente) = "APAE") Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagto_APAE.rpt")
  'ElseIf (PCliente = "Costa Andrade" And Cont > 1) Then
  ElseIf (UCase(PCliente) = "COSTA ANDRADE") Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagtoCA.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagto.rpt")
  End If

  'Final da alteração da impressão do relatório específico da APAE
  
  'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagto.rpt")
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1

  '********** Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("CCusto").Text = Chr(34) + FiltroCCusto + Chr(34)
        .GetItemByName("QuantParcela").Text = Chr(34) + Cont + Chr(34)
        .GetItemByName("VlBrutoNF1").Text = Chr(34) + CStr(XLD_VALORBRUTO) + Chr(34)
        .GetItemByName("VencParcela").Text = Chr(34) + XDtVenc + Chr(34)
        .GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
        .GetItemByName("JaFoiImpresso").Text = Chr(34) + XLT_EXISTE + Chr(34)
        .GetItemByName("Distribuido").Text = Chr(34) + XLT_DISTRIBUIDO + Chr(34)
        .GetItemByName("NumChamado").Text = Chr(34) + NumChamado + Chr(34)
        If (UCase(PCliente) = "APAE") Then
            .GetItemByName("Emitente").Text = Chr(34) + XGT_USUARIO + Chr(34)
            .GetItemByName("Cargo").Text = Chr(34) + XGT_CARGO + Chr(34)
        End If
    End With
  
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO, "ProcPagto_Temp"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO2, "ConsFINDistribCopiaCheque"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO3, "ConsFINDescAcresc"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  
  '///////Impressão do Anexo do Processo (p/notas distribuidas)//////
  If XLT_DISTRIBUIDO = "S" Then
    Dim XLO_RSRELATORIOANEXO As New ADODB.Recordset
    Dim XLR_RELATORIOANEXO As New CRAXDRT.Report
  
    Set XLO_RSRELATORIOANEXO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_fatura=" & Adodc1.Recordset.Fields("nofi_nr_Fatura") & "")
      
    'Abre o relatório
    FrmVisRelatorios_2.Caption = "Anexo do Processo de Pagamento"
    Set XLR_RELATORIOANEXO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\AnexoProcPagto.rpt")
  
    'Passa o recordset para o relatório
    XLR_RELATORIOANEXO.Database.SetDataSource XLO_RSRELATORIOANEXO, 3, 1
  
    '********** Ativa os Relatórios
    With XLR_RELATORIOANEXO.FormulaFields
        .GetItemByName("NumProcesso2").Text = Chr(34) + CStr(NumFat) + Chr(34)
    End With
    FrmVisRelatorios_2.SubVerRelatorio XLR_RELATORIOANEXO, ""
  End If
  '//////////////////////////////////////////////////////////////////
  
  TelaNotasFiscais.MousePointer = vbDefault
    
'  'Limpa a tabela temporária
'  If XControle <> "" Then
'    Sql = "DELETE FROM ProcPagto_Temp WHERE prte_cd_controle = '" & XControle & "'"
'    Conexao.Execute Sql
'  End If
  XFormulario = "TelaNotasFiscais"
  '*************** Imprimir relação de notas da fatura *******************
  If Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "3" Then
    If MsgBox("Deseja imprimir a relação de Notas Fiscais desta Fatura?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
      Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINNotasFiscais where nofi_nr_fatura = " & NumFat & "")
      Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas")
      
      'Abre o relatório
      FrmVisRelatorios.Caption = "Relação de Notas da Fatura Nº " + CStr(NumFat)
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelNotasProcPagto.rpt")
      
      'Passa o recordset para o relatório
      XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1

      '********** Ativa os Relatórios
      XGR_RELATORIO.FormulaFields.GetItemByName("CodEmpresa").Text = Chr(34) & _
        CStr(PCodEmpresa) & Chr(34)
      FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
      FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
      TelaNotasFiscais.MousePointer = vbDefault
    Else
      Exit Sub
    End If
  End If
End Sub

Private Sub CmdRecibo_Click()
    Dim ResTipoDoc As Object
    Dim ResControle As Object
    Dim codnota As Long
    Dim XControle As Integer
    
    SubQOpenRecordset ResTipoDoc, "SELECT NotasFiscais.nofi_cd_notafiscal, NotasFiscais.nofi_tx_controle, " & _
                " NotasFiscais.tido_cd_tipodoc, tipodocumentos.tido_cd_tipodoc, tipodocumentos.tido_tx_recibo " & _
                " FROM NotasFiscais, TipoDocumentos WHERE NotasFiscais.tido_cd_tipodoc = TipoDocumentos.tido_cd_tipodoc " & _
                " AND notasfiscais.nofi_cd_notafiscal = " & Adodc1.Recordset.Fields("nofi_cd_notafiscal"), "Dinamico"
    If ResTipoDoc!tido_tx_recibo = "1" And Adodc1.Recordset.Fields("nofi_tx_status") <> "P" Then 'And Adodc1.Recordset.Fields("nofi_tx_status") = "R" Then
        If Adodc1.Recordset.Fields("nofi_tx_TipoMov") = "S" Then
            TxtTextoRecibo.Text = "RECEBEMOS DE " + PEmpresa + ", CNPJ: " + XGT_CGC + ", ENDEREÇO: " + PEndereco + ", A IMPORTÂNCIA DE R$ " + TDBGrid1.Columns("Valor") + " (" + Trim(funExtenso(CDbl(TDBGrid1.Columns("Valor")))) + ")" + " REFERENTE A " + Adodc1.Recordset.Fields("nofi_tx_Historico")
            PanTextoRecibo.Left = (TDBGrid1.Width - PanTextoRecibo.Width) / 2
            PanTextoRecibo.Top = (TDBGrid1.Height - PanTextoRecibo.Height) / 2
            PanTextoRecibo.Visible = True
            Exit Sub
        Else
            If IsNull(ResTipoDoc!nofi_tx_controle) Then            'Emitir o recibo e gera o nº de controle
                If MsgBox("Confirma emissão do Recibo?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
                    If Adodc1.Recordset.Fields("nofi_tx_TipoMov") = "S" Then
                        ResTipoDoc!nofi_tx_controle = CStr(FunPegaGuardaUltimo("C"))
                    Else
                        ResTipoDoc!nofi_tx_controle = CStr(FunPegaGuardaUltimo("R"))
                    End If
'                    subImprimeRecibo (ResTipoDoc!nofi_tx_controle)
                  Else
                    Exit Sub
                End If
            Else
                If MsgBox("O recibo já foi emitido. Deseja reemitir ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                    'Reemitir sem gerar um novo número de controle
'                    subImprimeRecibo (ResTipoDoc!nofi_tx_controle)
                Else
                    Exit Sub
                End If
            End If
            ResTipoDoc.Update
        End If
    ElseIf ResTipoDoc!tido_tx_recibo <> "1" Then
        MsgBox "Este tipo de Documento não permite a emissão de Recibo.", vbCritical, "ATENÇÃO"
        'Solicitado pela C. Andrade para imprimir também para um compromissado
        'ElseIf Adodc1.Recordset.Fields("nofi_tx_status") <> "R" Then
    ElseIf Adodc1.Recordset.Fields("nofi_tx_status") = "P" Then
        MsgBox "O recibo não pode ser emitido para um documento com Status Previsto.", vbCritical, "ATENÇÃO"
    End If
    ResTipoDoc.Close
End Sub
    

Private Sub CmdRetEstorno_Click()
    PanEstorno.Visible = False
    PanEstorno.Enabled = False
End Sub

Private Sub cmdRetornarCancelamento_Click()
    
    Me.panCancelar.Visible = False
        
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSairRecibo_Click()
    PanTextoRecibo.Visible = False
End Sub



Private Sub CmdRemoverDistribuicao_Click()
    CmdInsereDistribuicao.Enabled = False
    CmdRemoverDistribuicao.Enabled = False
    CmdAlteraDistribuicao.Enabled = False
   
    TDBGrid3.Enabled = False
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor3, TDBGrid3
    End If

    If Vetor3(0, 0) <> Empty Then
        CmdInsereDistribuicao.Enabled = True
        CmdRemoverDistribuicao.Enabled = True
        CmdAlteraDistribuicao.Enabled = True
        
        LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 4)
        ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
    Else
        LblValorTotalDistrib.Caption = ""
        ValorTotalDistribuicao = 0
        CmdInsereDistribuicao.Enabled = True
    End If
   
    TDBGrid3.Enabled = True
 
    Exit Sub
End Sub



Private Sub CmdAlteraDistribuicao_Click()

    CmdInsereDistribuicao.Enabled = False
    CmdRemoverDistribuicao.Enabled = False
    CmdAlteraDistribuicao.Enabled = False
    
    XFT_CR = TDBGrid3.Columns(1).Text
    
    TDBGrid3.Enabled = False
    
    PanDistribuicao.Enabled = True
    PanDistribuicao.Top = 2280
    PanAltDistribRecolhimento.Height = 3945
    PanDistribuicao.Visible = True
    CmdGravaDistribuicao.Visible = False
    CmdDesisteDistribuicao.Visible = False
    XDistribuicao = "A"
    
    MskPcr2.Text = TDBGrid3.Columns(1).Text
    CboPcr2.BoundText = MskPcr2.Text
    XPcr = CboPcr2.BoundText
    MskConta2.Text = TDBGrid3.Columns(3).Text
    CboPlanoContas2.BoundText = MskConta2.Text
    xcc = MskConta2.Text
    TxtVldistribuicao.Text = Format(CDbl(TDBGrid3.Columns(4).Text), "standard")
    TxtPercDistribuicao.Text = Mid(TDBGrid3.Columns(7).Text, 1, (Len(TDBGrid3.Columns(7).Text) - 1))
    CboIdentificador.BoundText = TDBGrid3.Columns(8).Text
    CboIdentificador.Text = TDBGrid3.Columns(9).Text
    
    If MskPcr2.Enabled = True Then MskPcr2.SetFocus

End Sub



Private Sub CmdInsereDistribuicao_Click()

    TelaNotasFiscais.MousePointer = vbHourglass
    XDistribuicao = "I"
    
    XFT_CR = ""
    
    If Adodc1.Recordset.Fields("nofi_vl_Valor") = LblValorTotalDistrib Then
       MsgBox "As distribuições já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
    Else
        PanDistribuicao.Enabled = True
        PanDistribuicao.Top = 2280
        PanAltDistribRecolhimento.Height = 3945
        PanDistribuicao.Visible = True
        TxtPercDistribuicao.Text = ""
        TxtVldistribuicao.Text = ""
        MskPcr2.PromptInclude = False
        MskPcr2.Text = ""
        MskPcr2.PromptInclude = True
        CboPcr2.Text = ""
        MskConta2.Text = ""
        CboPlanoContas2.Text = ""
        CmdGravaDistribuicao.Visible = False
        CmdDesisteDistribuicao.Visible = False
        
        MskPcr2.SetFocus
        
        CmdInsereDistribuicao.Enabled = False
        CmdRemoverDistribuicao.Enabled = False
        CmdAlteraDistribuicao.Enabled = False
        
    End If
    
    TelaNotasFiscais.MousePointer = vbDefault
    
End Sub


Private Sub CmdSairDistribuicao_Click()

    PanDistribuicao.Visible = False
    PanDistribuicao.Enabled = False
    CmdGravaDistribuicao.Visible = True
    CmdDesisteDistribuicao.Visible = True
    PanAltDistribRecolhimento.Height = 2760
       
    XDistribuicao = ""
    TDBGrid3.Enabled = True
    
    CmdInsereDistribuicao.Enabled = True
    If Vetor3(0, 0) <> Empty Then
        CmdRemoverDistribuicao.Enabled = True
        CmdAlteraDistribuicao.Enabled = True
    End If
    
End Sub

'
' - Trata combo Centro de Custos da tela de Distribuição Recolhimento
' -------------------------------------------------------------------
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
        
'        ' Se informação de vinculação a nível de empresa for verdadeira
'        If pVinculaCrCc = True Then
        
            SubQOpenRecordset XLO_CHECAR, "SELECT cere_cd_Pcr FROM CR_CC where cere_cd_pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr"), Estatico
            
            If Not (XLO_CHECAR.EOF) Then
               
                If EntSaida = "E" Then
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                ElseIf EntSaida = "S" Then
                    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_cd_Conta IN (SELECT plco_cd_Conta FROM CR_CC WHERE cere_cd_Pcr = " & DatPCR.Recordset.Fields("cere_cd_pcr") & ") AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
                End If
                                                    
            End If
            
            XLO_CHECAR.Close
            Set XLO_CHECAR = Nothing
            
            ' Alterado em 21/03/22 - Paulo Garcia
            ' Força a (re)carga do CboPlanoContas sempre que haja mudança na opção do centro de custo,
            ' obrigando a escolha de nova conta que esteja associada ao centro de custo escolhido.
            ' Testa se o formulário está carregado, única garantia de que a execução desta subrotina
            ' será em função da opção do usuários
            ' ----------------------------------------------------------------------------------------
            If (CboPcr2.Visible = True And CboPcr2.Enabled = True) Then
            
                If (XFT_CR <> DatPCR.Recordset.Fields("cere_cd_pcr")) Then
                
                    MskConta2.Text = ""
                    CboPlanoContas2.Text = ""
                    CboPlanoContas2.BoundText = 0
                        
                End If
                    
                XFT_CR = DatPCR.Recordset.Fields("cere_cd_pcr")
                
            End If
        
        ' End If
                    
    End If

End Sub

'
' - Trata combo Plano de Contas da tela de Distribuição Recolhimento
' ------------------------------------------------------------------
Private Sub CboPlanoContas2_Change()
    If CboPlanoContas2.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas2.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas2.BoundText = MskConta2.Text
            CboPlanoContas2.SetFocus
            Exit Sub
        End If
        MskConta2.Text = CboPlanoContas2.BoundText
    End If
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
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta2.SetFocus
        End If
    End If
End Sub

Private Sub MskPcr_GotFocus()
'    Call subSelecionaMSK(MskPcr)
End Sub

Private Sub MskPcr_LostFocus()
'    If MskPcr.Text <> "" Then
'        If CboPcr.Text = "" Then
'            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
'            If MskPcr.Visible = True Then MskPcr.SetFocus
'        End If
'    End If
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

Private Sub TDBGrid1_DblClick()
CmdAlterar_Click
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    Dim Cont As Integer
    Dim SomaTotal As Integer
    Dim TotFiltro As Integer
    Dim VALOR As String
    Dim DBGCol As Integer
    Dim DBGLin As Integer
    Dim vargBookmark As Variant
    
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaNotasFiscais
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
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.Fields("tido_tx_Descricao") = "NOTA PROMISSÓRIA" Then
            CmdExcluir.Enabled = False
'            CmdEstorno.Enabled = False
'            cmdCancela.Enabled = False
            CmdAlteraDistribuicao.Enabled = False
'            CmdDevGlosa.Enabled = False
'            CmdPPSequencia.Enabled = False
'            CmdProcPagto.Enabled = False
        Else
            CmdExcluir.Enabled = True
'            CmdEstorno.Enabled = True
'            cmdCancela.Enabled = True
            CmdAlteraDistribuicao.Enabled = True
'            CmdDevGlosa.Enabled = True
'            CmdPPSequencia.Enabled = True
'            CmdProcPagto.Enabled = True
        End If
    End If
End Sub

'
' Entradas posteriores
Private Sub Form_Activate()
    Set Formulario = TelaNotasFiscais
    
    If Not PrimeiraVez Then
         subTelaValoresGlobais "P"
    Else
        Adodc1.Recordset.Find "nofi_dt_entrada >=" & Format$(Now, "DD/MM/YY")
        If Adodc1.Recordset.EOF And Not FunTabelaVazia(Adodc1) Then
            Adodc1.Recordset.MoveLast
        End If
    End If
     
    PrimeiraVez = False
    
'    'Se o usuário fizer alguma operação em outra tela, por exemplo alterar um contas a pagar
'    'através da tela de Duplicata, a variável sera modificada para S sendo necessário verificar
'    'quando a tela for reativada
'    If TelaNotasFiscais.Caption = "Manutenção e Consulta de Contas a Receber" Then
'        EntSaida = "E"
'    Else
'        EntSaida = "S"
'    End If
    
    '
    ' Alterado por Garcia, para tratar somente Contas a Pagar (Demanda JMJ)
    EntSaida = "S"
        

    If XGB_PQUALIEST = True Then
        TDBGrid1.Columns("Estoque").Visible = True
    End If
    
    If Not XGB_PCAPI Or EntSaida = "S" Then 'Não possui Capi ou é Saída, não mostrar título
        TDBGrid1.Columns("Título QualiCapi").Visible = False
    End If
    
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
    
End Sub

Private Sub Form_Load()

    Dim XLD_DATAENTRADA     As Date
    Dim XLI_QTDDIASFILTRO   As Integer
    Dim XLB_SIGILOSO        As Boolean
    Dim XLT_SQL             As String
        
'    If FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "12") Then
'     XLB_SIGILOSO = True
'     Else
'     XLB_SIGILOSO = False
'    End If
 
    '
    ' Alterado por Garcia, para tratar somente Contas a Pagar (Demanda JMJ)
     EntSaida = "S"
     XLB_SIGILOSO = False
      
    Set Formulario = TelaNotasFiscais
    Call subAjustaTelaLoc
    subManutencaoJanelasAtivas "I", "TelaNotasFiscais"
    
    'Verifica se o número de dias está definido no arquivo QualiAdmFin.ini. Se não estiver definido no arquivo, procura na tabela de empresas (campo empr_nr_QtdDiasFiltro).
    If Val(Trim(XGT_NUMDIAS)) <> 0 Then
        XLI_QTDDIASFILTRO = Val(Trim(FunNulo(XGT_NUMDIAS)))
        XLD_DATAENTRADA = DateAdd("d", (-Val(XGT_NUMDIAS)), Format(Now, "dd/MM/yyyy"))
    Else
        XLI_QTDDIASFILTRO = XGI_QTDDIASFILTRO
        XLD_DATAENTRADA = DateAdd("d", (-XGI_QTDDIASFILTRO), Format(Now, "dd/MM/yyyy"))
    End If

'    If XLI_QTDDIASFILTRO > 0 Then
'
'        If Not XGB_INCLUICANCELADOS Then 'Indica que não deve-se incluir os lancamentos cancelados
'            subCarregaDadosNV Adodc1, "", "nofi_dt_entrada", _
'                                    "", "SELECT * FROM ConsFINNotasFiscais " & _
'                                    " WHERE nofi_tx_tipomov='" & EntSaida & "' " & _
'                                    " AND ((nofi_tx_status = 'Realizado' " & _
'                                    " AND nofi_dt_entrada >= " & FunNuloData(XLD_DATAENTRADA, NomeSgbd) & ") " & _
'                                    " OR nofi_tx_status = 'Compromissado' " & _
'                                    " OR nofi_tx_status = 'Previsto') " & _
'                                    " AND empr_cd_empresa= " & PCodEmpresa & "" & _
'                                    "" & IIf(Not XLB_SIGILOSO, " AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL)", "") & ""
'        Else
'            subCarregaDadosNV Adodc1, "", "nofi_dt_entrada DESC", _
'                                    "", "SELECT TOP " & XLI_QTDDIASFILTRO & " * FROM ConsFINNotasFiscais " & _
'                                    " WHERE nofi_tx_tipomov='" & EntSaida & "' " & _
'                                    " AND ((nofi_tx_status = 'Realizado' " & _
'                                    " AND nofi_dt_entrada >= " & FunNuloData(XLD_DATAENTRADA, NomeSgbd) & ") " & _
'                                    " OR nofi_tx_status = 'Compromissado' " & _
'                                    " OR nofi_tx_status = 'Previsto' " & _
'                                    " OR nofi_tx_status = 'Cancelado') " & _
'                                    " AND empr_cd_empresa= " & PCodEmpresa & "" & _
'                                    "" & IIf(Not XLB_SIGILOSO, " AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL)", "") & ""
'        End If
'
'    Else
'
'        If Not XGB_INCLUICANCELADOS Then 'Indica que não deve-se incluir os lancamentos cancelados
'            subCarregaDadosNV Adodc1, "", "nofi_dt_entrada", _
'                                    "", "SELECT * FROM ConsFINNotasFiscais " & _
'                                    " WHERE nofi_tx_tipomov='" & EntSaida & "' " & _
'                                    " AND (nofi_tx_status = 'Realizado' " & _
'                                    " OR nofi_tx_status = 'Compromissado' " & _
'                                    " OR nofi_tx_status = 'Previsto') " & _
'                                    " AND empr_cd_empresa= " & PCodEmpresa & "" & _
'                                    "" & IIf(Not XLB_SIGILOSO, " AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL)", "") & ""
'        Else
'            subCarregaDadosNV Adodc1, "", "nofi_dt_entrada", _
'                                    "", "SELECT * FROM ConsFINNotasFiscais " & _
'                                    " WHERE nofi_tx_tipomov='" & EntSaida & "' " & _
'                                    " AND (nofi_tx_status = 'Realizado' " & _
'                                    " OR nofi_tx_status = 'Compromissado' " & _
'                                    " OR nofi_tx_status = 'Previsto' " & _
'                                    " OR nofi_tx_status = 'Cancelado') " & _
'                                    " AND empr_cd_empresa= " & PCodEmpresa & "" & _
'                                    "" & IIf(Not XLB_SIGILOSO, " AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL)", "") & ""
'        End If
'
'    End If
    
    '
    ' Monta o SQL para ler as Notas Fiscais e montar a Grid
    XLT_SQL = "SELECT * " & _
              "  FROM ConsFINNotasFiscais A" & _
              "       JOIN" & _
              "       Corretores B ON A.focl_cd_FornCli = B.focl_cd_FornCli" & _
              " WHERE nofi_tx_tipomov='S' " & _
              "   AND nofi_tx_status = 'Compromissado' " & _
              "    OR nofi_tx_status = 'Previsto' " & _
              "   AND empr_cd_empresa= " & PCodEmpresa & _
              IIf(Not XLB_SIGILOSO, " AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL)", "") & ""
              
    '
    ' Carrega Adodc1 (dados para o Grid)
    subCarregaDadosNV Adodc1, "", "nofi_dt_entrada DESC", "", XLT_SQL
              
'    If EntSaida = "E" Then
'
'        TelaNotasFiscais.Caption = "Manutenção e Consulta de Contas a Receber"
'        'TDBGrid1.Columns(8).Caption = "Nº Proc. Rec."
'        TDBGrid1.Columns(9).Caption = "Nº Proc. Rec."
'        TDBGrid1.Columns(2).Caption = "Cliente"
'        CmdDevGlosa.Caption = "&Glosa"
'        'CmdSair.Top = CmdDevGlosa.Top
'        'CmdSair.Left = CmdDevGlosa.Left
'        CmdSair.Top = CmdPPSequencia.Top
'        CmdSair.Left = CmdPPSequencia.Left
'        CmdEstorno.Left = CmdDevGlosa.Left
'        CmdDevGlosa.Width = 1350
'        CmdDevGlosa.Top = CmdProcPagto.Top
'        CmdDevGlosa.Left = CmdProcPagto.Left
'        CmdProcPagto.Visible = False
'        CmdPPSequencia.Visible = False
'
'    Else
'
'        If (PEmpresa = "QUALIDADOS C. E SISTEMAS LTDA SALVADOR" Or PEmpresa = "QUALINFO S. DE INFORMATICA LTDA - EPP SALVADOR") Then
'            TDBGrid1.Columns(14).Caption = "Pendente?"
'        End If
'
'    End If
    
    subHabilitaBotoes
    filtra_especial = True
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
    
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
'       CmdDuplicata.Enabled = False
'       CmdRecibo.Enabled = False
'       CmdProcPagto.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       
    Else
    
        If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
        End If
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaNotasFiscais"
End Sub


Private Sub TDBGrid1_SelChange(Cancel As Integer)

'Rotina para somar um conjunto de campos selecionados de uma coluna

Dim XLD_TOTALSELECAO As Variant
    Dim i As Integer

    i = 0
    
    XLD_TOTALSELECAO = 0
       
    If TDBGrid1.SelRange Then
        If TDBGrid1.SelStartCol = TDBGrid1.SelEndCol And (TDBGrid1.SelEndCol = 7) Then
        
            For i = 0 To TDBGrid1.SelBookmarks.Count - 1 Step 1
            
                If (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i))) <> "" Then
                    XLD_TOTALSELECAO = XLD_TOTALSELECAO + (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i)))
                End If
            Next
            
            MsgBox "Total selecionado: R$ " & Format(CStr(XLD_TOTALSELECAO), "0.00"), vbInformation, "Seleção"
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

Private Sub TxtVldistribuicao_GotFocus()
    Call subSelecionaTXT(TxtVldistribuicao)

End Sub

Private Sub TxtVldistribuicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtVldistribuicao_LostFocus()
    If TxtVldistribuicao.Text <> "" Then
        'TxtPercDistribuicao.Text = Format(CStr(((CDbl(TxtVldistribuicao.Text) / CDbl(TxtValor.Text)) * 100)), "standard")
        TxtPercDistribuicao.Text = ((CDbl(TxtVldistribuicao.Text) / Adodc1.Recordset.Fields("nofi_vl_Valor")) * 100)
    End If
End Sub

Private Sub TxtPercDistribuicao_GotFocus()
    Call subSelecionaTXT(TxtPercDistribuicao)
End Sub

Private Sub TxtPercDistribuicao_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPercDistribuicao_LostFocus()
    If TxtPercDistribuicao.Text <> "" Then
        TxtVldistribuicao.Text = Format$(((Adodc1.Recordset.Fields("nofi_vl_Valor") * CDbl(TxtPercDistribuicao.Text)) / 100), "##,##0.00")
    End If
End Sub
